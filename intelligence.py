"""
MS DOR Order Intelligence Engine

Tracks availability patterns, predicts restock windows, and optimizes
ordering strategy to maximize fill rates against competing bots.

Data flow:
  bot_script.py → log_scan() → availability_log.csv
  intelligence.py → analyze patterns → predictions + priority queue
  web_gui.py → Intelligence tab → visualize + configure
"""

import csv
import json
import os
import re
import statistics
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

# ── Data Files ──
LOG_FILE = Path('availability_log.csv')
LOG_FIELDS = [
    'timestamp', 'item_number', 'available_qty', 'was_available',
    'check_duration_ms', 'order_attempted', 'order_success',
]
PREDICTIONS_FILE = Path('restock_predictions.json')
WIN_LOG_FILE = Path('win_log.csv')
WIN_FIELDS = ['timestamp', 'item_number', 'qty_ordered', 'qty_available', 'compete_score']


def _ensure_log_file():
    if not LOG_FILE.exists():
        with open(LOG_FILE, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=LOG_FIELDS)
            writer.writeheader()


def _ensure_win_log():
    if not WIN_LOG_FILE.exists():
        with open(WIN_LOG_FILE, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=WIN_FIELDS)
            writer.writeheader()


# ── Scan Logging ──

def log_scan(item_number, available_qty, was_available, check_duration_ms=0,
             order_attempted=False, order_success=False):
    """Log a single availability scan result. Called from bot_script.py after every check."""
    _ensure_log_file()
    with open(LOG_FILE, 'a', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        writer.writerow({
            'timestamp': datetime.now().isoformat(),
            'item_number': str(item_number),
            'available_qty': available_qty,
            'was_available': was_available,
            'check_duration_ms': check_duration_ms,
            'order_attempted': order_attempted,
            'order_success': order_success,
        })


def log_win(item_number, qty_ordered, qty_available):
    """Log a successful order capture."""
    _ensure_win_log()
    compete_score = round(qty_ordered / max(qty_available, 1), 2)
    with open(WIN_LOG_FILE, 'a', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=WIN_FIELDS)
        writer.writerow({
            'timestamp': datetime.now().isoformat(),
            'item_number': str(item_number),
            'qty_ordered': qty_ordered,
            'qty_available': qty_available,
            'compete_score': compete_score,
        })


# ── Pattern Analysis ──

def _read_scan_log(days_back=30):
    """Read scan log, optionally filtered to recent N days."""
    if not LOG_FILE.exists():
        return []
    cutoff = datetime.now() - timedelta(days=days_back)
    rows = []
    with open(LOG_FILE, 'r', newline='') as f:
        for row in csv.DictReader(f):
            try:
                ts = datetime.fromisoformat(row['timestamp'])
                if ts >= cutoff:
                    row['_ts'] = ts
                    row['available_qty'] = int(row.get('available_qty', 0) or 0)
                    row['was_available'] = row.get('was_available', '').lower() == 'true'
                    rows.append(row)
            except (ValueError, KeyError):
                continue
    return rows


def _read_win_log(days_back=30):
    """Read win log."""
    if not WIN_LOG_FILE.exists():
        return []
    cutoff = datetime.now() - timedelta(days=days_back)
    rows = []
    with open(WIN_LOG_FILE, 'r', newline='') as f:
        for row in csv.DictReader(f):
            try:
                ts = datetime.fromisoformat(row['timestamp'])
                if ts >= cutoff:
                    row['_ts'] = ts
                    rows.append(row)
            except (ValueError, KeyError):
                continue
    return rows


def analyze_restock_patterns(days_back=30):
    """Analyze when items transition from 0 → available (restock events).

    Returns dict keyed by item_number:
    {
        'restock_events': [{'timestamp', 'qty', 'day_of_week', 'hour'}],
        'avg_restock_qty': float,
        'common_days': [int],     # 0=Mon, 6=Sun
        'common_hours': [int],    # 0-23
        'depletion_rate_per_min': float,  # how fast qty drops after restock
        'next_predicted_restock': str,    # ISO timestamp
        'confidence': float,              # 0.0-1.0
    }
    """
    rows = _read_scan_log(days_back)
    if not rows:
        return {}

    # Group by item
    by_item = defaultdict(list)
    for row in rows:
        by_item[row['item_number']].append(row)

    patterns = {}
    for item_num, scans in by_item.items():
        scans.sort(key=lambda x: x['_ts'])

        # Detect restock events: 0→positive transitions
        restocks = []
        depletion_samples = []
        prev_qty = 0
        prev_ts = None
        peak_qty = 0
        peak_ts = None

        for scan in scans:
            qty = scan['available_qty']
            ts = scan['_ts']

            # Restock detected: was 0, now positive
            if prev_qty == 0 and qty > 0:
                restocks.append({
                    'timestamp': ts.isoformat(),
                    'qty': qty,
                    'day_of_week': ts.weekday(),
                    'hour': ts.hour,
                    'minute': ts.minute,
                })
                peak_qty = qty
                peak_ts = ts

            # Depletion: qty dropped from peak
            if peak_ts and qty == 0 and peak_qty > 0:
                elapsed_min = max((ts - peak_ts).total_seconds() / 60, 1)
                rate = peak_qty / elapsed_min
                depletion_samples.append(rate)
                peak_qty = 0
                peak_ts = None

            prev_qty = qty
            prev_ts = ts

        if not restocks:
            patterns[item_num] = {
                'restock_events': [],
                'avg_restock_qty': 0,
                'common_days': [],
                'common_hours': [],
                'depletion_rate_per_min': 0,
                'next_predicted_restock': '',
                'confidence': 0.0,
                'total_scans': len(scans),
                'availability_rate': sum(1 for s in scans if s['was_available']) / len(scans),
            }
            continue

        # Analyze restock timing patterns
        days = [r['day_of_week'] for r in restocks]
        hours = [r['hour'] for r in restocks]
        qtys = [r['qty'] for r in restocks]

        # Most common days and hours
        day_counts = defaultdict(int)
        for d in days:
            day_counts[d] += 1
        hour_counts = defaultdict(int)
        for h in hours:
            hour_counts[h] += 1

        common_days = sorted(day_counts, key=day_counts.get, reverse=True)[:3]
        common_hours = sorted(hour_counts, key=hour_counts.get, reverse=True)[:3]

        # Predict next restock
        next_predicted = _predict_next_restock(restocks, common_days, common_hours)

        # Confidence based on data volume and consistency
        confidence = min(1.0, len(restocks) / 10)  # More events = higher confidence
        if len(set(days)) <= 2:
            confidence = min(1.0, confidence + 0.2)  # Consistent days boost
        if len(set(hours)) <= 3:
            confidence = min(1.0, confidence + 0.1)  # Consistent hours boost

        avg_depletion = statistics.mean(depletion_samples) if depletion_samples else 0

        patterns[item_num] = {
            'restock_events': restocks[-10:],  # Keep last 10
            'avg_restock_qty': round(statistics.mean(qtys), 1),
            'common_days': common_days,
            'common_hours': common_hours,
            'depletion_rate_per_min': round(avg_depletion, 2),
            'next_predicted_restock': next_predicted,
            'confidence': round(confidence, 2),
            'total_scans': len(scans),
            'availability_rate': round(sum(1 for s in scans if s['was_available']) / len(scans), 3),
        }

    return patterns


def _predict_next_restock(restocks, common_days, common_hours):
    """Predict the next restock time based on observed patterns."""
    if not restocks or not common_days or not common_hours:
        return ''

    now = datetime.now()
    target_hour = common_hours[0]

    # Find next occurrence of a common day
    for offset in range(1, 8):
        candidate = now + timedelta(days=offset)
        if candidate.weekday() in common_days:
            predicted = candidate.replace(hour=target_hour, minute=0, second=0, microsecond=0)
            return predicted.isoformat()

    return ''


# ── Priority Scoring ──

def score_item_priority(item_number, sales_velocity=0, margin=0, restock_pattern=None):
    """Score an item 0-100 for bot processing priority.

    Higher score = process first.

    Factors:
      - Sales velocity (how fast it sells) — 30%
      - Competition level (how fast qty depletes) — 25%
      - Restock predictability (can we time it?) — 20%
      - Margin/revenue contribution — 15%
      - Scarcity (low avg restock qty) — 10%
    """
    score = 0
    pattern = restock_pattern or {}

    # Velocity score (0-30): faster selling = higher priority
    vel_score = min(30, sales_velocity * 5)
    score += vel_score

    # Competition score (0-25): faster depletion = more competition
    depletion = pattern.get('depletion_rate_per_min', 0)
    if depletion > 5:
        score += 25  # Depletes in minutes — very competitive
    elif depletion > 1:
        score += 15
    elif depletion > 0.1:
        score += 8

    # Predictability score (0-20): higher confidence = we can time it
    confidence = pattern.get('confidence', 0)
    score += confidence * 20

    # Margin score (0-15)
    margin_score = min(15, margin * 1.5)
    score += margin_score

    # Scarcity score (0-10): low restock qty = scarce
    avg_qty = pattern.get('avg_restock_qty', 100)
    if avg_qty <= 12:
        score += 10
    elif avg_qty <= 24:
        score += 6
    elif avg_qty <= 48:
        score += 3

    return round(min(100, score), 1)


def build_priority_queue(items, patterns=None):
    """Sort items by priority score, highest first.

    items: list of dicts with at least 'item_number', optionally 'velocity', 'margin'
    patterns: restock pattern dict from analyze_restock_patterns()
    """
    if patterns is None:
        patterns = {}

    scored = []
    for item in items:
        item_num = str(item.get('item_number', '')).strip()
        pattern = patterns.get(item_num, {})
        priority = score_item_priority(
            item_num,
            sales_velocity=float(item.get('velocity', 0) or 0),
            margin=float(item.get('margin', 0) or 0),
            restock_pattern=pattern,
        )
        scored.append({
            **item,
            'priority_score': priority,
            'predicted_restock': pattern.get('next_predicted_restock', ''),
            'competition_level': (
                'high' if pattern.get('depletion_rate_per_min', 0) > 5
                else 'medium' if pattern.get('depletion_rate_per_min', 0) > 1
                else 'low'
            ),
        })

    scored.sort(key=lambda x: x['priority_score'], reverse=True)
    return scored


# ── Adaptive Scheduling ──

def get_optimal_check_schedule(patterns):
    """Given restock patterns, determine when and how frequently to check each item.

    Returns list of schedule entries:
    {
        'item_number': str,
        'check_interval_sec': int,   # How often to check
        'hot_windows': [...],        # Time windows to check aggressively
        'cold_windows': [...],       # Time windows to check less
    }
    """
    schedule = []
    for item_num, pattern in patterns.items():
        confidence = pattern.get('confidence', 0)
        depletion = pattern.get('depletion_rate_per_min', 0)

        # Base interval: competitive items get checked more often
        if depletion > 5:
            base_interval = 30  # Check every 30 seconds
        elif depletion > 1:
            base_interval = 60
        elif confidence > 0.5:
            base_interval = 120
        else:
            base_interval = 300  # Check every 5 minutes

        # Hot windows: around predicted restock times
        hot_windows = []
        common_days = pattern.get('common_days', [])
        common_hours = pattern.get('common_hours', [])
        for day in common_days:
            for hour in common_hours:
                # 30 min window before and after
                hot_windows.append({
                    'day_of_week': day,
                    'start_hour': max(0, hour - 1),
                    'end_hour': min(23, hour + 1),
                    'interval_sec': max(15, base_interval // 3),  # 3x more frequent
                })

        schedule.append({
            'item_number': item_num,
            'check_interval_sec': base_interval,
            'hot_windows': hot_windows,
            'priority': score_item_priority(item_num, restock_pattern=pattern),
        })

    schedule.sort(key=lambda x: x['priority'], reverse=True)
    return schedule


# ── Competition Analysis ──

def analyze_competition(days_back=14):
    """Analyze competition patterns — how fast items deplete after restock."""
    rows = _read_scan_log(days_back)
    if not rows:
        return {'items': [], 'summary': {}}

    by_item = defaultdict(list)
    for row in rows:
        by_item[row['item_number']].append(row)

    item_stats = []
    for item_num, scans in by_item.items():
        scans.sort(key=lambda x: x['_ts'])
        total = len(scans)
        available = sum(1 for s in scans if s['was_available'])
        availability_rate = available / total if total > 0 else 0

        # Find depletion events
        depletion_times = []
        peak_qty = 0
        peak_ts = None
        for scan in scans:
            qty = scan['available_qty']
            ts = scan['_ts']
            if qty > peak_qty:
                peak_qty = qty
                peak_ts = ts
            if peak_ts and qty == 0 and peak_qty > 0:
                elapsed = (ts - peak_ts).total_seconds() / 60
                depletion_times.append(elapsed)
                peak_qty = 0
                peak_ts = None

        avg_depletion_min = statistics.mean(depletion_times) if depletion_times else 0

        if avg_depletion_min > 0 and avg_depletion_min < 5:
            competition = 'extreme'
        elif avg_depletion_min < 30:
            competition = 'high'
        elif avg_depletion_min < 120:
            competition = 'medium'
        else:
            competition = 'low'

        item_stats.append({
            'item_number': item_num,
            'total_scans': total,
            'availability_rate': round(availability_rate, 3),
            'avg_depletion_minutes': round(avg_depletion_min, 1),
            'competition_level': competition,
            'depletion_events': len(depletion_times),
        })

    item_stats.sort(key=lambda x: x['avg_depletion_minutes'])

    # Summary
    total_items = len(item_stats)
    extreme = sum(1 for i in item_stats if i['competition_level'] == 'extreme')
    high = sum(1 for i in item_stats if i['competition_level'] == 'high')

    return {
        'items': item_stats,
        'summary': {
            'total_tracked': total_items,
            'extreme_competition': extreme,
            'high_competition': high,
            'avg_availability_rate': round(
                statistics.mean(i['availability_rate'] for i in item_stats), 3
            ) if item_stats else 0,
        },
    }


# ── Win Rate Analysis ──

def analyze_win_rate(days_back=14):
    """Analyze order success rate."""
    scans = _read_scan_log(days_back)
    wins = _read_win_log(days_back)

    total_attempts = sum(1 for s in scans if s.get('order_attempted', '').lower() == 'true')
    total_success = sum(1 for s in scans if s.get('order_success', '').lower() == 'true')
    win_rate = total_success / max(total_attempts, 1)

    # By item
    by_item = defaultdict(lambda: {'attempts': 0, 'successes': 0})
    for s in scans:
        if s.get('order_attempted', '').lower() == 'true':
            item = s['item_number']
            by_item[item]['attempts'] += 1
            if s.get('order_success', '').lower() == 'true':
                by_item[item]['successes'] += 1

    item_rates = []
    for item_num, stats in by_item.items():
        rate = stats['successes'] / max(stats['attempts'], 1)
        item_rates.append({
            'item_number': item_num,
            'attempts': stats['attempts'],
            'successes': stats['successes'],
            'win_rate': round(rate, 3),
        })
    item_rates.sort(key=lambda x: x['win_rate'], reverse=True)

    return {
        'overall_win_rate': round(win_rate, 3),
        'total_attempts': total_attempts,
        'total_successes': total_success,
        'by_item': item_rates,
    }


# ── Dashboard Summary ──

def get_intelligence_summary(days_back=14):
    """Full intelligence dashboard data."""
    patterns = analyze_restock_patterns(days_back)
    competition = analyze_competition(days_back)
    win_rate = analyze_win_rate(days_back)

    # Items approaching restock window
    upcoming_restocks = []
    now = datetime.now()
    for item_num, pattern in patterns.items():
        pred = pattern.get('next_predicted_restock', '')
        if pred:
            try:
                pred_dt = datetime.fromisoformat(pred)
                hours_until = (pred_dt - now).total_seconds() / 3600
                if 0 < hours_until < 48:
                    upcoming_restocks.append({
                        'item_number': item_num,
                        'predicted_time': pred,
                        'hours_until': round(hours_until, 1),
                        'confidence': pattern.get('confidence', 0),
                        'avg_qty': pattern.get('avg_restock_qty', 0),
                    })
            except (ValueError, TypeError):
                pass

    upcoming_restocks.sort(key=lambda x: x['hours_until'])

    return {
        'patterns': {k: {kk: vv for kk, vv in v.items() if kk != 'restock_events'}
                     for k, v in patterns.items()},
        'competition': competition,
        'win_rate': win_rate,
        'upcoming_restocks': upcoming_restocks[:20],
        'total_items_tracked': len(patterns),
        'data_since': (datetime.now() - timedelta(days=days_back)).isoformat(),
    }
