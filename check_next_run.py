#!/usr/bin/env python3
from datetime import datetime, timedelta
import json

# Load state and config
with open('automation_state.json') as f:
    state = json.load(f)

with open('automation_config.json') as f:
    config = json.load(f)

# Parse last check time
last_check_str = state['last_checked']
last_check = datetime.fromisoformat(last_check_str.replace('Z', '+00:00'))

# Get interval
interval_minutes = config.get('poll_interval_minutes', 10)

# Calculate next run
next_run = last_check + timedelta(minutes=interval_minutes)

# Get current time
now = datetime.now(last_check.tzinfo)

# Calculate time until next run
time_until = next_run - now
minutes_until = int(time_until.total_seconds() / 60)
seconds_until = int(time_until.total_seconds() % 60)

print(f"הסבב האחרון: {last_check.strftime('%d/%m/%Y %H:%M:%S')}")
print(f"מרווח זמן: {interval_minutes} דקות")
print(f"הסבב הבא: {next_run.strftime('%d/%m/%Y %H:%M:%S')}")
print(f"זמן עד הסבב הבא: {minutes_until} דקות ו-{seconds_until} שניות")
