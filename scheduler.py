import os
import subprocess
import sys
import time
from datetime import datetime, timedelta


def parse_run_time(value):
    try:
        hour_text, minute_text = value.split(":", 1)
        hour = int(hour_text)
        minute = int(minute_text)
    except ValueError as exc:
        raise ValueError("B2B_RUN_AT must use HH:MM format, for example 09:00") from exc

    if not 0 <= hour <= 23 or not 0 <= minute <= 59:
        raise ValueError("B2B_RUN_AT must be a valid 24-hour time")
    return hour, minute


def next_run_at(hour, minute):
    now = datetime.now()
    run_at = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if run_at <= now:
        run_at += timedelta(days=1)
    return run_at


def run_crawler():
    started_at = datetime.now()
    print(f"[scheduler] Starting crawler at {started_at:%Y-%m-%d %H:%M:%S}", flush=True)

    result = subprocess.run(
        [sys.executable, "b2b_excel.py"],
        cwd=os.path.dirname(os.path.abspath(__file__)),
        text=True,
    )

    finished_at = datetime.now()
    print(
        f"[scheduler] Crawler finished at {finished_at:%Y-%m-%d %H:%M:%S} "
        f"with exit code {result.returncode}",
        flush=True,
    )
    return result.returncode


def main():
    hour, minute = parse_run_time(os.getenv("B2B_RUN_AT", "09:00"))
    run_on_start = os.getenv("B2B_RUN_ON_START", "false").lower() in {"1", "true", "yes"}

    print(f"[scheduler] Daily crawler scheduled at {hour:02d}:{minute:02d}", flush=True)
    if run_on_start:
        run_crawler()

    while True:
        run_at = next_run_at(hour, minute)
        sleep_seconds = max(1, int((run_at - datetime.now()).total_seconds()))
        print(f"[scheduler] Next run: {run_at:%Y-%m-%d %H:%M:%S}", flush=True)
        time.sleep(sleep_seconds)
        run_crawler()


if __name__ == "__main__":
    main()
