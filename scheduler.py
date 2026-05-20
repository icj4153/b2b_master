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
        raise ValueError("Run time must use HH:MM format, for example 09:00") from exc

    if not 0 <= hour <= 23 or not 0 <= minute <= 59:
        raise ValueError("Run time must be a valid 24-hour time")
    return hour, minute


def parse_run_times(value):
    run_times = []
    for item in value.split(","):
        item = item.strip()
        if not item:
            continue
        run_times.append(parse_run_time(item))

    if not run_times:
        raise ValueError("B2B_RUN_TIMES must include at least one HH:MM time")

    return sorted(set(run_times))


def format_run_times(run_times):
    return ", ".join(f"{hour:02d}:{minute:02d}" for hour, minute in run_times)


def next_run_at(run_times):
    now = datetime.now()
    candidates = [
        now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        for hour, minute in run_times
    ]
    future_candidates = [candidate for candidate in candidates if candidate > now]
    if future_candidates:
        return min(future_candidates)

    first_hour, first_minute = run_times[0]
    return (now + timedelta(days=1)).replace(
        hour=first_hour,
        minute=first_minute,
        second=0,
        microsecond=0,
    )


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
    run_times_value = os.getenv("B2B_RUN_TIMES") or os.getenv("B2B_RUN_AT", "09:00")
    run_times = parse_run_times(run_times_value)
    run_on_start = os.getenv("B2B_RUN_ON_START", "false").lower() in {"1", "true", "yes"}

    print(f"[scheduler] Daily crawler scheduled at: {format_run_times(run_times)}", flush=True)
    if run_on_start:
        run_crawler()

    while True:
        run_at = next_run_at(run_times)
        sleep_seconds = max(1, int((run_at - datetime.now()).total_seconds()))
        print(f"[scheduler] Next run: {run_at:%Y-%m-%d %H:%M:%S}", flush=True)
        time.sleep(sleep_seconds)
        run_crawler()


if __name__ == "__main__":
    main()
