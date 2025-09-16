import psutil
import time
import threading
import csv
import os
from datetime import datetime, timedelta
from pynput import mouse
import win32gui
import win32process

# =========================
# Configurable variables
# =========================
TARGET_PROCESS = "Resolve.exe" # prog name, Autodesk.exe, Premiere.exe, etc.
IDLE_TIMEOUT = 20  # seconds
CSV_FILE = "resolve_time_log.csv"
EXCLUSIONS_FILE = "project-exclusions.ini"

WEEKLY_WAGE = 1.0 # YOUR WEEKLY WAGE!!!!!!! Your salary. (Hopefully more than one dollar)
DAILY_WAGE = WEEKLY_WAGE / 5
MONTHLY_WAGE = WEEKLY_WAGE * (52 / 12)
YEARLY_WAGE = WEEKLY_WAGE * 52


def load_project_exclusions():
    exclusions = set()
    if os.path.isfile(EXCLUSIONS_FILE):
        with open(EXCLUSIONS_FILE, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    exclusions.add(line)
    return exclusions


class Timer:
    def __init__(self):
        self.total_time = 0.0
        self.running = False
        self.last_active = time.time()
        self._start_time = None
        self._lock = threading.Lock()

    def start(self):
        with self._lock:
            if not self.running:
                self.running = True
                self._start_time = time.time()
                print("Timer started.")

    def pause(self):
        with self._lock:
            if self.running:
                elapsed = time.time() - self._start_time
                self.total_time += elapsed
                self.running = False
                timestamp = datetime.now().strftime("%H:%M")
                print(f"Timer paused at {timestamp}. Session time: {elapsed:.2f} seconds")

                # Print pay statistics when pausing
                stats = calculate_statistics()
                if stats:
                    print(
                        f"${stats['day']:.2f}/hr today ({stats['day_hours']:.2f} hours worked so far)\n"
                        f"${stats['week']:.2f}/hr this week ({stats['week_hours']:.2f} hours worked so far)\n"
                        f"${stats['month']:.2f}/hr this month ({stats['month_hours']:.2f} hours worked so far)\n"
                        f"${stats['year']:.2f}/hr this year ({stats['year_hours']:.2f} hours worked so far)\n"
                        "----------------------------------------"
                    )
                return elapsed
        return 0

    def get_time(self):
        with self._lock:
            if self.running:
                return self.total_time + (time.time() - self._start_time)
            return self.total_time


def is_resolve_running():
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == TARGET_PROCESS:
            return proc.info['pid']
    return None


def get_active_window_title():
    hwnd = win32gui.GetForegroundWindow()
    if hwnd:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        try:
            proc = psutil.Process(pid)
            if proc.name() == TARGET_PROCESS:
                return win32gui.GetWindowText(hwnd)
        except psutil.NoSuchProcess:
            return None
    return None


def get_project_name(window_title):
    if " - " in window_title:
        return window_title.split(" - ")[-1].strip()
    return "Unknown"


def mouse_listener(timer):
    def on_move(x, y):
        timer.last_active = time.time()
    with mouse.Listener(on_move=on_move) as listener:
        listener.join()


def log_time_to_csv(project_name, duration):
    file_exists = os.path.isfile(CSV_FILE)
    with open(CSV_FILE, mode='a', newline='') as file:
        writer = csv.writer(file)
        if not file_exists:
            writer.writerow(["Date", "Project", "Duration (seconds)"])
        writer.writerow([datetime.now().isoformat(), project_name, round(duration, 2)])


def calculate_statistics():
    if not os.path.isfile(CSV_FILE):
        return None

    exclusions = load_project_exclusions()
    now = datetime.now()
    today_date = now.date()
    current_year, current_month = now.year, now.month

    # Week boundaries (Sunday → Saturday)
    weekday = now.weekday()  # Monday=0, Sunday=6
    days_since_sunday = (weekday + 1) % 7
    week_start = now - timedelta(days=days_since_sunday)
    week_start = week_start.replace(hour=0, minute=0, second=0, microsecond=0)
    week_end = week_start + timedelta(days=6, hours=23, minutes=59, seconds=59)

    total_day = total_week = total_month = total_year = 0.0

    with open(CSV_FILE, newline='') as file:
        reader = csv.DictReader(file)
        for row in reader:
            try:
                entry_time = datetime.fromisoformat(row["Date"])
                project = row["Project"].strip()
                duration = float(row["Duration (seconds)"])
            except Exception:
                continue

            # Skip excluded projects
            if project in exclusions:
                continue

            if entry_time.date() == today_date:
                total_day += duration

            if week_start <= entry_time <= week_end:
                total_week += duration

            if entry_time.year == current_year and entry_time.month == current_month:
                total_month += duration

            if entry_time.year == current_year:
                total_year += duration

    def safe_hourly(total_seconds, wage):
        hours = total_seconds / 3600
        return (wage / hours) if hours > 0 else 0.0

    return {
        "day": safe_hourly(total_day, DAILY_WAGE),
        "week": safe_hourly(total_week, WEEKLY_WAGE),
        "month": safe_hourly(total_month, MONTHLY_WAGE),
        "year": safe_hourly(total_year, YEARLY_WAGE),
        "day_hours": total_day / 3600,
        "week_hours": total_week / 3600,
        "month_hours": total_month / 3600,
        "year_hours": total_year / 3600,
    }


def main():
    print("Resolve Work Tracker Loaded and running | Waiting for Resolve...")

    timer = Timer()
    mouse_thread = threading.Thread(target=mouse_listener, args=(timer,), daemon=True)
    mouse_thread.start()

    current_project = ""

    try:
        while True:
            pid = is_resolve_running()
            window_title = get_active_window_title()
            if pid and window_title:
                current_project = get_project_name(window_title)
                if (time.time() - timer.last_active) < IDLE_TIMEOUT:
                    timer.start()
                else:
                    elapsed = timer.pause()
                    if elapsed > 0:
                        log_time_to_csv(current_project, elapsed)
            else:
                elapsed = timer.pause()
                if elapsed > 0 and current_project:
                    log_time_to_csv(current_project, elapsed)
            time.sleep(5)
    except KeyboardInterrupt:
        final_elapsed = timer.pause()
        if final_elapsed > 0 and current_project:
            log_time_to_csv(current_project, final_elapsed)
        print("\nExiting. Total time tracked:", timer.get_time(), "seconds")


if __name__ == "__main__":
    main()
