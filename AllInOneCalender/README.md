# AllInOneCalendar Task Prototype

This is a lightweight Flutter prototype for the AllInOneCalendar iOS app focusing on calendar-based task management. The app keeps tasks in-memory, lets you pick a date from a calendar, and showcases the UX for adding, editing, completing, and deleting tasks on that day.

## Features
- Calendar-first home screen with a square day grid on a bright yellow background that shows up to five tasks directly in each date cell with tiny, elided labels plus a built-in [+] tile for quick adds, laid out with extra-compact spacing so one full month fits on a single mobile screen.
- Create tasks with title, optional notes, category selection (choose an existing one or add a new category inline), and an auto-filled due date from the selected calendar day that also shows in the task list and inside the calendar cells.
- Edit or delete existing tasks by tapping a task (or using the edit icon) or swiping to delete.
- Toggle completion with a checkbox that visually strikes through completed tasks.
- Minimal Material 3 styling suitable for iOS with Cupertino icons available.
- The current prototype version (v0.1.0) is visible in the app bar for quick reference.
- Single-column layout on all viewports that keeps the calendar and daily task list stacked and leans on the centered pop-up composer for edits and adds, keeping focus on the calendar grid.

## Getting Started
1. Ensure Flutter (3.13+) is installed and an iOS simulator or device is available.
2. Fetch dependencies:
   ```sh
   flutter pub get
   ```
3. Run the app on iOS:
   ```sh
   flutter run
   ```

## Next Steps
- Persist tasks locally (e.g., `shared_preferences` or SQLite).
- Add schedule and weight management modules to expand toward the full AllInOneCalendar vision.
- Sync tasks with a backend for multi-device support.
