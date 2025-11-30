# AllInOneCalendar Task Prototype

This is a lightweight Flutter prototype for the AllInOneCalendar iOS app focusing on calendar-based task management. The app keeps tasks in-memory, lets you pick a date from a calendar, and showcases the UX for adding, editing, completing, and deleting tasks on that day.

## Features
- Calendar-first home screen with a square day grid that shows up to five tasks directly in each date cell.
- Create tasks with title, optional notes, and an auto-filled due date from the selected calendar day that also shows in the task list and inside the calendar cells.
- Edit or delete existing tasks via swipe-to-delete or edit action.
- Toggle completion with a checkbox that visually strikes through completed tasks.
- Minimal Material 3 styling suitable for iOS with Cupertino icons available.

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
