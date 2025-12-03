# AllInOneCalendar Task Prototype

This is a lightweight Flutter prototype for the AllInOneCalendar iOS app focusing on calendar-based task management. The app keeps tasks in-memory, lets you pick a date from a calendar, and showcases the UX for adding and editing tasks pinned to specific days.

## Features
- Calendar-first home screen with a square day grid on a bright yellow background that shows up to five tasks directly in each date cell with tiny, elided labels plus a built-in [+] tile for quick adds, laid out with extra-compact spacing so one full month fits on a single mobile screen.
- Create tasks with title, optional notes, an optional icon, category selection (choose an existing one or add a new category inline), and a date＋time assignment from the selected calendar day using 30-minute time slots that surface inside the calendar cells.
- Configure repeat rules directly in the composer: none, every day, every N days, or weekly with a specific weekday.
- Toggle between title labels and icon chips (time-prefixed) for the per-day previews via the app bar shortcut.
- Tap any existing task label inside the calendar to open the pop-up editor for that task, or tap the inline [+] tile to compose a task for that day without leaving the grid.
- Manage category colors from a dedicated palette shortcut (alongside task add/toggle), with preset swatches, HEX input, and the ability to add categories directly from the palette screen.
- Minimal Material 3 styling suitable for iOS with Cupertino icons available.
- The app bar shows the latest pull request creation time instead of a static version tag.
- Single-column layout centered on the calendar grid and pop-up composer, without a separate task list pane beneath the calendar.
- The calendar grid and day cells keep a fixed aspect ratio so the monthly view and task previews remain visually consistent even when the screen size changes (mobile or Flutter web).

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
