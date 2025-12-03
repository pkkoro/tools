import 'dart:math' as math;

import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../models/task.dart';
import '../widgets/category_palette_sheet.dart';
import '../widgets/task_form_sheet.dart';

class TaskListScreen extends StatefulWidget {
  const TaskListScreen({super.key});

  @override
  State<TaskListScreen> createState() => _TaskListScreenState();
}

class _TaskListScreenState extends State<TaskListScreen> {
  static const String _latestPrCreatedAt = 'PR: 2025/12/01 14:39';

  late DateTime _selectedDate;
  late DateTime _currentMonthStart;
  bool _showIcons = false;
  bool _showTime = true;
  bool _showTitle = true;

  @override
  void initState() {
    super.initState();
    _selectedDate = _startOfDay(DateTime.now());
    _currentMonthStart = _startOfMonth(_selectedDate);
  }

  @override
  Widget build(BuildContext context) {
    final tasks = context.watch<TaskState>().tasks;
    final categoryColors = context.watch<TaskState>().categoryColors;
    final tasksByDay = <DateTime, List<Task>>{};
    final monthStart = _currentMonthStart;
    final monthEnd = DateTime(monthStart.year, monthStart.month + 1, 0);

    for (final task in tasks) {
      final occurrences = task.occurrencesBetween(monthStart, monthEnd);
      for (final occurrence in occurrences) {
        final dayKey = _startOfDay(occurrence);
        tasksByDay.putIfAbsent(dayKey, () => []).add(
              task.copyWith(scheduledAt: occurrence),
            );
      }
    }

    for (final entry in tasksByDay.entries) {
      entry.value.sort((a, b) => a.scheduledAt.compareTo(b.scheduledAt));
    }

    final monthDays = _monthDays;
    final theme = Theme.of(context);
    return Scaffold(
      appBar: AppBar(
        titleSpacing: 12,
        title: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text('タスク管理'),
            Text(
              _latestPrCreatedAt,
              style: theme.textTheme.labelSmall,
            ),
          ],
        ),
        actions: [
          IconButton(
            tooltip: _showIcons ? 'タイトル表示に切り替え' : 'アイコン表示に切り替え',
            icon: Icon(
              _showIcons ? Icons.title : Icons.tag_faces_outlined,
              color: _showIcons ? theme.colorScheme.primary : null,
            ),
            onPressed: () => setState(() => _showIcons = !_showIcons),
          ),
          IconButton(
            tooltip: _showTime ? '時間を隠す' : '時間を表示',
            icon: Icon(
              Icons.schedule,
              color: _showTime ? theme.colorScheme.primary : null,
            ),
            onPressed: () => setState(() => _showTime = !_showTime),
          ),
          IconButton(
            tooltip: _showTitle ? 'タイトルを隠す' : 'タイトルを表示',
            icon: Icon(
              Icons.text_fields,
              color: _showTitle ? theme.colorScheme.primary : null,
            ),
            onPressed: () => setState(() => _showTitle = !_showTitle),
          ),
          IconButton(
            tooltip: 'カテゴリカラー設定',
            icon: const Icon(Icons.palette_outlined),
            onPressed: () => _openCategoryPalette(context),
          ),
          IconButton(
            icon: const Icon(Icons.add_task),
            onPressed: () => _openComposer(context),
          ),
        ],
      ),
      body: SafeArea(
        child: Center(
          child: ConstrainedBox(
            constraints: const BoxConstraints(maxWidth: 1400),
            child: Padding(
              padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
              child: _CalendarCard(
                currentMonthStart: _currentMonthStart,
                selectedDate: _selectedDate,
                tasksByDay: tasksByDay,
                categoryColors: categoryColors,
                showIcons: _showIcons,
                showTime: _showTime,
                showTitle: _showTitle,
                monthDays: monthDays,
                onChangeMonth: _changeMonth,
                onSelectDay: (day) {
                  setState(() {
                    _selectedDate = _startOfDay(day);
                    _currentMonthStart = _startOfMonth(day);
                  });
                },
                onAddForDay: (day) => _openComposerForDay(context, _startOfDay(day)),
                onEditTask: _handleEditFromCalendar,
              ),
            ),
          ),
        ),
      ),
      floatingActionButton: FloatingActionButton.extended(
        onPressed: () => _openComposer(context),
        label: const Text('追加'),
        icon: const Icon(Icons.add),
      ),
    );
  }

  void _openComposerForDay(BuildContext context, DateTime date) {
    setState(() {
      _selectedDate = date;
      _currentMonthStart = _startOfMonth(date);
    });
    _openComposer(context, initialDate: _withDefaultTime(date));
  }

  void _handleEditFromCalendar(DateTime day, Task task) {
    final normalized = _startOfDay(day);
    setState(() {
      _selectedDate = normalized;
      _currentMonthStart = _startOfMonth(normalized);
    });
    _openComposer(context, task: task, initialDate: task.scheduledAt);
  }

  void _openCategoryPalette(BuildContext context) {
    showModalBottomSheet<void>(
      context: context,
      isScrollControlled: true,
      builder: (_) => const CategoryPaletteSheet(),
    );
  }

  void _openComposer(BuildContext context, {Task? task, DateTime? initialDate}) {
    showModalBottomSheet<void>(
      context: context,
      isScrollControlled: true,
      builder: (_) => LayoutBuilder(
        builder: (context, constraints) {
          final maxWidth = constraints.maxWidth > 600 ? 520.0 : constraints.maxWidth;
          return Center(
            child: ConstrainedBox(
              constraints: BoxConstraints(maxWidth: maxWidth),
              child: Padding(
                padding: EdgeInsets.only(
                  bottom: MediaQuery.of(context).viewInsets.bottom,
                ),
                child: TaskFormSheet(
                  task: task,
                  initialDate: initialDate ?? _withDefaultTime(_selectedDate),
                ),
              ),
            ),
          );
        },
      ),
    );
  }

  bool _isSameDay(DateTime a, DateTime b) {
    return a.year == b.year && a.month == b.month && a.day == b.day;
  }

  DateTime _startOfDay(DateTime date) {
    return DateTime(date.year, date.month, date.day);
  }

  DateTime _withDefaultTime(DateTime date) {
    final now = DateTime.now();
    final rawMinutes = (now.minute / 30).round() * 30;
    final minute = rawMinutes % 60;
    final extraHour = rawMinutes >= 60 ? 1 : 0;
    final hour = (now.hour + extraHour) % 24;
    return DateTime(date.year, date.month, date.day, hour, minute);
  }

  DateTime _startOfMonth(DateTime date) {
    return DateTime(date.year, date.month);
  }

  void _changeMonth(int delta) {
    final newMonth = DateTime(_currentMonthStart.year, _currentMonthStart.month + delta, 1);
    final lastDay = DateTime(newMonth.year, newMonth.month + 1, 0).day;
    final safeDay = math.min(_selectedDate.day, lastDay);
    setState(() {
      _currentMonthStart = newMonth;
      _selectedDate = _startOfDay(DateTime(newMonth.year, newMonth.month, safeDay));
    });
  }

  List<DateTime?> get _monthDays {
    final firstDay = _currentMonthStart;
    final daysInMonth = DateTime(firstDay.year, firstDay.month + 1, 0).day;
    final leadingEmptySlots = firstDay.weekday % 7;

    final days = <DateTime?>[];
    days.addAll(List<DateTime?>.filled(leadingEmptySlots, null));
    for (var i = 0; i < daysInMonth; i++) {
      days.add(DateTime(firstDay.year, firstDay.month, i + 1));
    }
    final remainder = days.length % 7;
    if (remainder != 0) {
      days.addAll(List<DateTime?>.filled(7 - remainder, null));
    }
    return days;
  }
}

class _CalendarCard extends StatelessWidget {
  const _CalendarCard({
    required this.currentMonthStart,
    required this.selectedDate,
    required this.tasksByDay,
    required this.categoryColors,
    required this.showIcons,
    required this.showTime,
    required this.showTitle,
    required this.monthDays,
    required this.onChangeMonth,
    required this.onSelectDay,
    required this.onAddForDay,
    required this.onEditTask,
  });

  final DateTime currentMonthStart;
  final DateTime selectedDate;
  final Map<DateTime, List<Task>> tasksByDay;
  final Map<String, Color> categoryColors;
  final bool showIcons;
  final bool showTime;
  final bool showTitle;
  final List<DateTime?> monthDays;
  final void Function(int delta) onChangeMonth;
  final void Function(DateTime day) onSelectDay;
  final void Function(DateTime day) onAddForDay;
  final void Function(DateTime day, Task task) onEditTask;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    return Card(
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14)),
      elevation: 0,
      color: Colors.yellow.shade100,
      child: Padding(
        padding: const EdgeInsets.all(8),
        child: LayoutBuilder(
          builder: (context, constraints) {
            final rows = (monthDays.length / 7).ceil();
            final gridAspectRatio = 7 / rows;

            return Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    Row(
                      children: [
                        IconButton(
                          onPressed: () => onChangeMonth(-1),
                          icon: const Icon(Icons.chevron_left),
                          visualDensity: VisualDensity.compact,
                        ),
                        Text(
                          '${currentMonthStart.year}年 ${currentMonthStart.month}月',
                          style: theme.textTheme.titleSmall?.copyWith(
                            fontWeight: FontWeight.w600,
                            fontSize: 13,
                          ),
                        ),
                        IconButton(
                          onPressed: () => onChangeMonth(1),
                          icon: const Icon(Icons.chevron_right),
                          visualDensity: VisualDensity.compact,
                        ),
                      ],
                    ),
                    Text(
                      '${tasksByDay[selectedDate]?.length ?? 0}件',
                      style: theme.textTheme.labelSmall,
                    ),
                  ],
                ),
                const SizedBox(height: 4),
                Row(
                  children: const [
                    Expanded(child: Center(child: Text('日', style: TextStyle(fontSize: 11)))),
                    Expanded(child: Center(child: Text('月', style: TextStyle(fontSize: 11)))),
                    Expanded(child: Center(child: Text('火', style: TextStyle(fontSize: 11)))),
                    Expanded(child: Center(child: Text('水', style: TextStyle(fontSize: 11)))),
                    Expanded(child: Center(child: Text('木', style: TextStyle(fontSize: 11)))),
                    Expanded(child: Center(child: Text('金', style: TextStyle(fontSize: 11)))),
                    Expanded(child: Center(child: Text('土', style: TextStyle(fontSize: 11)))),
                  ],
                ),
                const SizedBox(height: 2),
                AspectRatio(
                  aspectRatio: gridAspectRatio,
                  child: GridView.builder(
                    shrinkWrap: true,
                    physics: const NeverScrollableScrollPhysics(),
                    gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(
                      crossAxisCount: 7,
                      childAspectRatio: 1,
                    ),
                    itemCount: monthDays.length,
                    itemBuilder: (context, index) {
                      final day = monthDays[index];
                      return _DayCell(
                        date: day,
                        isSelected: day != null && _isSameDay(day, selectedDate),
                        onTap: day != null ? () => onSelectDay(day) : null,
                        onAdd: day != null ? () => onAddForDay(day) : null,
                        tasks:
                            day != null ? (tasksByDay[_startOfDay(day)] ?? []) : const [],
                        categoryColors: categoryColors,
                        showIcons: showIcons,
                        showTime: showTime,
                        showTitle: showTitle,
                        onEditTask: day != null
                            ? (task) {
                                onSelectDay(day);
                                onEditTask(day, task);
                              }
                            : null,
                      );
                    },
                  ),
                ),
              ],
            );
          },
        ),
      ),
    );
  }

  bool _isSameDay(DateTime a, DateTime b) {
    return a.year == b.year && a.month == b.month && a.day == b.day;
  }

  DateTime _startOfDay(DateTime date) {
    return DateTime(date.year, date.month, date.day);
  }
}

class _DayCell extends StatelessWidget {
  const _DayCell({
    required this.date,
    required this.tasks,
    required this.categoryColors,
    required this.showIcons,
    required this.showTime,
    required this.showTitle,
    this.isSelected = false,
    this.onTap,
    this.onAdd,
    this.onEditTask,
  });

  final DateTime? date;
  final List<Task> tasks;
  final Map<String, Color> categoryColors;
  final bool showIcons;
  final bool showTime;
  final bool showTitle;
  final bool isSelected;
  final VoidCallback? onTap;
  final VoidCallback? onAdd;
  final void Function(Task task)? onEditTask;

  @override
  Widget build(BuildContext context) {
    if (date == null) {
      return const AspectRatio(
        aspectRatio: 1,
        child: SizedBox.shrink(),
      );
    }

    final theme = Theme.of(context);
    const maxVisibleTasks = 5;
    final visibleTasks = tasks.take(maxVisibleTasks).toList();
    final remaining = tasks.length - visibleTasks.length;

    return GestureDetector(
      onTap: onTap,
      child: AnimatedContainer(
        duration: const Duration(milliseconds: 150),
        margin: const EdgeInsets.all(2.5),
        padding: const EdgeInsets.all(4),
        decoration: BoxDecoration(
          color: isSelected ? Colors.amber.shade200 : Colors.yellow.shade200,
          border: Border.all(
            color: isSelected ? theme.colorScheme.primary : Colors.amber.shade300,
            width: 1.5,
          ),
          borderRadius: BorderRadius.circular(12),
        ),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Text(
                  '${date!.day}',
                  style: theme.textTheme.titleSmall?.copyWith(
                    fontWeight: FontWeight.bold,
                    fontSize: 13,
                  ),
                ),
                const Spacer(),
                if (tasks.isNotEmpty)
                  Container(
                    padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 1.5),
                    decoration: BoxDecoration(
                      color: theme.colorScheme.primary.withOpacity(0.15),
                      borderRadius: BorderRadius.circular(8),
                    ),
                    child: Text(
                      '${tasks.length}',
                      style: theme.textTheme.labelSmall?.copyWith(fontSize: 10),
                    ),
                  ),
              ],
            ),
            const SizedBox(height: 2),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  ...visibleTasks.map(
                    (task) => Padding(
                      padding: const EdgeInsets.only(top: 3),
                      child: _TaskPreviewTile(
                        label: showTitle ? _previewLabel(task) : null,
                        timeLabel: showTime ? _timeLabel(task) : null,
                        icon: showIcons ? (task.icon ?? Icons.task_alt_outlined) : null,
                        color: _taskColor(task, theme),
                        onTap: onEditTask != null ? () => onEditTask!(task) : null,
                      ),
                    ),
                  ),
                  const Spacer(),
                  Align(
                    alignment: Alignment.centerRight,
                    child: _AddPreviewTile(onTap: onAdd),
                  ),
                ],
              ),
            ),
            if (remaining > 0)
              Padding(
                padding: const EdgeInsets.only(top: 1),
                child: Text(
                  '…他${remaining}件',
                  style: theme.textTheme.labelSmall?.copyWith(fontSize: 10),
                ),
              ),
          ],
        ),
      ),
    );
  }

  Color _taskColor(Task task, ThemeData theme) {
    final key = task.category.isNotEmpty ? task.category : task.id;
    final resolved =
        key.isNotEmpty ? categoryColors[key] ?? theme.primaryColorLight : null;
    return (resolved ?? Colors.lightBlue.shade200).withOpacity(0.9);
  }

  String _previewLabel(Task task) {
    return task.title;
  }

  String _timeLabel(Task task) {
    final hour = task.scheduledAt.hour.toString().padLeft(2, '0');
    final minute = task.scheduledAt.minute.toString().padLeft(2, '0');
    return '$hour:$minute';
  }
}

class _TaskPreviewTile extends StatelessWidget {
  const _TaskPreviewTile({
    required this.color,
    this.timeLabel,
    this.label,
    this.icon,
    this.onTap,
  });

  final String? label;
  final String? timeLabel;
  final IconData? icon;
  final Color color;
  final VoidCallback? onTap;

  @override
  Widget build(BuildContext context) {
    final showClock = timeLabel != null && timeLabel!.isNotEmpty;
    final showText = label != null && label!.isNotEmpty;
    final showIcon = icon != null;
    final hasContent = showClock || showText || showIcon;

    return ConstrainedBox(
      constraints: const BoxConstraints(minWidth: 36, minHeight: 18, maxWidth: 150),
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(7),
        child: Container(
          padding: const EdgeInsets.symmetric(horizontal: 4, vertical: 2),
          decoration: BoxDecoration(
            color: color,
            borderRadius: BorderRadius.circular(7),
            border: Border.all(
              color: color.withOpacity(0.8),
            ),
          ),
          child: Row(
            mainAxisSize: MainAxisSize.min,
            children: [
              if (showIcon) ...[
                Icon(
                  icon,
                  size: 14,
                  color: Colors.black.withOpacity(0.8),
                ),
                if (showClock || showText) const SizedBox(width: 3),
              ],
              if (showClock)
                Text(
                  timeLabel!,
                  style: Theme.of(context).textTheme.labelSmall?.copyWith(
                        fontSize: 10,
                        color: Colors.black.withOpacity(0.78),
                      ),
                ),
              if (showClock && showText) const SizedBox(width: 3),
              if (showText)
                Flexible(
                  child: Text(
                    label!,
                    maxLines: 1,
                    overflow: TextOverflow.ellipsis,
                    style: Theme.of(context).textTheme.labelSmall?.copyWith(
                          fontSize: 10,
                          color: Colors.black.withOpacity(0.85),
                          height: 1.05,
                        ),
                  ),
                ),
              if (!hasContent)
                Text(
                  '…',
                  style: Theme.of(context).textTheme.labelSmall?.copyWith(
                        fontSize: 10,
                        color: Colors.black.withOpacity(0.7),
                      ),
                ),
            ],
          ),
        ),
      ),
    );
  }
}

class _AddPreviewTile extends StatelessWidget {
  const _AddPreviewTile({this.onTap});

  final VoidCallback? onTap;

  @override
  Widget build(BuildContext context) {
    return ConstrainedBox(
      constraints: const BoxConstraints(minWidth: 32, minHeight: 18, maxWidth: 80),
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(8),
        child: Container(
          padding: const EdgeInsets.symmetric(horizontal: 5, vertical: 2),
          decoration: BoxDecoration(
            color: Theme.of(context).colorScheme.primary.withOpacity(0.08),
            borderRadius: BorderRadius.circular(8),
            border: Border.all(
              color: Theme.of(context).colorScheme.primary.withOpacity(0.4),
            ),
          ),
          child: Center(
            child: Icon(
              Icons.add,
              size: 13,
              color: Theme.of(context).colorScheme.primary,
            ),
          ),
        ),
      ),
    );
  }
}

