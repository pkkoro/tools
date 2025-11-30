import 'dart:math' as math;

import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../models/task.dart';
import '../widgets/task_form_sheet.dart';

class TaskListScreen extends StatefulWidget {
  const TaskListScreen({super.key});

  @override
  State<TaskListScreen> createState() => _TaskListScreenState();
}

class _TaskListScreenState extends State<TaskListScreen> {
  late DateTime _selectedDate;
  late DateTime _currentMonthStart;

  @override
  void initState() {
    super.initState();
    _selectedDate = _startOfDay(DateTime.now());
    _currentMonthStart = _startOfMonth(_selectedDate);
  }

  @override
  Widget build(BuildContext context) {
    final tasks = context.watch<TaskState>().tasks;
    final tasksByDay = <DateTime, List<Task>>{};
    for (final task in tasks) {
      if (task.dueDate == null) continue;
      final dayKey = _startOfDay(task.dueDate!);
      tasksByDay.putIfAbsent(dayKey, () => []).add(task);
    }

    final monthDays = _monthDays;
    final tasksForDay = tasksByDay[_selectedDate] ?? [];

    return Scaffold(
      appBar: AppBar(
        title: const Text('タスク管理'),
        actions: [
          IconButton(
            icon: const Icon(Icons.add_task),
            onPressed: () => _openComposer(context),
          )
        ],
      ),
      body: SafeArea(
        child: Column(
          children: [
            Padding(
              padding: const EdgeInsets.fromLTRB(16, 16, 16, 8),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    'メイン画面のカレンダーから日付を選んでタスクを登録',
                    style: Theme.of(context).textTheme.titleMedium,
                  ),
                  const SizedBox(height: 12),
                  Card(
                    shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
                    elevation: 0,
                    color: Theme.of(context).colorScheme.surfaceVariant.withOpacity(0.5),
                    child: Padding(
                      padding: const EdgeInsets.all(12),
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Row(
                            mainAxisAlignment: MainAxisAlignment.spaceBetween,
                            children: [
                              Row(
                                children: [
                                  IconButton(
                                    onPressed: () => _changeMonth(-1),
                                    icon: const Icon(Icons.chevron_left),
                                  ),
                                  Text(
                                    '${_currentMonthStart.year}年 ${_currentMonthStart.month}月',
                                    style: Theme.of(context).textTheme.titleMedium,
                                  ),
                                  IconButton(
                                    onPressed: () => _changeMonth(1),
                                    icon: const Icon(Icons.chevron_right),
                                  ),
                                ],
                              ),
                              Chip(
                                label: Text('${tasksForDay.length}件'),
                              ),
                            ],
                          ),
                          const SizedBox(height: 8),
                          Row(
                            children: const [
                              Expanded(child: Center(child: Text('日'))),
                              Expanded(child: Center(child: Text('月'))),
                              Expanded(child: Center(child: Text('火'))),
                              Expanded(child: Center(child: Text('水'))),
                              Expanded(child: Center(child: Text('木'))),
                              Expanded(child: Center(child: Text('金'))),
                              Expanded(child: Center(child: Text('土'))),
                            ],
                          ),
                          const SizedBox(height: 4),
                          GridView.builder(
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
                                isSelected: day != null && _isSameDay(day, _selectedDate),
                                onTap: day != null
                                    ? () => setState(() {
                                          _selectedDate = _startOfDay(day);
                                          _currentMonthStart = _startOfMonth(day);
                                        })
                                    : null,
                                onAdd: day != null
                                    ? () => _openComposerForDay(context, _startOfDay(day))
                                    : null,
                                tasks: day != null ? (tasksByDay[_startOfDay(day)] ?? []) : const [],
                              );
                            },
                          ),
                          const SizedBox(height: 8),
                          Row(
                            mainAxisAlignment: MainAxisAlignment.spaceBetween,
                            children: [
                              Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  Text(
                                    '${_selectedDate.year}/${_selectedDate.month}/${_selectedDate.day} のタスク',
                                    style: Theme.of(context).textTheme.bodyLarge,
                                  ),
                                  const SizedBox(height: 4),
                                  Text(
                                    '日付のマスに最大5件まで登録済みタスクが表示されます',
                                    style: Theme.of(context).textTheme.bodySmall,
                                  ),
                                ],
                              ),
                              FilledButton.tonalIcon(
                                onPressed: () => _openComposer(context),
                                icon: const Icon(Icons.add),
                                label: const Text('この日に追加'),
                              ),
                            ],
                          ),
                        ],
                      ),
                    ),
                  ),
                ],
              ),
            ),
            const Divider(height: 1),
            Expanded(
              child: tasksForDay.isEmpty
                  ? _EmptyState(selectedDate: _selectedDate)
                  : ListView.separated(
                      padding: const EdgeInsets.all(16),
                      itemCount: tasksForDay.length,
                      separatorBuilder: (_, __) => const SizedBox(height: 12),
                      itemBuilder: (context, index) {
                        final task = tasksForDay[index];
                        return _TaskTile(task: task, onEdit: () => _openComposer(context, task: task));
                      },
                    ),
            ),
          ],
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
    _openComposer(context, initialDate: date);
  }

  void _openComposer(BuildContext context, {Task? task, DateTime? initialDate}) {
    showModalBottomSheet<void>(
      context: context,
      isScrollControlled: true,
      builder: (_) => Padding(
        padding: EdgeInsets.only(
          bottom: MediaQuery.of(context).viewInsets.bottom,
        ),
        child: TaskFormSheet(
          task: task,
          initialDate: initialDate ?? _selectedDate,
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

class _DayCell extends StatelessWidget {
  const _DayCell({
    required this.date,
    required this.tasks,
    this.isSelected = false,
    this.onTap,
    this.onAdd,
  });

  final DateTime? date;
  final List<Task> tasks;
  final bool isSelected;
  final VoidCallback? onTap;
  final VoidCallback? onAdd;

  @override
  Widget build(BuildContext context) {
    if (date == null) {
      return const AspectRatio(
        aspectRatio: 1,
        child: SizedBox.shrink(),
      );
    }

    final theme = Theme.of(context);
    final visibleTasks = tasks.take(5).toList();
    final remaining = tasks.length - visibleTasks.length;

    return GestureDetector(
      onTap: onTap,
      child: AspectRatio(
        aspectRatio: 1,
        child: AnimatedContainer(
          duration: const Duration(milliseconds: 150),
          margin: const EdgeInsets.all(4),
          padding: const EdgeInsets.all(6),
          decoration: BoxDecoration(
            color: isSelected
                ? theme.colorScheme.primaryContainer
                : theme.colorScheme.surfaceVariant,
            border: Border.all(
              color: isSelected ? theme.colorScheme.primary : Colors.transparent,
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
                    style: theme.textTheme.titleMedium?.copyWith(fontWeight: FontWeight.bold),
                  ),
                  const Spacer(),
                  if (tasks.isNotEmpty)
                    Container(
                      padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                      decoration: BoxDecoration(
                        color: theme.colorScheme.primary.withOpacity(0.15),
                        borderRadius: BorderRadius.circular(8),
                      ),
                      child: Text(
                        '${tasks.length}',
                        style: theme.textTheme.labelSmall,
                      ),
                    ),
                ],
              ),
              const SizedBox(height: 4),
              Expanded(
                child: Wrap(
                  spacing: 6,
                  runSpacing: 6,
                  children: [
                    ...visibleTasks.map(
                      (task) => _TaskPreviewTile(label: task.title),
                    ),
                    _AddPreviewTile(onTap: onAdd),
                  ],
                ),
              ),
              if (remaining > 0)
                Padding(
                  padding: const EdgeInsets.only(top: 2),
                  child: Text(
                    '…他${remaining}件',
                    style: theme.textTheme.labelSmall,
                  ),
                ),
            ],
          ),
        ),
      ),
    );
  }
}

class _TaskPreviewTile extends StatelessWidget {
  const _TaskPreviewTile({required this.label});

  final String label;

  @override
  Widget build(BuildContext context) {
    return ConstrainedBox(
      constraints: const BoxConstraints(minWidth: 48, minHeight: 28, maxWidth: 120),
      child: Container(
        padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 4),
        decoration: BoxDecoration(
          color: Theme.of(context).colorScheme.surface,
          borderRadius: BorderRadius.circular(8),
          border: Border.all(
            color: Theme.of(context).colorScheme.outlineVariant,
          ),
        ),
        child: Text(
          label,
          maxLines: 1,
          overflow: TextOverflow.ellipsis,
          style: Theme.of(context).textTheme.labelSmall,
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
      constraints: const BoxConstraints(minWidth: 48, minHeight: 28, maxWidth: 120),
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(8),
        child: Container(
          padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 4),
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
              size: 16,
              color: Theme.of(context).colorScheme.primary,
            ),
          ),
        ),
      ),
    );
  }
}

class _TaskTile extends StatelessWidget {
  const _TaskTile({required this.task, required this.onEdit});

  final Task task;
  final VoidCallback onEdit;

  @override
  Widget build(BuildContext context) {
    final taskState = context.read<TaskState>();

    return Dismissible(
      key: ValueKey(task.id),
      background: Container(
        decoration: BoxDecoration(
          color: Theme.of(context).colorScheme.errorContainer,
          borderRadius: BorderRadius.circular(12),
        ),
        padding: const EdgeInsets.symmetric(horizontal: 16),
        alignment: Alignment.centerLeft,
        child: const Icon(Icons.delete),
      ),
      secondaryBackground: Container(
        decoration: BoxDecoration(
          color: Theme.of(context).colorScheme.errorContainer,
          borderRadius: BorderRadius.circular(12),
        ),
        padding: const EdgeInsets.symmetric(horizontal: 16),
        alignment: Alignment.centerRight,
        child: const Icon(Icons.delete),
      ),
      onDismissed: (_) => taskState.removeTask(task.id),
      child: ListTile(
        tileColor: Theme.of(context).colorScheme.surfaceVariant,
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
        leading: Checkbox(
          value: task.isCompleted,
          onChanged: (_) => taskState.toggleComplete(task.id),
        ),
        title: Text(
          task.title,
          style: task.isCompleted
              ? const TextStyle(
                  decoration: TextDecoration.lineThrough,
                  color: Colors.grey,
                )
              : null,
        ),
        subtitle: () {
          final meta = <Widget>[];
          if (task.dueDate != null) {
            meta.add(
              Text(
                '期限: ${task.dueDate!.year}/${task.dueDate!.month}/${task.dueDate!.day}',
                style: Theme.of(context).textTheme.bodySmall,
              ),
            );
          }
          if (task.notes.isNotEmpty) {
            meta.add(
              Padding(
                padding: const EdgeInsets.only(top: 4),
                child: Text(
                  task.notes,
                  maxLines: 2,
                  overflow: TextOverflow.ellipsis,
                ),
              ),
            );
          }

          if (meta.isEmpty) return null;

          return Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            mainAxisSize: MainAxisSize.min,
            children: meta,
          );
        }(),
        trailing: IconButton(
          icon: const Icon(Icons.edit),
          onPressed: onEdit,
        ),
      ),
    );
  }
}

class _EmptyState extends StatelessWidget {
  const _EmptyState({required this.selectedDate});

  final DateTime selectedDate;

  @override
  Widget build(BuildContext context) {
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Icon(
            Icons.calendar_today,
            size: 48,
            color: Theme.of(context).colorScheme.primary,
          ),
          const SizedBox(height: 12),
          Text(
            '${selectedDate.month}月${selectedDate.day}日のタスクはありません',
            style: Theme.of(context).textTheme.bodyLarge,
          ),
          const SizedBox(height: 8),
          const Text('右下のボタンからタスクを追加してください'),
        ],
      ),
    );
  }
}
