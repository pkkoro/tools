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

  @override
  void initState() {
    super.initState();
    _selectedDate = _startOfDay(DateTime.now());
  }

  @override
  Widget build(BuildContext context) {
    final tasks = context.watch<TaskState>().tasks;
    final tasksForDay = tasks
        .where((task) => task.dueDate != null && _isSameDay(task.dueDate!, _selectedDate))
        .toList();

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
      body: Column(
        children: [
          Padding(
            padding: const EdgeInsets.fromLTRB(16, 16, 16, 8),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'カレンダーから日付を選択してタスクを登録',
                  style: Theme.of(context).textTheme.titleMedium,
                ),
                const SizedBox(height: 12),
                CalendarDatePicker(
                  initialDate: _selectedDate,
                  firstDate: DateTime(_selectedDate.year - 1),
                  lastDate: DateTime(_selectedDate.year + 2),
                  onDateChanged: (date) => setState(() {
                    _selectedDate = _startOfDay(date);
                  }),
                ),
                const SizedBox(height: 4),
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    Text(
                      '${_selectedDate.year}/${_selectedDate.month}/${_selectedDate.day}のタスク',
                      style: Theme.of(context).textTheme.bodyLarge,
                    ),
                    Chip(
                      label: Text('${tasksForDay.length}件'),
                    ),
                  ],
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
      floatingActionButton: FloatingActionButton.extended(
        onPressed: () => _openComposer(context),
        label: const Text('追加'),
        icon: const Icon(Icons.add),
      ),
    );
  }

  void _openComposer(BuildContext context, {Task? task}) {
    showModalBottomSheet<void>(
      context: context,
      isScrollControlled: true,
      builder: (_) => Padding(
        padding: EdgeInsets.only(
          bottom: MediaQuery.of(context).viewInsets.bottom,
        ),
        child: TaskFormSheet(
          task: task,
          initialDate: _selectedDate,
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
