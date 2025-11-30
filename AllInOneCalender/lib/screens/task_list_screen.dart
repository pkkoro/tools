import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../models/task.dart';
import '../widgets/task_form_sheet.dart';

class TaskListScreen extends StatelessWidget {
  const TaskListScreen({super.key});

  @override
  Widget build(BuildContext context) {
    final tasks = context.watch<TaskState>().tasks;

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
      body: tasks.isEmpty
          ? const _EmptyState()
          : ListView.separated(
              padding: const EdgeInsets.all(16),
              itemCount: tasks.length,
              separatorBuilder: (_, __) => const SizedBox(height: 12),
              itemBuilder: (context, index) {
                final task = tasks[index];
                return _TaskTile(task: task);
              },
            ),
      floatingActionButton: FloatingActionButton.extended(
        onPressed: () => _openComposer(context),
        label: const Text('追加'),
        icon: const Icon(Icons.add),
      ),
    );
  }

  void _openComposer(BuildContext context) {
    showModalBottomSheet<void>(
      context: context,
      isScrollControlled: true,
      builder: (_) => Padding(
        padding: EdgeInsets.only(
          bottom: MediaQuery.of(context).viewInsets.bottom,
        ),
        child: const TaskFormSheet(),
      ),
    );
  }
}

class _TaskTile extends StatelessWidget {
  const _TaskTile({required this.task});

  final Task task;

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
        subtitle: task.notes.isNotEmpty
            ? Text(
                task.notes,
                maxLines: 2,
                overflow: TextOverflow.ellipsis,
              )
            : null,
        trailing: IconButton(
          icon: const Icon(Icons.edit),
          onPressed: () => _openEditSheet(context, task),
        ),
      ),
    );
  }

  void _openEditSheet(BuildContext context, Task task) {
    showModalBottomSheet<void>(
      context: context,
      isScrollControlled: true,
      builder: (_) => Padding(
        padding: EdgeInsets.only(
          bottom: MediaQuery.of(context).viewInsets.bottom,
        ),
        child: TaskFormSheet(task: task),
      ),
    );
  }
}

class _EmptyState extends StatelessWidget {
  const _EmptyState();

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
          const Text('まだタスクがありません'),
          const SizedBox(height: 8),
          const Text('右下のボタンからタスクを追加してください'),
        ],
      ),
    );
  }
}
