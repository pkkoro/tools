import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import 'package:uuid/uuid.dart';
import '../models/task.dart';

class TaskFormSheet extends StatefulWidget {
  const TaskFormSheet({super.key, this.task, this.initialDate});

  final Task? task;
  final DateTime? initialDate;

  @override
  State<TaskFormSheet> createState() => _TaskFormSheetState();
}

class _TaskFormSheetState extends State<TaskFormSheet> {
  final _formKey = GlobalKey<FormState>();
  late final TextEditingController _titleController;
  late final TextEditingController _notesController;
  late DateTime _scheduledAt;
  late TimeOfDay _timeOfDay;
  String _selectedCategory = '';
  bool _initializedCategory = false;
  IconData? _selectedIcon;

  @override
  void initState() {
    super.initState();
    _titleController = TextEditingController(text: widget.task?.title ?? '');
    _notesController = TextEditingController(text: widget.task?.notes ?? '');
    final initialDateTime = widget.task?.scheduledAt ??
        widget.initialDate ??
        DateTime.now();
    _selectedIcon = widget.task?.icon;
    _timeOfDay = widget.task != null
        ? TimeOfDay(
            hour: widget.task!.scheduledAt.hour,
            minute: widget.task!.scheduledAt.minute,
          )
        : _roundToNearestHalfHour(TimeOfDay.fromDateTime(initialDateTime));
    _scheduledAt = DateTime(
      initialDateTime.year,
      initialDateTime.month,
      initialDateTime.day,
      _timeOfDay.hour,
      _timeOfDay.minute,
    );
  }

  @override
  void didChangeDependencies() {
    super.didChangeDependencies();
    if (_initializedCategory) return;
    final categories = context.read<TaskState>().categories;
    final existing = widget.task?.category;
    if (existing != null && existing.isNotEmpty) {
      _selectedCategory = existing;
    } else if (categories.isNotEmpty) {
      _selectedCategory = categories.first;
    }
    _initializedCategory = true;
  }

  @override
  void dispose() {
    _titleController.dispose();
    _notesController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final isEditing = widget.task != null;
    final theme = Theme.of(context);
    final taskState = context.watch<TaskState>();
    final addCategoryLabel = '＋ カテゴリを追加';
    final categoryOptions = [
      ...taskState.categories,
      if (_selectedCategory.isNotEmpty && !taskState.categories.contains(_selectedCategory))
        _selectedCategory,
      addCategoryLabel,
    ];

    final timeOptions = List.generate(
      48,
      (index) {
        final hour = index ~/ 2;
        final minute = (index % 2) * 30;
        final display = '${hour.toString().padLeft(2, '0')}:${minute.toString().padLeft(2, '0')}';
        return DropdownMenuItem<TimeOfDay>(
          value: TimeOfDay(hour: hour, minute: minute),
          child: Text(display),
        );
      },
    );

    final iconOptions = <IconData>[
      Icons.work_outline,
      Icons.home_outlined,
      Icons.school,
      Icons.fitness_center,
      Icons.shopping_bag_outlined,
      Icons.emoji_events_outlined,
      Icons.coffee_outlined,
      Icons.flight_takeoff,
      Icons.medical_services_outlined,
      Icons.music_note,
      Icons.sports_esports,
      Icons.lightbulb_outline,
    ];

    return Padding(
      padding: const EdgeInsets.fromLTRB(20, 20, 20, 32),
      child: Form(
        key: _formKey,
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(
              isEditing ? 'タスクを編集' : '新しいタスク',
              style: theme.textTheme.titleLarge,
            ),
            const SizedBox(height: 16),
            TextFormField(
              controller: _titleController,
              decoration: const InputDecoration(
                labelText: 'タイトル',
                border: OutlineInputBorder(),
              ),
              validator: (value) {
                if (value == null || value.trim().isEmpty) {
                  return 'タイトルを入力してください';
                }
                return null;
              },
            ),
            const SizedBox(height: 12),
            TextFormField(
              controller: _notesController,
              minLines: 2,
              maxLines: 4,
              decoration: const InputDecoration(
                labelText: 'メモ',
                border: OutlineInputBorder(),
              ),
            ),
            const SizedBox(height: 12),
            DropdownButtonFormField<IconData?>(
              value: _selectedIcon,
              items: [
                const DropdownMenuItem<IconData?>(
                  value: null,
                  child: Text('アイコンなし'),
                ),
                ...iconOptions.map(
                  (icon) => DropdownMenuItem<IconData?>(
                    value: icon,
                    child: Row(
                      children: [
                        Icon(icon, size: 18),
                        const SizedBox(width: 6),
                        Text(_iconLabel(icon)),
                      ],
                    ),
                  ),
                ),
                if (_selectedIcon != null && !iconOptions.contains(_selectedIcon))
                  DropdownMenuItem<IconData?>(
                    value: _selectedIcon,
                    child: Row(
                      children: [
                        Icon(_selectedIcon, size: 18),
                        const SizedBox(width: 6),
                        const Text('カスタムアイコン'),
                      ],
                    ),
                  ),
              ],
              decoration: const InputDecoration(
                labelText: 'アイコン',
                border: OutlineInputBorder(),
              ),
              onChanged: (value) => setState(() => _selectedIcon = value),
            ),
            const SizedBox(height: 12),
            DropdownButtonFormField<String>(
              value: categoryOptions.contains(_selectedCategory)
                  ? _selectedCategory
                  : (taskState.categories.isNotEmpty ? taskState.categories.first : addCategoryLabel),
              items: [
                ...taskState.categories.map(
                  (c) => DropdownMenuItem(value: c, child: Text(c)),
                ),
                const DropdownMenuItem(
                  value: '＋ カテゴリを追加',
                  child: Text('＋ カテゴリを追加'),
                ),
              ],
              decoration: const InputDecoration(
                labelText: 'カテゴリ',
                border: OutlineInputBorder(),
              ),
              onChanged: (value) async {
                if (value == null) return;
                if (value == addCategoryLabel) {
                  final newCategory = await _promptAddCategory();
                  if (newCategory != null) {
                    taskState.addCategory(newCategory);
                    setState(() => _selectedCategory = newCategory);
                  }
                } else {
                  setState(() => _selectedCategory = value);
                }
              },
            ),
            const SizedBox(height: 12),
            Row(
              children: [
                Expanded(
                  child: Text('日付: ${_scheduledAt.year}/${_scheduledAt.month}/${_scheduledAt.day}'),
                ),
                TextButton.icon(
                  onPressed: _pickDate,
                  icon: const Icon(Icons.calendar_today),
                  label: const Text('日付を変更'),
                )
              ],
            ),
            const SizedBox(height: 12),
            DropdownButtonFormField<TimeOfDay>(
              value: _timeOfDay,
              items: timeOptions,
              decoration: const InputDecoration(
                labelText: '時刻 (30分刻み)',
                border: OutlineInputBorder(),
              ),
              onChanged: (value) {
                if (value == null) return;
                setState(() {
                  _timeOfDay = value;
                  _scheduledAt = DateTime(
                    _scheduledAt.year,
                    _scheduledAt.month,
                    _scheduledAt.day,
                    _timeOfDay.hour,
                    _timeOfDay.minute,
                  );
                });
              },
            ),
            const SizedBox(height: 20),
            Row(
              mainAxisAlignment: MainAxisAlignment.end,
              children: [
                TextButton(
                  onPressed: () => Navigator.pop(context),
                  child: const Text('キャンセル'),
                ),
                const SizedBox(width: 8),
                FilledButton(
                  onPressed: _submit,
                  child: Text(isEditing ? '更新' : '追加'),
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }

  Future<void> _pickDate() async {
    final now = DateTime.now();
    final initialDate = _scheduledAt;
    final newDate = await showDatePicker(
      context: context,
      initialDate: initialDate,
      firstDate: DateTime(now.year - 1),
      lastDate: DateTime(now.year + 3),
    );
    if (newDate == null) return;
    setState(() {
      _scheduledAt = DateTime(
        newDate.year,
        newDate.month,
        newDate.day,
        _timeOfDay.hour,
        _timeOfDay.minute,
      );
    });
  }

  void _submit() {
    if (!_formKey.currentState!.validate()) return;
    final taskState = context.read<TaskState>();
    final isEditing = widget.task != null;
    final id = widget.task?.id ?? const Uuid().v4();
    final selectedCategory = _selectedCategory.isNotEmpty
        ? _selectedCategory
        : (context.read<TaskState>().categories.isNotEmpty
            ? context.read<TaskState>().categories.first
            : '');

    final task = Task(
      id: id,
      title: _titleController.text.trim(),
      notes: _notesController.text.trim(),
      scheduledAt: _scheduledAt,
      category: selectedCategory,
      iconCodePoint: _selectedIcon?.codePoint,
      iconFontFamily: _selectedIcon?.fontFamily,
      iconFontPackage: _selectedIcon?.fontPackage,
      isCompleted: widget.task?.isCompleted ?? false,
    );

    if (isEditing) {
      taskState.updateTask(task);
    } else {
      taskState.addTask(task);
    }
    Navigator.pop(context);
  }

  Future<String?> _promptAddCategory() async {
    final controller = TextEditingController();
    final result = await showDialog<String>(
      context: context,
      builder: (context) {
        return AlertDialog(
          title: const Text('カテゴリを追加'),
          content: TextField(
            controller: controller,
            decoration: const InputDecoration(
              labelText: 'カテゴリ名',
              border: OutlineInputBorder(),
            ),
            autofocus: true,
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.pop(context),
              child: const Text('キャンセル'),
            ),
            FilledButton(
              onPressed: () => Navigator.pop(context, controller.text.trim()),
              child: const Text('追加'),
            ),
          ],
        );
      },
    );
    controller.dispose();
    if (result == null || result.isEmpty) return null;
    return result;
  }

  String _iconLabel(IconData icon) {
    if (icon == Icons.work_outline) return '仕事';
    if (icon == Icons.home_outlined) return 'プライベート';
    if (icon == Icons.school) return '学習';
    if (icon == Icons.fitness_center) return '運動';
    if (icon == Icons.shopping_bag_outlined) return '買い物';
    if (icon == Icons.emoji_events_outlined) return 'イベント';
    if (icon == Icons.coffee_outlined) return 'カフェ';
    if (icon == Icons.flight_takeoff) return '旅行';
    if (icon == Icons.medical_services_outlined) return '通院';
    if (icon == Icons.music_note) return '音楽';
    if (icon == Icons.sports_esports) return 'ゲーム';
    if (icon == Icons.lightbulb_outline) return 'アイデア';
    return 'アイコン';
  }

  TimeOfDay _roundToNearestHalfHour(TimeOfDay time) {
    final totalMinutes = time.hour * 60 + time.minute;
    final rounded = (totalMinutes / 30).round() * 30;
    final hour = (rounded ~/ 60) % 24;
    final minute = rounded % 60;
    return TimeOfDay(hour: hour, minute: minute);
  }
}
