import 'package:flutter/foundation.dart';

class Task {
  Task({
    required this.id,
    required this.title,
    this.notes = '',
    required this.scheduledAt,
    this.category = '',
    this.isCompleted = false,
  });

  final String id;
  String title;
  String notes;
  DateTime scheduledAt;
  String category;
  bool isCompleted;

  Task copyWith({
    String? title,
    String? notes,
    DateTime? scheduledAt,
    String? category,
    bool? isCompleted,
  }) {
    return Task(
      id: id,
      title: title ?? this.title,
      notes: notes ?? this.notes,
      scheduledAt: scheduledAt ?? this.scheduledAt,
      category: category ?? this.category,
      isCompleted: isCompleted ?? this.isCompleted,
    );
  }

  Map<String, dynamic> toJson() => {
        'id': id,
        'title': title,
        'notes': notes,
        'scheduledAt': scheduledAt.toIso8601String(),
        'category': category,
        'isCompleted': isCompleted,
      };

  factory Task.fromJson(Map<String, dynamic> json) {
    return Task(
      id: json['id'] as String,
      title: json['title'] as String,
      notes: json['notes'] as String? ?? '',
      scheduledAt: DateTime.tryParse(json['scheduledAt'] as String) ??
          DateTime.now(),
      category: json['category'] as String? ?? '',
      isCompleted: json['isCompleted'] as bool? ?? false,
    );
  }
}

class TaskState extends ChangeNotifier {
  final List<Task> _tasks = [];
  final List<String> _categories = ['仕事', 'プライベート'];

  List<Task> get tasks => List.unmodifiable(_tasks);
  List<String> get categories => List.unmodifiable(_categories);

  void addTask(Task task) {
    _tasks.insert(0, task);
    notifyListeners();
  }

  void toggleComplete(String id) {
    final index = _tasks.indexWhere((task) => task.id == id);
    if (index == -1) return;
    final current = _tasks[index];
    _tasks[index] = current.copyWith(isCompleted: !current.isCompleted);
    notifyListeners();
  }

  void updateTask(Task updatedTask) {
    final index = _tasks.indexWhere((task) => task.id == updatedTask.id);
    if (index == -1) return;
    _tasks[index] = updatedTask;
    notifyListeners();
  }

  void addCategory(String category) {
    if (category.trim().isEmpty) return;
    if (_categories.contains(category)) return;
    _categories.add(category);
    notifyListeners();
  }

  void removeTask(String id) {
    _tasks.removeWhere((task) => task.id == id);
    notifyListeners();
  }
}
