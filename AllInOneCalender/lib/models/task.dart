import 'package:flutter/material.dart';

class Task {
  Task({
    required this.id,
    required this.title,
    this.notes = '',
    required this.scheduledAt,
    this.category = '',
    this.iconCodePoint,
    this.iconFontFamily,
    this.iconFontPackage,
    this.isCompleted = false,
  });

  final String id;
  String title;
  String notes;
  DateTime scheduledAt;
  String category;
  int? iconCodePoint;
  String? iconFontFamily;
  String? iconFontPackage;
  bool isCompleted;

  IconData? get icon {
    if (iconCodePoint == null) return null;
    return IconData(
      iconCodePoint!,
      fontFamily: iconFontFamily ?? 'MaterialIcons',
      fontPackage: iconFontPackage,
    );
  }

  Task copyWith({
    String? title,
    String? notes,
    DateTime? scheduledAt,
    String? category,
    IconData? icon,
    int? iconCodePoint,
    String? iconFontFamily,
    String? iconFontPackage,
    bool? isCompleted,
  }) {
    return Task(
      id: id,
      title: title ?? this.title,
      notes: notes ?? this.notes,
      scheduledAt: scheduledAt ?? this.scheduledAt,
      category: category ?? this.category,
      iconCodePoint: iconCodePoint ?? icon?.codePoint ?? this.iconCodePoint,
      iconFontFamily:
          iconFontFamily ?? icon?.fontFamily ?? this.iconFontFamily,
      iconFontPackage:
          iconFontPackage ?? icon?.fontPackage ?? this.iconFontPackage,
      isCompleted: isCompleted ?? this.isCompleted,
    );
  }

  Map<String, dynamic> toJson() => {
        'id': id,
        'title': title,
        'notes': notes,
        'scheduledAt': scheduledAt.toIso8601String(),
        'category': category,
        'iconCodePoint': iconCodePoint,
        'iconFontFamily': iconFontFamily,
        'iconFontPackage': iconFontPackage,
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
      iconCodePoint: json['iconCodePoint'] as int?,
      iconFontFamily: json['iconFontFamily'] as String?,
      iconFontPackage: json['iconFontPackage'] as String?,
      isCompleted: json['isCompleted'] as bool? ?? false,
    );
  }
}

class TaskState extends ChangeNotifier {
  final List<Task> _tasks = [];
  final List<String> _categories = ['仕事', 'プライベート'];
  final Map<String, Color> _categoryColors = {
    '仕事': Colors.blueAccent.shade100,
    'プライベート': Colors.pinkAccent.shade100,
  };

  static final List<Color> _defaultPalette = [
    Colors.lightBlue.shade200,
    Colors.teal.shade200,
    Colors.pink.shade200,
    Colors.orange.shade200,
    Colors.green.shade200,
    Colors.purple.shade200,
  ];

  List<Task> get tasks => List.unmodifiable(_tasks);
  List<String> get categories => List.unmodifiable(_categories);
  Map<String, Color> get categoryColors => Map.unmodifiable(_categoryColors);

  void addTask(Task task) {
    _ensureCategoryExists(task.category);
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
    _ensureCategoryExists(updatedTask.category);
    final index = _tasks.indexWhere((task) => task.id == updatedTask.id);
    if (index == -1) return;
    _tasks[index] = updatedTask;
    notifyListeners();
  }

  void addCategory(String category) {
    if (category.trim().isEmpty) return;
    if (_categories.contains(category)) return;
    _categories.add(category);
    _categoryColors[category] =
        _defaultPalette[category.hashCode.abs() % _defaultPalette.length];
    notifyListeners();
  }

  void setCategoryColor(String category, Color color) {
    if (category.trim().isEmpty) return;
    _categoryColors[category] = color;
    if (!_categories.contains(category)) {
      _categories.add(category);
    }
    notifyListeners();
  }

  Color colorForCategory(String category) {
    if (_categoryColors.containsKey(category)) return _categoryColors[category]!;
    final fallback =
        _defaultPalette[category.hashCode.abs() % _defaultPalette.length];
    _categoryColors[category] = fallback;
    return fallback;
  }

  void _ensureCategoryExists(String category) {
    if (category.isEmpty) return;
    if (!_categories.contains(category)) {
      _categories.add(category);
    }
    _categoryColors.putIfAbsent(
      category,
      () => _defaultPalette[category.hashCode.abs() % _defaultPalette.length],
    );
  }

  void removeTask(String id) {
    _tasks.removeWhere((task) => task.id == id);
    notifyListeners();
  }
}
