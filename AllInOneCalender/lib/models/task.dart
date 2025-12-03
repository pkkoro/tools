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
    this.recurrence = Recurrence.none,
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
  Recurrence recurrence;
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
    Recurrence? recurrence,
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
      recurrence: recurrence ?? this.recurrence,
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
        'recurrence': recurrence.toJson(),
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
      recurrence: Recurrence.fromJson(
        json['recurrence'] as Map<String, dynamic>?,
      ),
      isCompleted: json['isCompleted'] as bool? ?? false,
    );
  }

  List<DateTime> occurrencesBetween(DateTime start, DateTime end) {
    final normalizedStart = DateTime(start.year, start.month, start.day);
    final normalizedEnd = DateTime(end.year, end.month, end.day, 23, 59, 59);
    final List<DateTime> results = [];

    DateTime next = scheduledAt;
    if (next.isAfter(normalizedEnd)) return results;

    switch (recurrence.type) {
      case RecurrenceType.none:
        if (!next.isBefore(normalizedStart) && !next.isAfter(normalizedEnd)) {
          results.add(next);
        }
        return results;
      case RecurrenceType.daily:
        next = next.isBefore(normalizedStart)
            ? _advanceToStart(normalizedStart, next, 1)
            : next;
        break;
      case RecurrenceType.everyNDays:
        final interval = recurrence.intervalDays ?? 1;
        if (normalizedStart.isBefore(scheduledAt)) {
          next = scheduledAt;
        } else {
          final offset = _daysBetween(scheduledAt, normalizedStart);
          final remainder = offset % interval;
          final daysToAdd = remainder == 0 ? 0 : interval - remainder;
          next = scheduledAt.add(Duration(days: offset + daysToAdd));
        }
        break;
      case RecurrenceType.weekly:
        final targetWeekday = recurrence.weekday ?? scheduledAt.weekday;
        final anchor = normalizedStart.isAfter(scheduledAt)
            ? normalizedStart
            : DateTime(scheduledAt.year, scheduledAt.month, scheduledAt.day);
        next = _nextWeekdayOnOrAfter(anchor, targetWeekday);
        if (next.isBefore(scheduledAt)) {
          next = next.add(const Duration(days: 7));
        }
        break;
    }

    while (!next.isAfter(normalizedEnd)) {
      results.add(
        DateTime(next.year, next.month, next.day, scheduledAt.hour,
            scheduledAt.minute),
      );

      switch (recurrence.type) {
        case RecurrenceType.none:
          return results;
        case RecurrenceType.daily:
          next = next.add(const Duration(days: 1));
          break;
        case RecurrenceType.everyNDays:
          next = next.add(Duration(days: recurrence.intervalDays ?? 1));
          break;
        case RecurrenceType.weekly:
          next = next.add(const Duration(days: 7));
          break;
      }
    }

    return results;
  }

  int _daysBetween(DateTime from, DateTime to) {
    final start = DateTime(from.year, from.month, from.day);
    final end = DateTime(to.year, to.month, to.day);
    return end.difference(start).inDays;
  }

  DateTime _advanceToStart(DateTime start, DateTime base, int intervalDays) {
    final diff = _daysBetween(base, start);
    final steps = (diff / intervalDays).ceil();
    return base.add(Duration(days: steps * intervalDays));
  }

  DateTime _nextWeekdayOnOrAfter(DateTime start, int weekday) {
    final delta = (weekday - start.weekday + 7) % 7;
    return start.add(Duration(days: delta));
  }
}

enum RecurrenceType { none, daily, everyNDays, weekly }

class Recurrence {
  const Recurrence({
    required this.type,
    this.intervalDays,
    this.weekday,
  });

  final RecurrenceType type;
  final int? intervalDays;
  final int? weekday;

  factory Recurrence.none() => const Recurrence(type: RecurrenceType.none);

  Map<String, dynamic> toJson() => {
        'type': type.name,
        'intervalDays': intervalDays,
        'weekday': weekday,
      };

  factory Recurrence.fromJson(Map<String, dynamic>? json) {
    if (json == null || json['type'] == null) return Recurrence.none();
    final typeString = json['type'] as String?;
    final type = RecurrenceType.values.firstWhere(
      (e) => e.name == typeString,
      orElse: () => RecurrenceType.none,
    );
    return Recurrence(
      type: type,
      intervalDays: json['intervalDays'] as int?,
      weekday: json['weekday'] as int?,
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
