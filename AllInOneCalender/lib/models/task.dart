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
    this.recurrence = const Recurrence.none(),
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
    final effectiveStartDate = recurrence.startDate ?? scheduledAt;
    final effectiveEnd = recurrence.endDate;
    final cappedEnd = effectiveEnd != null && effectiveEnd.isBefore(normalizedEnd)
        ? DateTime(
            effectiveEnd.year, effectiveEnd.month, effectiveEnd.day, 23, 59, 59)
        : normalizedEnd;

    final List<DateTime> results = [];

    DateTime next = DateTime(
      scheduledAt.year,
      scheduledAt.month,
      scheduledAt.day,
      scheduledAt.hour,
      scheduledAt.minute,
    );

    if (next.isBefore(effectiveStartDate)) {
      next = DateTime(
        effectiveStartDate.year,
        effectiveStartDate.month,
        effectiveStartDate.day,
        scheduledAt.hour,
        scheduledAt.minute,
      );
    }

    if (next.isAfter(cappedEnd)) return results;

    switch (recurrence.type) {
      case RecurrenceType.none:
        if (!next.isBefore(normalizedStart) && !next.isAfter(cappedEnd)) {
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
        final anchor = DateTime(
          effectiveStartDate.year,
          effectiveStartDate.month,
          effectiveStartDate.day,
          scheduledAt.hour,
          scheduledAt.minute,
        );
        if (normalizedStart.isBefore(anchor)) {
          next = anchor;
        } else {
          final offset = _daysBetween(anchor, normalizedStart);
          final remainder = offset % interval;
          final daysToAdd = remainder == 0 ? 0 : interval - remainder;
          next = anchor.add(Duration(days: offset + daysToAdd));
        }
        break;
      case RecurrenceType.weekly:
        final targetWeekdays = recurrence.weekdays ?? [scheduledAt.weekday];
        final anchor = normalizedStart.isAfter(effectiveStartDate)
            ? normalizedStart
            : DateTime(effectiveStartDate.year, effectiveStartDate.month,
                effectiveStartDate.day);
        next = _nextWeekdayOnOrAfter(anchor, targetWeekdays);
        break;
    }

    while (!next.isAfter(cappedEnd)) {
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
          next = _nextWeekdayOnOrAfter(
              next.add(const Duration(days: 1)),
              recurrence.weekdays ?? [scheduledAt.weekday]);
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

  DateTime _nextWeekdayOnOrAfter(DateTime start, List<int> weekdays) {
    final normalized = weekdays.isEmpty
        ? [scheduledAt.weekday]
        : weekdays.map((d) => ((d - 1) % 7) + 1).toList();
    for (var i = 0; i < 7; i++) {
      final candidate = start.add(Duration(days: i));
      if (normalized.contains(candidate.weekday)) {
        return candidate;
      }
    }
    return start;
  }
}

enum RecurrenceType { none, daily, everyNDays, weekly }

class Recurrence {
  const Recurrence({
    required this.type,
    this.intervalDays,
    this.weekdays,
    this.startDate,
    this.endDate,
  });

  const Recurrence.none() : this(type: RecurrenceType.none);

  final RecurrenceType type;
  final int? intervalDays;
  final List<int>? weekdays;
  final DateTime? startDate;
  final DateTime? endDate;

  Map<String, dynamic> toJson() => {
        'type': type.name,
        'intervalDays': intervalDays,
        'weekdays': weekdays,
        'startDate': startDate?.toIso8601String(),
        'endDate': endDate?.toIso8601String(),
      };

  factory Recurrence.fromJson(Map<String, dynamic>? json) {
    if (json == null || json['type'] == null) return Recurrence.none();
    final typeString = json['type'] as String?;
    final type = RecurrenceType.values.firstWhere(
      (e) => e.name == typeString,
      orElse: () => RecurrenceType.none,
    );
    List<int>? parsedWeekdays;
    if (json['weekdays'] is List) {
      parsedWeekdays = (json['weekdays'] as List)
          .whereType<int>()
          .map((d) => ((d - 1) % 7) + 1)
          .toList();
    } else if (json['weekday'] is int) {
      final weekday = json['weekday'] as int;
      parsedWeekdays = [(((weekday) - 1) % 7) + 1];
    }
    final start = json['startDate'] as String?;
    final end = json['endDate'] as String?;
    return Recurrence(
      type: type,
      intervalDays: json['intervalDays'] as int?,
      weekdays: parsedWeekdays,
      startDate: start != null ? DateTime.tryParse(start) : null,
      endDate: end != null ? DateTime.tryParse(end) : null,
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
