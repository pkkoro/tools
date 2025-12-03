import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

import '../models/task.dart';

class CategoryPaletteSheet extends StatefulWidget {
  const CategoryPaletteSheet({super.key});

  @override
  State<CategoryPaletteSheet> createState() => _CategoryPaletteSheetState();
}

class _CategoryPaletteSheetState extends State<CategoryPaletteSheet> {
  static const _dialogPadding = EdgeInsets.symmetric(horizontal: 20, vertical: 16);
  final _newCategoryController = TextEditingController();
  Color _newCategoryColor = Colors.lightBlue.shade200;

  final List<Color> _palette = [
    Colors.lightBlue.shade200,
    Colors.blueAccent.shade100,
    Colors.teal.shade200,
    Colors.green.shade200,
    Colors.orange.shade200,
    Colors.pink.shade200,
    Colors.purple.shade200,
    Colors.amber.shade200,
    Colors.cyan.shade200,
  ];

  @override
  void dispose() {
    _newCategoryController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final state = context.watch<TaskState>();
    final categories = state.categories;
    return Padding(
      padding: EdgeInsets.only(
        left: 20,
        right: 20,
        bottom: MediaQuery.of(context).viewInsets.bottom + 24,
        top: 20,
      ),
      child: SafeArea(
        child: SingleChildScrollView(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            mainAxisSize: MainAxisSize.min,
            children: [
              Row(
                children: [
                  const Icon(Icons.palette_outlined),
                  const SizedBox(width: 8),
                  Text(
                    'カテゴリごとの色を設定',
                    style: Theme.of(context).textTheme.titleLarge,
                  ),
                ],
              ),
              const SizedBox(height: 12),
              Text(
                'カレンダーの色をカテゴリ単位で調整できます。HEXコードを入力するか、プリセットの色を選択してください。',
                style: Theme.of(context).textTheme.bodySmall,
              ),
              const SizedBox(height: 16),
              ListView.separated(
                shrinkWrap: true,
                physics: const NeverScrollableScrollPhysics(),
                itemCount: categories.length,
                separatorBuilder: (_, __) => const Divider(height: 20),
                itemBuilder: (context, index) {
                  final category = categories[index];
                  final color = state.categoryColors[category] ??
                      state.colorForCategory(category);
                  return _CategoryRow(
                    category: category,
                    color: color,
                    onChange: () async {
                      final picked = await _pickColor(context, color);
                      if (picked != null) {
                        state.setCategoryColor(category, picked);
                      }
                    },
                  );
                },
              ),
              const SizedBox(height: 20),
              const Divider(),
              const SizedBox(height: 12),
              Text(
                'カテゴリを追加して色を決める',
                style: Theme.of(context).textTheme.titleMedium,
              ),
              const SizedBox(height: 8),
              Row(
                children: [
                  Expanded(
                    child: TextField(
                      controller: _newCategoryController,
                      decoration: const InputDecoration(
                        labelText: 'カテゴリ名',
                        border: OutlineInputBorder(),
                      ),
                    ),
                  ),
                  const SizedBox(width: 12),
                  _ColorDot(color: _newCategoryColor, size: 28),
                  TextButton(
                    onPressed: () async {
                      final picked = await _pickColor(context, _newCategoryColor);
                      if (picked != null) {
                        setState(() => _newCategoryColor = picked);
                      }
                    },
                    child: const Text('色変更'),
                  ),
                ],
              ),
              const SizedBox(height: 12),
              Align(
                alignment: Alignment.centerRight,
                child: FilledButton.icon(
                  onPressed: () {
                    final name = _newCategoryController.text.trim();
                    if (name.isEmpty) return;
                    state.setCategoryColor(name, _newCategoryColor);
                    _newCategoryController.clear();
                  },
                  icon: const Icon(Icons.add),
                  label: const Text('カテゴリを追加'),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Future<Color?> _pickColor(BuildContext context, Color initial) async {
    Color selected = initial;
    final controller = TextEditingController(text: _toHex(initial));
    final result = await showDialog<Color>(
      context: context,
      builder: (context) {
        return StatefulBuilder(
          builder: (context, setStateDialog) {
            return AlertDialog(
              title: const Text('色を選択'),
              contentPadding: _dialogPadding,
              content: Column(
                mainAxisSize: MainAxisSize.min,
                children: [
                  Wrap(
                    spacing: 8,
                    runSpacing: 8,
                    children: [
                      ..._palette.map(
                        (color) => GestureDetector(
                          onTap: () => setStateDialog(() => selected = color),
                          child: _ColorDot(
                            color: color,
                            size: 32,
                            isSelected: selected.value == color.value,
                          ),
                        ),
                      ),
                    ],
                  ),
                  const SizedBox(height: 12),
                  TextField(
                    controller: controller,
                    decoration: const InputDecoration(
                      labelText: 'HEX (#RRGGBB)',
                      border: OutlineInputBorder(),
                    ),
                    onChanged: (value) {
                      final parsed = _fromHex(value);
                      if (parsed != null) {
                        setStateDialog(() => selected = parsed);
                      }
                    },
                  ),
                ],
              ),
              actionsPadding: _dialogPadding,
              actions: [
                TextButton(
                  onPressed: () => Navigator.pop(context),
                  child: const Text('キャンセル'),
                ),
                FilledButton(
                  onPressed: () => Navigator.pop(context, selected),
                  child: const Text('決定'),
                ),
              ],
            );
          },
        );
      },
    );
    controller.dispose();
    return result;
  }

  Color? _fromHex(String value) {
    final normalized = value.replaceFirst('#', '');
    if (normalized.length != 6) return null;
    final hex = int.tryParse(normalized, radix: 16);
    if (hex == null) return null;
    return Color(0xFF000000 | hex);
  }

  String _toHex(Color color) {
    return '#${color.value.toRadixString(16).padLeft(8, '0').substring(2)}';
  }
}

class _CategoryRow extends StatelessWidget {
  const _CategoryRow({
    required this.category,
    required this.color,
    required this.onChange,
  });

  final String category;
  final Color color;
  final VoidCallback onChange;

  @override
  Widget build(BuildContext context) {
    return Row(
      children: [
        _ColorDot(color: color),
        const SizedBox(width: 12),
        Expanded(
          child: Text(
            category,
            style: Theme.of(context).textTheme.titleMedium,
          ),
        ),
        TextButton.icon(
          onPressed: onChange,
          icon: const Icon(Icons.palette_outlined),
          label: const Text('色を変更'),
        ),
      ],
    );
  }
}

class _ColorDot extends StatelessWidget {
  const _ColorDot({
    required this.color,
    this.size = 22,
    this.isSelected = false,
  });

  final Color color;
  final double size;
  final bool isSelected;

  @override
  Widget build(BuildContext context) {
    return Container(
      width: size,
      height: size,
      decoration: BoxDecoration(
        color: color,
        shape: BoxShape.circle,
        border: Border.all(
          color: isSelected ? Theme.of(context).colorScheme.primary : Colors.black12,
          width: isSelected ? 2 : 1,
        ),
        boxShadow: [
          if (isSelected)
            BoxShadow(
              color: Theme.of(context).colorScheme.primary.withOpacity(0.2),
              blurRadius: 6,
            )
        ],
      ),
    );
  }
}
