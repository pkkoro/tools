import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import 'models/task.dart';
import 'screens/task_list_screen.dart';

void main() {
  runApp(const AllInOneCalendarApp());
}

class AllInOneCalendarApp extends StatelessWidget {
  const AllInOneCalendarApp({super.key});

  @override
  Widget build(BuildContext context) {
    return ChangeNotifierProvider(
      create: (_) => TaskState(),
      child: MaterialApp(
        debugShowCheckedModeBanner: false,
        title: 'AllInOneCalendar Tasks',
        theme: ThemeData(
          colorScheme: ColorScheme.fromSeed(seedColor: Colors.indigo),
          useMaterial3: true,
        ),
        home: const TaskListScreen(),
      ),
    );
  }
}
