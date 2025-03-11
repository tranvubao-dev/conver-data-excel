import 'package:flutter/material.dart';

class CommonFileList extends StatelessWidget {
  final String title;
  final List<Map<String, dynamic>> files;
  final bool isBuy;
  final Function(int, bool, String) removeFile;
  final String debtCode; // Mã nợ (No156, No153, No211,...)

  const CommonFileList({
    Key? key,
    required this.title,
    required this.files,
    required this.isBuy,
    required this.removeFile,
    required this.debtCode,
  }) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return Expanded(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Text(
            title,
            style: const TextStyle(
              fontSize: 16,
              fontWeight: FontWeight.bold,
              color: Colors.black,
            ),
          ),
          SizedBox(
            height: 300,
            child: ListView.builder(
              shrinkWrap: true,
              itemCount: files.length,
              itemBuilder: (context, index) {
                final file = files[index];
                return Card(
                  elevation: 10,
                  margin:
                      const EdgeInsets.symmetric(vertical: 4, horizontal: 8),
                  child: ListTile(
                    title: Text(
                      file['name']!,
                      style: const TextStyle(
                        fontSize: 11,
                        fontWeight: FontWeight.bold,
                        color: Colors.black,
                      ),
                    ),
                    trailing: IconButton(
                      icon: const Icon(Icons.delete, color: Colors.red),
                      onPressed: () => removeFile(index, isBuy, debtCode),
                    ),
                  ),
                );
              },
            ),
          ),
        ],
      ),
    );
  }
}
