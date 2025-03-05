import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'package:flutter/services.dart';
import 'package:html_to_excel/logic/excel_utils.dart';
import 'logic/code_generator.dart';
import 'common/common_button.dart';

class FilterDataPage extends StatefulWidget {
  const FilterDataPage({super.key});

  @override
  _FilterDataPageState createState() => _FilterDataPageState();
}

class _FilterDataPageState extends State<FilterDataPage> {
  final TextEditingController _controller = TextEditingController();
  Uint8List? barcodePng;
  String? barcodeSvg;
  List<Map<String, String>> uploadedFiles =
      []; // Chỉ chứa 1 file// File đã tải lên
  List<List<String>> dataExcel = [];
  //////////////

  void _pickExcelFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'xls'],
    );

    if (result != null) {
      setState(() {
        uploadedFiles = [
          {
            'name': result.files.single.name,
            'bytes': base64Encode(result.files.single
                .bytes!), // Dùng base64Encode để chuyển đổi bytes thành chuỗi
          }
        ]; // Luôn chỉ chứa 1 file
      });

      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Đã chọn file: ${result.files.single.name}')),
      );
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Không có file nào được chọn')),
      );
    }
  }

  void _removeFile(int index) {
    setState(() {
      uploadedFiles.clear(); // Xóa file đã tải lên
    });
  }

  Future<void> pickAndProcessExcelFile() async {
    if (uploadedFiles.isEmpty) {
      print("Không có file nào được chọn!");
      return;
    }

    try {
      // Lấy bytes từ uploadedFiles
      String base64String = uploadedFiles.first['bytes']!;
      var bytes = base64Decode(base64String);

      // Đọc file Excel từ bytes
      var excel = Excel.decodeBytes(bytes);

      // Lấy sheet đầu tiên
      var sheetName = excel.tables.keys.first;
      var table = excel.tables[sheetName];

      if (table == null || table.rows.isEmpty) {
        print("Lỗi: Không tìm thấy dữ liệu trong file Excel!");
        return;
      }

      // Thêm tiêu đề (giữ nguyên tiêu đề gốc và thêm cột mới "Generated Code")
      List<String> header = table.rows.first
          .map((cell) => cell?.value?.toString() ?? "")
          .toList();
      header.insert(1, "Generated Code");
      dataExcel.add(header);

      // Xử lý từng hàng và tạo mã từ cột A (columnIndex = 0)
      for (var i = 1; i < table.rows.length; i++) {
        // Kiểm tra nếu dòng hiện tại không có dữ liệu
        if (table.rows[i].isEmpty) {
          print("Bỏ qua dòng $i vì không có dữ liệu.");
          continue;
        }

        List<String> rowData =
            table.rows[i].map((cell) => cell?.value?.toString() ?? "").toList();

        // Kiểm tra nếu không có cột A hoặc giá trị cột A bị rỗng
        if (rowData.isEmpty ||
            rowData.length < 1 ||
            rowData[0].trim().isEmpty) {
          print("Bỏ qua dòng $i vì cột A trống.");
          continue;
        }

        // Lấy giá trị từ cột A và tạo mã
        String itemCode = rowData[0];
        String encodedCode = CodeGenerator.generateItemCode(itemCode);

        // Đảm bảo rowData có đủ cột trước khi thêm cột mới
        while (rowData.length < header.length - 1) {
          rowData.add(""); // Thêm cột trống nếu thiếu
        }
        rowData.insert(1, encodedCode);
        dataExcel.add(rowData);
      }
      // Xuất file Excel
    } catch (e) {
      print("Lỗi xử lý file: $e");
    }
    ExcelUtils.createExcelFileCode(dataExcel, "processed_excel.xlsx");
    resetState();
  }

  void resetState() {
    setState(() {
      dataExcel = []; // Đặt lại thanh tiêu đề
      uploadedFiles = []; // Xóa danh sách file
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
        appBar: AppBar(
          title: const Text('Lọc dữ liệu mua bán hàng'),
          backgroundColor: Colors.transparent, // Làm trong suốt AppBar
          elevation: 0, // Xóa bóng AppBar
        ),
        body: Container(
          child: Padding(
            padding: const EdgeInsets.all(16.0),
            child: Column(
              children: [
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                          mainAxisAlignment: MainAxisAlignment.center,
                          children: [
                            // ElevatedButton.icon(
                            //   onPressed: _pickExcelFile,
                            //   icon: const Icon(Icons.upload_file),
                            //   label: const Text('Chưa hoàn thiện'),
                            //   style: ElevatedButton.styleFrom(
                            //     padding: const EdgeInsets.symmetric(
                            //         vertical: 16, horizontal: 32),
                            //     textStyle: const TextStyle(fontSize: 18),
                            //     elevation: 10, // Tạo độ cao cho nút
                            //     shape: RoundedRectangleBorder(
                            //       borderRadius: BorderRadius.circular(
                            //           10), // Đường viền tròn cho nút
                            //     ),
                            //     backgroundColor: const Color.fromARGB(
                            //         255, 219, 237, 252), // Màu nền nút
                            //     shadowColor: const Color.fromARGB(
                            //         255, 226, 235, 250), // Màu bóng đổ
                            //   ).copyWith(
                            //     elevation: WidgetStateProperty.all<double>(
                            //         12), // Tăng độ cao khi nhấn
                            //     shadowColor: WidgetStateProperty.all<Color>(
                            //         Colors.blue[800]!),
                            //   ),
                            // ),
                            // const SizedBox(width: 30.0),
                            // ElevatedButton(
                            //   onPressed: uploadedFiles.isNotEmpty
                            //       ? pickAndProcessExcelFile
                            //       : null, // Vô hiệu hóa nút nếu không có file
                            //   style: ElevatedButton.styleFrom(
                            //     padding: const EdgeInsets.symmetric(
                            //         vertical: 16, horizontal: 32),
                            //     textStyle: const TextStyle(fontSize: 18),
                            //     elevation: 13,
                            //     backgroundColor: uploadedFiles.isNotEmpty
                            //         ? const Color.fromARGB(255, 219, 237,
                            //             252) // Màu khi nút được kích hoạt
                            //         : Colors.grey, // Màu khi nút bị vô hiệu hóa
                            //   ),
                            //   child: const Text('Tạo lọc dữ liệu Excel'),
                            // ),
                            // CommonButton(
                            //   onPressed: _pickExcelFile,
                            //   label: 'Chưa hoàn thiện',
                            //   icon: Icons.upload_file,
                            // ),
                            // const SizedBox(width: 30.0),
                            // CommonButton(
                            //   onPressed: uploadedFiles.isNotEmpty
                            //       ? pickAndProcessExcelFile
                            //       : null,
                            //   label: 'Tạo lọc dữ liệu Excel',
                            //   isDisabled: uploadedFiles.isEmpty,
                            // ),
                          ]),
                      const SizedBox(height: 20),
                      Expanded(
                        child: uploadedFiles.isEmpty
                            ? const Center(
                                child: Text('Chưa có file nào được tải lên'))
                            : ListView.builder(
                                itemCount: uploadedFiles.length,
                                itemBuilder: (context, index) {
                                  final file = uploadedFiles[index];
                                  return Card(
                                    elevation: 10,
                                    margin: const EdgeInsets.symmetric(
                                        vertical: 8, horizontal: 16),
                                    child: ListTile(
                                      leading: const Icon(Icons.file_present,
                                          size: 40),
                                      title: Text(file['name'] ?? ''),
                                      trailing: IconButton(
                                        icon: const Icon(Icons.delete,
                                            color: Colors.red),
                                        onPressed: () => _removeFile(index),
                                      ),
                                    ),
                                  );
                                },
                              ),
                      ),
                    ],
                  ),
                ),
              ],
            ),
          ),
        ));
  }
}
