import 'dart:convert';

import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:html_to_excel/logic/excel_utils.dart';
import 'logic/code_generator.dart';
import 'package:barcode/barcode.dart';
import 'dart:ui' as ui;
import 'package:flutter_svg/flutter_svg.dart';
import 'package:file_picker/file_picker.dart';

class CreateCodePage extends StatefulWidget {
  const CreateCodePage({super.key});

  @override
  _CreateCodePageState createState() => _CreateCodePageState();
}

class _CreateCodePageState extends State<CreateCodePage> {
  final TextEditingController _controller = TextEditingController();
  String _generatedCode = '';
  Uint8List? barcodePng;
  String? barcodeSvg;
  List<Map<String, String>> uploadedFiles =
      []; // Chỉ chứa 1 file// File đã tải lên
  List<List<String>> dataExcel = [];
  //////////////
  void _generateCode() {
    setState(() {
      if (_controller.text.isNotEmpty) {
        _generatedCode = CodeGenerator.generateItemCode(_controller.text);
        generateBarcode(_generatedCode);
      }
    });
  }

  Future<void> generateBarcode(String data) async {
    // Tạo barcode
    final barcode = Barcode.code128(); // Hoặc Barcode.qrCode()

    // Chuyển barcode thành SVG
    barcodeSvg = barcode.toSvg(
      data,
      width: 300, // Chiều rộng
      height: 100, // Chiều cao
    );
    // Chuyển SVG thành PNG
    final svgString = barcode.toSvg(data, width: 300, height: 100);
    final pngBytes = await convertSvgToPng(svgString);

    setState(() {
      barcodePng = pngBytes;
    });
  }

  Future<Uint8List> convertSvgToPng(String svgString,
      {double width = 300, double height = 100}) async {
    try {
      // Phân tích SVG
      final svgDrawableRoot = await svg.fromSvgString(svgString, 'barcode');

      // Tạo hình ảnh từ SVG
      final picture = svgDrawableRoot.toPicture(size: Size(width, height));
      final image = await picture.toImage(width.toInt(), height.toInt());

      // Chuyển hình ảnh thành định dạng PNG
      final byteData = await image.toByteData(format: ui.ImageByteFormat.png);
      return byteData!.buffer.asUint8List();
    } catch (e) {
      throw Exception("Lỗi khi chuyển đổi SVG sang PNG: $e");
    }
  }

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
          title: const Text('Tạo Mã CODE'),
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
                      TextField(
                        controller: _controller,
                        decoration: InputDecoration(
                          labelText: 'Nhập tên mặt hàng',
                          border: const OutlineInputBorder(),
                          suffixIcon: IconButton(
                            icon: const Icon(Icons.clear),
                            onPressed: () {
                              setState(() {
                                _controller
                                    .clear(); // Xóa nội dung của TextField
                                _generatedCode = ''; // Xóa mã mặt hàng
                              });
                            },
                          ),
                        ),
                      ),
                      const SizedBox(height: 20.0),
                      ElevatedButton(
                        onPressed: _generateCode,
                        style: ElevatedButton.styleFrom(
                          padding: const EdgeInsets.symmetric(
                              vertical: 16, horizontal: 32),
                          textStyle: const TextStyle(fontSize: 18),
                          elevation: 10, // Tạo độ cao cho nút
                          shape: RoundedRectangleBorder(
                            borderRadius: BorderRadius.circular(
                                10), // Đường viền tròn cho nút
                          ),
                          backgroundColor: const Color.fromARGB(
                              255, 219, 237, 252), // Màu nền nút
                          shadowColor: const Color.fromARGB(
                              255, 226, 235, 250), // Màu bóng đổ
                        ).copyWith(
                          elevation: WidgetStateProperty.all<double>(
                              12), // Tăng độ cao khi nhấn
                          shadowColor:
                              WidgetStateProperty.all<Color>(Colors.blue[800]!),
                        ),
                        child: const Text('Mã hóa'),
                      ),
                      const SizedBox(height: 20.0),
                      const Text(
                        'Mã code:',
                        style: TextStyle(
                            fontSize: 16, fontWeight: FontWeight.bold),
                      ),
                      const SizedBox(height: 20.0),
                      Row(
                        children: [
                          Expanded(
                            child: SelectableText(
                              _generatedCode,
                              style: const TextStyle(
                                  fontSize: 22,
                                  color: Color.fromARGB(255, 0, 0, 0)),
                            ),
                          ),
                          const Text(
                            'Copy:', // Dòng chữ thêm
                            style: TextStyle(
                                fontSize: 16, fontWeight: FontWeight.bold),
                          ),
                          IconButton(
                            icon: const Icon(Icons.copy),
                            onPressed: () {
                              Clipboard.setData(
                                  ClipboardData(text: _generatedCode));
                              ScaffoldMessenger.of(context).showSnackBar(
                                const SnackBar(
                                    content: Text('Đã sao chép mã code')),
                              );
                            },
                          ),
                        ],
                      ),
                      Center(
                        child: barcodeSvg != null
                            ? SvgPicture.string(
                                barcodeSvg!) // Hiển thị mã vạch SVG
                            : const CircularProgressIndicator(),
                      ),
                      const SizedBox(height: 30.0),
                      Row(children: [
                        ElevatedButton.icon(
                          onPressed: _pickExcelFile,
                          icon: const Icon(Icons.upload_file),
                          label: const Text('Tải lên file Excel'),
                          style: ElevatedButton.styleFrom(
                            padding: const EdgeInsets.symmetric(
                                vertical: 16, horizontal: 32),
                            textStyle: const TextStyle(fontSize: 18),
                            elevation: 10, // Tạo độ cao cho nút
                            shape: RoundedRectangleBorder(
                              borderRadius: BorderRadius.circular(
                                  10), // Đường viền tròn cho nút
                            ),
                            backgroundColor: const Color.fromARGB(
                                255, 219, 237, 252), // Màu nền nút
                            shadowColor: const Color.fromARGB(
                                255, 226, 235, 250), // Màu bóng đổ
                          ).copyWith(
                            elevation: WidgetStateProperty.all<double>(
                                12), // Tăng độ cao khi nhấn
                            shadowColor: WidgetStateProperty.all<Color>(
                                Colors.blue[800]!),
                          ),
                        ),
                        const SizedBox(width: 30.0),
                        ElevatedButton(
                          onPressed: uploadedFiles.isNotEmpty
                              ? pickAndProcessExcelFile
                              : null, // Vô hiệu hóa nút nếu không có file
                          style: ElevatedButton.styleFrom(
                            padding: const EdgeInsets.symmetric(
                                vertical: 16, horizontal: 32),
                            textStyle: const TextStyle(fontSize: 18),
                            elevation: 13,
                            backgroundColor: uploadedFiles.isNotEmpty
                                ? const Color.fromARGB(255, 219, 237,
                                    252) // Màu khi nút được kích hoạt
                                : Colors.grey, // Màu khi nút bị vô hiệu hóa
                          ),
                          child: const Text('Tạo mã file Excel'),
                        ),
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

extension on Svg {
  fromSvgString(String svgString, String s) {}
}
