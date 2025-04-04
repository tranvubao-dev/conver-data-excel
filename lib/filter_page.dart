import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'package:flutter/services.dart';
import 'package:flutter_fireworks/fireworks_display.dart';
import 'package:html_to_excel/common_fireworks.dart';
import 'package:html_to_excel/logic/excel_utils.dart';
import 'package:lottie/lottie.dart';

class FilterDataPage extends StatefulWidget {
  const FilterDataPage({super.key});

  @override
  _FilterDataPageState createState() => _FilterDataPageState();
}

class _FilterDataPageState extends State<FilterDataPage> {
  List<Map<String, dynamic>> uploadedFileA = [];
  List<Map<String, dynamic>> uploadedFileB = [];
  List<List<dynamic>> dataExcel = [];
  bool isProcessing = false;
  List<List<dynamic>> dataA = [];
  List<List<dynamic>> dataB = [];
  Map<String, int> invoiceCounterSale = {};
  Map<String, int> invoiceCounterBuy = {};
  final CommonFireworks commonFireworks = CommonFireworks();
  //////////////

  void _pickExcelFile(bool fileA) async {
    // _progressController.add(0);
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'xls'],
    );

    if (result != null) {
      setState(() {
        if (fileA) {
          uploadedFileA = [
            {
              'name': result.files.single.name,
              'bytes': base64Encode(result.files.single
                  .bytes!), // Dùng base64Encode để chuyển đổi bytes thành chuỗi
            }
          ]; // Luôn chỉ chứa 1 file
        } else {
          uploadedFileB = [
            {
              'name': result.files.single.name,
              'bytes': base64Encode(result.files.single
                  .bytes!), // Dùng base64Encode để chuyển đổi bytes thành chuỗi
            }
          ];
        }
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
      uploadedFileA.clear(); // Xóa file đã tải lên
      uploadedFileB.clear();
    });
  }

  Future<void> pickAndProcessExcelFile() async {
    if (uploadedFileA.isNotEmpty) {
      getDataExcelFile(uploadedFileA, dataA);
      dataExcel.addAll(dataA);
    }
    if (uploadedFileB.isNotEmpty) {
      getDataExcelFile(uploadedFileB, dataB);
      dataExcel.addAll(dataB);
    }

    Map<String, List<List<dynamic>>> groupedData = {};
    for (var row in dataExcel) {
      if (row.length > 3) {
        String invoiceNumber = row[2]; // "Số:1514" hoặc "Số:1492"
        String date = row[3];
        String key = "$invoiceNumber - $date";

        if (!groupedData.containsKey(key)) {
          groupedData[key] = [];
        }
        groupedData[key]!.add(row);
      }
    }

    // Bước 2: Sắp xếp nhóm theo ngày row[3]
    var sortedEntries = groupedData.entries.toList()
      ..sort((a, b) {
        String dateA = a.value.first[3]; // Lấy ngày từ nhóm A
        String dateB = b.value.first[3]; // Lấy ngày từ nhóm B

        DateTime parsedDateA = DateTime.parse(
            "${dateA.substring(4, 8)}-${dateA.substring(2, 4)}-${dateA.substring(0, 2)}");
        DateTime parsedDateB = DateTime.parse(
            "${dateB.substring(4, 8)}-${dateB.substring(2, 4)}-${dateB.substring(0, 2)}");
        return parsedDateA.compareTo(parsedDateB); // Sắp xếp tăng dần theo ngày
      });

    // Bước 3: Chuyển về List<List<String>>
    List<List<dynamic>> sortedDataDay = [];
    for (var entry in sortedEntries) {
      sortedDataDay.addAll(
          entry.value); // Thêm tất cả dòng của nhóm vào danh sách kết quả
    }

    List<List<dynamic>> convertedData = sortedDataDay.map((row) {
      String rawDate = row[3]; // Lấy ngày dạng ddMMyyyy
      if (rawDate.length == 8) {
        String formattedDate =
            "${rawDate.substring(0, 2)}/${rawDate.substring(2, 4)}/${rawDate.substring(4, 8)}";
        row[3] = formattedDate; // Cập nhật ngày mới vào danh sách
      }
      return row;
    }).toList();

    int dataNumberPN = 1;
    int dataNumberPX = 1;
    String previousRow2 = ""; // Lưu giá trị row[2] của dòng trước đó
    String previousRow1 = ""; // Lưu giá trị row[1] của dòng trước đó

    for (var row in convertedData) {
      if (row.length > 2 && row[1].isNotEmpty) {
        String prefix = row[1].substring(0, 2);
        if (row[2] == previousRow2) {
          row[1] = previousRow1;
        } else {
          if (prefix == "PN") {
            row[1] = "PN$dataNumberPN";
            dataNumberPN++;
          } else if (prefix == "PX") {
            row[1] = "PX$dataNumberPX";
            dataNumberPX++;
          } else {
            print("Không xác định");
          }
        }

        // Cập nhật giá trị của dòng trước đó
        previousRow2 = row[2];
        previousRow1 = row[1];
      }
    }

    ExcelUtils.createExcel(convertedData, "FileTổngHợp .xlsx");
    resetState();
    commonFireworks.firework();

    /// Pháo hoa
    setState(() {
      isProcessing = false; // Cho phép nhấn lại sau khi hoàn thành
    });
  }

  void getDataExcelFile(
      List<Map<String, dynamic>> input, List<List<dynamic>> output) async {
    try {
      // Lấy bytes từ uploadedFiles
      String base64String = input.first['bytes']!;
      var bytes = base64Decode(base64String);
      var excel = Excel.decodeBytes(bytes);

      // Lấy sheet đầu tiên
      var sheetName = excel.tables.keys.first;
      var table = excel.tables[sheetName];

      if (table == null || table.rows.length <= 2) {
        print("Lỗi: File không có đủ dữ liệu!");
        return;
      }

      // Xử lý từng dòng dữ liệu (bắt đầu từ dòng thứ 3)
      for (var i = 2; i < table.rows.length; i++) {
        if (table.rows[i].isEmpty) {
          print("Bỏ qua dòng $i vì không có dữ liệu.");
          continue;
        }

        // Chuyển đổi tất cả giá trị trong hàng thành String
        List<dynamic> rowData =
            table.rows[i].map((cell) => cell?.value?.toString() ?? "").toList();

        // Đảm bảo danh sách có ít nhất 3 cột trước khi truy xuất
        // String column1 = rowData.length > 0 ? rowData[0] : "N/A";
        // String column2 = rowData.length > 1 ? rowData[1] : "N/A";
        String column3 = rowData.length > 2 ? rowData[2] : "N/A";
        String column4 = rowData.length > 3 ? rowData[3] : "N/A";
        String column5 = rowData.length > 4 ? rowData[4] : "N/A";
        String column6 = rowData.length > 5 ? rowData[5] : "N/A";
        String column7 = rowData.length > 6 ? rowData[6] : "N/A";
        String column8 = rowData.length > 7 ? rowData[7] : "N/A";
        String column9 = rowData.length > 8 ? rowData[8] : "N/A";
        String column10 = rowData.length > 9 ? rowData[9] : "N/A";
        String column11 = rowData.length > 10 ? rowData[10] : "N/A";
        String column12 = rowData.length > 11 ? rowData[11] : "N/A";
        String column13 = rowData.length > 12 ? rowData[12] : "N/A";
        String column14 = rowData.length > 13 ? rowData[13] : "N/A";
        var dayElement = column4;
        String dateText = dayElement;

        RegExp regex = RegExp(r'(\d{1,2})/(\d{1,2})/(\d{4})');
        var match = regex.firstMatch(dateText);
        if (match != null) {
          // Trích xuất ngày, tháng, năm
          String day = match.group(1)!.padLeft(2, '0'); // Đảm bảo 2 chữ số
          String month = match.group(2)!.padLeft(2, '0');
          String year = match.group(3)!;

          // Kết hợp thành "ddMMyyyy"
          dayElement = "$day$month$year";
        }
        String key = "$column3 - $dayElement";
        String counterFormatted = column7.isEmpty
            ? getInvoiceCounter(invoiceCounterBuy, key)
            : getInvoiceCounter(invoiceCounterSale, key);
        String numberVotes = "CTKT";
        if (column6.isEmpty && column7.isEmpty) {
          numberVotes = "";
        } else {
          if (column11 == "642") {
            numberVotes = "CTKT";
          } else {
            numberVotes =
                column7.isEmpty ? "PN$counterFormatted" : "PX$counterFormatted";
          }
        }

        List<dynamic> sortedRow = [
          "",
          numberVotes,
          column3,
          dayElement,
          column5,
          column6,
          column7,
          column8,
          column9,
          column10,
          column11,
          column12,
          column13,
          column14,
        ];

        // Thêm vào dataBank
        output.add(sortedRow);
      }
      print("4 - Dữ liệu đã xử lý xong, chuẩn bị ghi file.");
    } catch (e) {
      print("Lỗi xử lý file: $e");
    }
  }

  void resetState() {
    setState(() {
      dataExcel = []; // Đặt lại thanh tiêu đề
      uploadedFileA = []; // Xóa danh sách file
      uploadedFileB = [];
      dataA = [];
      dataB = [];
      invoiceCounterSale.clear();
      invoiceCounterBuy.clear();
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
        appBar: AppBar(
          title: const Text('Gộp 2 file dữ liệu'),
          backgroundColor: Colors.transparent, // Làm trong suốt AppBar
          elevation: 0, // Xóa bóng AppBar
        ),
        body: Stack(children: [
          Container(
            child: Padding(
              padding: const EdgeInsets.all(16.0),
              child: Column(
                children: [
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.center,
                      children: [
                        Row(
                            mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                            children: [
                              ElevatedButton.icon(
                                onPressed: () => _pickExcelFile(true),
                                style: ElevatedButton.styleFrom(
                                  padding: const EdgeInsets.symmetric(
                                      vertical: 16, horizontal: 32),
                                  textStyle: const TextStyle(fontSize: 18),
                                  elevation: 10,
                                  shape: RoundedRectangleBorder(
                                    borderRadius: BorderRadius.circular(10),
                                  ),
                                  backgroundColor: const Color.fromARGB(
                                      255, 219, 237, 252), // Màu nền nút
                                  shadowColor:
                                      const Color.fromARGB(255, 226, 235, 250),
                                ).copyWith(
                                  elevation:
                                      WidgetStateProperty.all<double>(12),
                                  shadowColor: WidgetStateProperty.all<Color>(
                                      Colors.blue[800]!),
                                ),
                                icon: const Icon(Icons.upload_file),
                                label: const Text('Tải lên file Excel A'),
                              ),
                              const SizedBox(width: 30.0),
                              ElevatedButton.icon(
                                onPressed: () => _pickExcelFile(false),
                                style: ElevatedButton.styleFrom(
                                  padding: const EdgeInsets.symmetric(
                                      vertical: 16, horizontal: 32),
                                  textStyle: const TextStyle(fontSize: 18),
                                  elevation: 10,
                                  shape: RoundedRectangleBorder(
                                    borderRadius: BorderRadius.circular(10),
                                  ),
                                  backgroundColor: const Color.fromARGB(
                                      255, 219, 237, 252), // Màu nền nút
                                  shadowColor:
                                      const Color.fromARGB(255, 226, 235, 250),
                                ).copyWith(
                                  elevation:
                                      WidgetStateProperty.all<double>(12),
                                  shadowColor: WidgetStateProperty.all<Color>(
                                      Colors.blue[800]!),
                                ),
                                icon: const Icon(Icons.upload_file),
                                label: const Text('Tải lên file Excel B'),
                              ),
                            ]),
                        const SizedBox(height: 40),
                        Row(
                            mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                            children: [
                              Expanded(
                                child: uploadedFileA.isEmpty
                                    ? const Center(
                                        child: Text(
                                            'Chưa có file nào được tải lên'))
                                    : ListView.builder(
                                        shrinkWrap: true,
                                        itemCount: uploadedFileA.length,
                                        itemBuilder: (context, index) {
                                          final file = uploadedFileA[index];
                                          return Card(
                                            elevation: 10,
                                            margin: const EdgeInsets.symmetric(
                                                vertical: 8, horizontal: 16),
                                            child: ListTile(
                                              leading: const Icon(
                                                  Icons.file_present,
                                                  size: 40),
                                              title: Text(file['name'] ?? ''),
                                              trailing: IconButton(
                                                icon: const Icon(Icons.delete,
                                                    color: Colors.red),
                                                onPressed: () =>
                                                    _removeFile(index),
                                              ),
                                            ),
                                          );
                                        },
                                      ),
                              ),
                              Expanded(
                                child: uploadedFileB.isEmpty
                                    ? const Center(
                                        child: Text(
                                            'Chưa có file nào được tải lên'))
                                    : ListView.builder(
                                        shrinkWrap: true,
                                        itemCount: uploadedFileB.length,
                                        itemBuilder: (context, index) {
                                          final file = uploadedFileB[index];
                                          return Card(
                                            elevation: 10,
                                            margin: const EdgeInsets.symmetric(
                                                vertical: 8, horizontal: 16),
                                            child: ListTile(
                                              leading: const Icon(
                                                  Icons.file_present,
                                                  size: 40),
                                              title: Text(file['name'] ?? ''),
                                              trailing: IconButton(
                                                icon: const Icon(Icons.delete,
                                                    color: Colors.red),
                                                onPressed: () =>
                                                    _removeFile(index),
                                              ),
                                            ),
                                          );
                                        },
                                      ),
                              ),
                            ]),
                        const SizedBox(height: 80),
                        ElevatedButton(
                          onPressed: uploadedFileA.isNotEmpty ||
                                  uploadedFileB.isNotEmpty
                              ? pickAndProcessExcelFile
                              : null, // Vô hiệu hóa nút nếu không có file
                          style: ElevatedButton.styleFrom(
                            padding: const EdgeInsets.symmetric(
                                vertical: 16, horizontal: 32),
                            textStyle: const TextStyle(fontSize: 18),
                            elevation: 13,
                            backgroundColor: uploadedFileA.isNotEmpty ||
                                    uploadedFileB.isNotEmpty
                                ? const Color.fromARGB(255, 219, 237,
                                    252) // Màu khi nút được kích hoạt
                                : Colors.grey, // Màu khi nút bị vô hiệu hóa
                          ),
                          child: const Text('Tạo File tổng hợp Excel'),
                        ),
                      ],
                    ),
                  ),
                ],
              ),
            ),
          ),
          Container(
            child: FireworksDisplay(
                controller: commonFireworks.fireworksController),
          ),
          // Positioned(
          //   bottom: 16,
          //   left: 16,
          // child: FloatingActionButton(
          //   onPressed: () {
          //     // Xử lý sự kiện khi nhấn vào icon
          //     print("Icon ở góc trái được nhấn!");
          //   },
          //   // backgroundColor: Colors.blue,
          //   child: Lottie.asset(
          //     'assets/rabbit.json', // Animation từ ảnh cá nhân của bạn
          //     width: 120,
          //     height: 120,
          //     repeat: true,
          //   ),
          // ),
          // ),
        ]));
  }

  String getInvoiceCounter(Map<String, int> invoiceCounter, String key) {
    if (!invoiceCounter.containsKey(key)) {
      invoiceCounter[key] = invoiceCounter.length + 1;
    }
    return invoiceCounter[key]!.toString().padLeft(3, '0');
  }
}
