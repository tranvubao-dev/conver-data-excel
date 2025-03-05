import 'dart:html' as html; // Dành cho tải file trong Flutter web
import 'dart:io';
import 'dart:math';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:html/parser.dart' show parse;
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'logic/code_generator.dart';
import 'logic/zip_file_utils.dart';
import 'logic/excel_utils.dart';
import 'logic/safe_value.dart';
import 'package:flutter_fireworks/flutter_fireworks.dart';

class SalesPage extends StatefulWidget {
  const SalesPage({super.key});

  @override
  _SalesPageState createState() => _SalesPageState();
}

class _SalesPageState extends State<SalesPage> {
  String? _fileName;
  File? _excelFile;
  String? _downloadUrl;
  List<List<String>> dataExcel = [];
  List<List<String>> listTitle = [
    [
      "Ngày",
      "Thứ tự",
      "Mã KH/NCC",
      "Tên KH/NCC",
      "Người phụ trách",
      "Kho xuất HH/NVL",
      "Loại giao dịch",
      "Hình thức thanh toán",
      "Đơn vị tiền tệ",
      "Tỷ giá",
      "Hình thức vận chuyển",
      "Mã mặt hàng",
      "Tên mặt hàng",
      "Thông số đặc tả",
      "Số lượng",
      "Đơn giá",
      "Tổng tiền hàng",
      "% Chiết khấu",
      "Số tiền chiết khấu",
      "Số tiền ngoại tệ",
      "Thành tiền sau chiết khấu",
      "% Thuế",
      "Tiền thuế",
      "Ghi chú",
      "Lập phiếu sản xuất",
      "S/L phụ",
      "Địa chỉ",
      "Số điện thoại",
    ],
  ];
  List<Map<String, dynamic>> uploadedFiles =
      []; // List để lưu các file đã tải lên
  Map<String, String> _htmlContents = {};
  bool isAuxiliaryTemplateEnabled = false;

  final fireworksController = FireworksController(
    // Define a list of colors for the fireworks explosions
    // They will be picked randomly from this list for each explosion
    colors: [
      Color(0xFFFF4C40), // Coral
      Color(0xFF6347A6), // Purple Haze
      Color(0xFF7FB13B), // Greenery
      Color(0xFF82A0D1), // Serenity Blue
      Color(0xFFF7B3B2), // Rose Quartz
      Color(0xFF864542), // Marsala
      Color(0xFFB04A98), // Orchid
      Color(0xFF008F6C), // Sea Green
      Color(0xFFFFD033), // Pastel Yellow
      Color(0xFFFF6F7C), // Pink Grapefruit
    ],
    // The fastest explosion in seconds
    minExplosionDuration: 0.5,
    // The slowest explosion in seconds
    maxExplosionDuration: 3.5,
    // The minimum number of particles in an explosion
    minParticleCount: 125,
    // The maximum number of particles in an explosion
    maxParticleCount: 275,
    // The duration for particles to fade out in seconds
    fadeOutDuration: 0.4,
  );

  void firework() {
    fireworksController.fireMultipleRockets(
        minRockets: 20,
        maxRockets: 50,
        launchWindow: Duration(milliseconds: 600));
  }

  // Hàm xóa file khỏi danh sách
  void removeFile(int index) {
    setState(() {
      uploadedFiles.removeAt(index);
    });
  }

  /////
  void resetState() {
    setState(() {
      _fileName = null;
      _excelFile = null;
      _downloadUrl = null;
      dataExcel = []; // Đặt lại thanh tiêu đề
      uploadedFiles = []; // Xóa danh sách file
      _htmlContents = {}; // Xóa nội dung HTML
    });
  }

  ///

  void convertToExcel() {
    dataExcel.addAll(listTitle);
    _htmlContents.forEach((fileName, content) {
      try {
        final document = parse(content);

        // Lấy nội dung trong thẻ <body>
        final body = document.body;
        if (body != null) {
          print('Nội dung trong <body>:');
        } else {
          print('Không tìm thấy thẻ <body> trong file HTML.');
        }
        // Tao list data
        final sampleNumberElement = document.querySelector('.main-title');
        final tables = document.querySelector('.di-value');
        if (sampleNumberElement != null && tables != null) {
          var invoiceNumber = document
              .querySelectorAll('.top-content .code-content b')[2]
              .text
              .trim();

          String cleanedInvoiceNumber =
              invoiceNumber.replaceAll(RegExp(r'\s+'), '');

          RegExp regExp = RegExp(r'\d+'); // Tìm các ký tự số (0-9)
          String sttNumber = regExp.stringMatch(cleanedInvoiceNumber) ?? '';
          sttNumber = sttNumber.padLeft(5, '0');
          sttNumber =
              sttNumber.substring(sttNumber.length - 4); // Lấy 4 ký tự cuối
          var dayElement =
              document.querySelectorAll('.title-heading div p').first.text;
          String dayText = dayElement;

          // Sử dụng RegExp để trích xuất ngày, tháng, năm
          RegExp regex = RegExp(r'Ngày (\d{1,2}) tháng (\d{1,2}) năm (\d{4})');
          var match = regex.firstMatch(dayText);

          if (match != null) {
            // Trích xuất các phần tử ngày, tháng, năm
            String day = match.group(1)!.padLeft(2, '0');
            String month = match.group(2)!.padLeft(2, '0');
            String year = match.group(3)!;

            // Kết hợp thành định dạng "20122024"
            dayElement = "$day$month$year";
          } else {
            print("Không tìm thấy ngày trong nội dung HTML.");
          }
          ///////// Title Heading
          var titleHeading = document
              .querySelectorAll('.title-heading .main-title')
              .first
              .text;

          ///////////////////////////////////////////////
          var customerName = document // Tên người mua
              .querySelectorAll('.list-fill-out .data-item')[5]
              .querySelector('.di-value div')!
              .text;
          ///////////////////////////////////////////////

          var customerCode = document // Mã số thuế người bán
              .querySelectorAll('.list-fill-out .data-item')[1]
              .querySelector('.di-value div')!
              .text;
          ///////////////////////////////////////////////
          var buyerCode = document // Mã số thuế người mua
              .querySelectorAll('.list-fill-out .data-item')[7]
              .querySelector('.di-value div')!
              .text;
          ///////////////////////////////////////////////
          var customerAddress = document // Địa chỉ người mua
              .querySelectorAll('.list-fill-out .data-item')[8]
              .querySelector('.di-value div')!
              .text;
          ///////////////////////////////////////////////
          var customerPhone = document //Số tài khoản người bán
              .querySelectorAll('.list-fill-out .data-item')[9]
              .querySelector('.di-value div')!
              .text;
          /////////////////////////////////////////////////

          var dataRows = document
              .querySelectorAll('.content-info .res-tb')
              .first
              .querySelectorAll('tbody tr');
          List<List<String>> rows = [];
          for (var row in dataRows) {
            var cells = row.querySelectorAll('td');

            rows.add(cells.map((e) => e.text.trim()).toList());
          }
          List<List<String>> dataExcelDetail;
          // Các index cần lấy
          if (titleHeading.trim().toUpperCase() == "HOÁ ĐƠN GIÁ TRỊ GIA TĂNG") {
            dataExcelDetail = rows.map((row) {
              // Tạo một danh sách mới với chuỗi rỗng ở các vị trí cụ thể
              return [
                dayElement, // Ngày
                sttNumber, // Thứ tự
                buyerCode, // Mã KH/NCC
                customerName, // Tên KH/NCC
                customerCode, // Người phụ trách
                "KX", // Kho xuất HH/NVL
                "", // Loại giao dịch
                "CK", // Hình thức thanh toán
                "", // Đơn vị tiền tệ
                "", // Tỷ giá
                "NVGH", // Hình thức vận chuyển
                CodeGenerator.generateItemCode(
                    SafeValueHandler.safeValue(row, 2)
                        .toString()), // Mã mặt hàng
                SafeValueHandler.safeValue(row, 2)
                    .toString()
                    .replaceAll(RegExp(r'[^a-zA-Z0-9À-ỹ ]'),
                        '') // Loại bỏ ký tự đặc biệt
                    .toUpperCase() // Chuyển thành chữ in hoa
                    .substring(
                        0,
                        min(
                            95,
                            SafeValueHandler.safeValue(row, 2)
                                .toString()
                                .replaceAll(RegExp(r'[^a-zA-Z0-9À-ỹ ]'), '')
                                .length)), // Tên mặt hàng
                SafeValueHandler.safeValue(row, 3)
                    .toString()
                    .toUpperCase(), // Thông số đặc tả
                SafeValueHandler.safeValue(row, 4)
                    .toString()
                    .replaceAll(".", ""), // Số lượng
                SafeValueHandler.safeValue(row, 5)
                    .toString()
                    .replaceAll(".", ""), // Đơn giá
                SafeValueHandler.safeValue(row, 8)
                    .toString()
                    .replaceAll(".", ""), // Tổng tiền hàng
                SafeValueHandler.safeValue(row, 6)
                    .toString()
                    .replaceAll("%", ""), // % Chiết khấu
                "", // Số tiền chiết khấu
                "", // Số tiền ngoại tệ
                "", // Thành tiền sau chiết khấu
                SafeValueHandler.safeValue(row, 7)
                    .toString()
                    .replaceAll("%", ""), // % Thuế
                "", // Tiền thuế
                cleanedInvoiceNumber, // Ghi chú
                "", // Lập phiếu sản xuất
                "", // S/L phụ
                customerAddress, // Địa chỉ
                customerPhone, // Số điện thoại
              ];
            }).toList();
          } else {
            dataExcelDetail = rows.map((row) {
              // Tạo một danh sách mới với chuỗi rỗng ở các vị trí cụ thể
              return [
                dayElement, // Ngày
                sttNumber, // Thứ tự
                buyerCode, // Mã KH/NCC
                customerName, // Tên KH/NCC
                customerCode, // Người phụ trách
                "KX", // Kho xuất HH/NVL
                "", // Loại giao dịch
                "CK", // Hình thức thanh toán
                "", // Đơn vị tiền tệ
                "", // Tỷ giá
                "NVGH", // Hình thức vận chuyển
                CodeGenerator.generateItemCode(
                    SafeValueHandler.safeValue(row, 2)
                        .toString()), // Mã mặt hàng
                SafeValueHandler.safeValue(row, 2)
                    .toString()
                    .replaceAll(RegExp(r'[^a-zA-Z0-9À-ỹ ]'),
                        '') // Loại bỏ ký tự đặc biệt
                    .toUpperCase() // Chuyển thành chữ in hoa
                    .substring(
                        0,
                        min(
                            95,
                            SafeValueHandler.safeValue(row, 2)
                                .toString()
                                .replaceAll(RegExp(r'[^a-zA-Z0-9À-ỹ ]'), '')
                                .length)), // Tên mặt hàng
                SafeValueHandler.safeValue(row, 3)
                    .toString()
                    .toUpperCase(), // Thông số đặc tả
                SafeValueHandler.safeValue(row, 4)
                    .toString()
                    .replaceAll(".", ""), // Số lượng
                SafeValueHandler.safeValue(row, 5)
                    .toString()
                    .replaceAll(".", ""), // Đơn giá
                SafeValueHandler.safeValue(row, 7)
                    .toString()
                    .replaceAll(".", ""), // Tổng tiền hàng
                SafeValueHandler.safeValue(row, 6)
                    .toString()
                    .replaceAll("%", ""), // % Chiết khấu
                "", // Số tiền chiết khấu
                "", // Số tiền ngoại tệ
                "", // Thành tiền sau chiết khấu
                "0", // % Thuế
                "", // Tiền thuế
                cleanedInvoiceNumber, // Ghi chú
                "", // Lập phiếu sản xuất
                "", // S/L phụ
                customerAddress, // Địa chỉ
                customerPhone, // Số điện thoại
              ];
            }).toList();
          }

          for (var row in dataExcelDetail) {
            if (row.length > 17) {
              // Lấy giá trị tiền và phần trăm
              String valueString = row[16]; // Số tiền trước thuế
              String discountString = row[17]; // % Chiết khấu
              String percentString = row[21]; // % Thuế

              // Chuyển sang số
              double value = double.tryParse(valueString) ?? 0;
              double percent = double.tryParse(percentString) ?? 0;
              double discount = double.tryParse(discountString) ?? 0;

              // Tính giá trị mới
              double result = value * discount / 100;
              double discountMoney = value - result;
              double percentMoney = discountMoney * percent / 100;

              // Chèn lại giá trị vào vị trí thứ 18
              row[18] =
                  result.toStringAsFixed(0); // Định dạng 0 chữ số thập phân
              row[20] = discountMoney
                  .toStringAsFixed(0); // Định dạng 0 chữ số thập phân
              row[22] = percentMoney.toStringAsFixed(0);
            }
          }

          // Thêm data vào bảng excel có tiêu đề
          dataExcel.addAll(dataExcelDetail);
          // Gọi hàm tạo file Excel
        } else {
          print('Không tìm thấy giá trị mẫu số.');
        }
      } catch (e) {
        print('Đã xảy ra lỗi: $e');
      }
    });
    ExcelUtils.createExcelFile(
        dataExcel, "template_banhang.xlsx", isAuxiliaryTemplateEnabled, false);
    resetState();
    firework();
  }

  // Tải file Excel về
  void downloadExcel() async {
    if (_excelFile == null) return;

    // Kiểm tra quyền lưu trữ
    var status = await Permission.storage.request();
    if (status.isGranted) {
      // Lưu file vào bộ nhớ thiết bị
      final directory = await getExternalStorageDirectory();
      final savedPath = '${directory?.path}/downloaded_file.xlsx';
      await _excelFile!.copy(savedPath);

      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('File đã được tải về: $savedPath')),
      );
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Cần cấp quyền lưu trữ')),
      );
    }
  }

  void handlePickZipFiles() async {
    // Sử dụng hàm từ file zip_file_utils.dart
    final files = await ZipFileUtils.pickZipFiles();

    if (files.isNotEmpty) {
      setState(() {
        uploadedFiles = files;
      });

      // Duyệt qua từng file để xử lý nội dung
      for (var file in files) {
        final Uint8List? zipBytes = file['data'];
        if (zipBytes != null) {
          final invoiceContent =
              await ZipFileUtils.extractInvoiceHtml(zipBytes);

          if (invoiceContent != null) {
            setState(() {
              _htmlContents[file['name']] = invoiceContent;
            });
          } else {
            ScaffoldMessenger.of(context).showSnackBar(
              SnackBar(
                  content: Text(
                      'Không tìm thấy file invoice.html trong: ${file['name']}')),
            );
            removeFile(
                uploadedFiles.indexWhere((f) => f['name'] == file['name']));
          }
        }
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
        appBar: AppBar(
          title: const Text(
            'Tạo Template bán hàng từ Hoá Đơn Giá Trị Gia Tăng',
            style: TextStyle(color: Colors.black),
          ),
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
                    child: Center(
                      child: Column(
                        mainAxisSize: MainAxisSize.min,
                        children: [
                          ElevatedButton.icon(
                            onPressed: () => handlePickZipFiles(),
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
                            icon: const Icon(Icons.upload_file),
                            label: const Text('Tải lên file .zip Hoá Đơn'),
                          ),
                          const SizedBox(height: 20),
                          if (_fileName != null) const SizedBox(height: 20),
                          Expanded(
                            child: ListView.builder(
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
                                    title: Text(file['name']),
                                    trailing: IconButton(
                                      icon: const Icon(Icons.delete,
                                          color: Colors.red),
                                      onPressed: () =>
                                          removeFile(index), // Xóa file
                                    ),
                                  ),
                                );
                              },
                            ),
                          ),
                          const SizedBox(height: 20),
                          // 🟢 Nút ON/OFF (Luôn bật - Không thể tắt)
                          Row(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              const Text('Tạo Template Phụ',
                                  style: TextStyle(fontSize: 18)),
                              const SizedBox(width: 10),
                              Switch(
                                value: isAuxiliaryTemplateEnabled,
                                onChanged: (bool value) {
                                  setState(() {
                                    isAuxiliaryTemplateEnabled = value;
                                  });
                                },
                                activeColor: Colors.green,
                              ),
                              SizedBox(
                                width:
                                    40, // Định nghĩa chiều rộng cố định để tránh dịch chuyển
                                child: Text(
                                  isAuxiliaryTemplateEnabled ? 'ON' : 'OFF',
                                  textAlign:
                                      TextAlign.center, // Canh giữa chữ ON/OFF
                                  style: TextStyle(
                                    fontSize: 18,
                                    fontWeight: FontWeight.bold,
                                    color: isAuxiliaryTemplateEnabled
                                        ? Colors.green
                                        : Colors.grey,
                                  ),
                                ),
                              ),
                            ],
                          ),
                          const SizedBox(height: 20),
                          ElevatedButton(
                            onPressed: uploadedFiles.isNotEmpty
                                ? convertToExcel
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
                            child: const Text('Tạo Template bán hàng Excel'),
                          ),
                          const SizedBox(height: 20),
                          if (_downloadUrl != null)
                            ElevatedButton(
                              onPressed: () {
                                final anchor =
                                    html.AnchorElement(href: _downloadUrl!)
                                      ..target = 'blank'
                                      ..download = 'file.xlsx';
                              },
                              style: ElevatedButton.styleFrom(
                                padding: const EdgeInsets.symmetric(
                                    vertical: 16, horizontal: 32),
                                textStyle: const TextStyle(fontSize: 18),
                              ),
                              child: const Text('Tải về tệp Excel'),
                            ),
                        ],
                      ),
                    ),
                  ),
                ],
              ),
            ),
          ),
          Container(
            child: FireworksDisplay(controller: fireworksController),
          ),
        ]));
  }
}
