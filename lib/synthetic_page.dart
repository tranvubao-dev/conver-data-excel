import 'dart:math';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:html/parser.dart';
import 'package:html_to_excel/logic/excel_utils.dart';
import 'package:html_to_excel/logic/safe_value.dart';
import 'package:html_to_excel/logic/zip_file_utils.dart';
import 'logic/code_generator.dart';
import 'dart:html' as html; // Dành cho tải file trong Flutter web

class SyntheticPage extends StatefulWidget {
  const SyntheticPage({super.key});

  @override
  _SyntheticPageState createState() => _SyntheticPageState();
}

class _SyntheticPageState extends State<SyntheticPage> {
  Uint8List? barcodePng;
  String? barcodeSvg;
  List<Map<String, dynamic>> uploadedFileSale =
      []; // List để lưu các file hoá đơn bán hàng
  List<Map<String, dynamic>> uploadedFileBuy =
      []; // List để lưu các file hoá đơn mua hàng
  List<List<String>> dataExcel = [];
  String? _fileName;

  bool isAuxiliaryTemplateEnabled = false;
  Map<String, Object> _htmlContents = {};
  //////////////

  void resetState() {
    setState(() {
      dataExcel = []; // Đặt lại thanh tiêu đề
      uploadedFileSale = []; // Xóa danh sách file
      uploadedFileBuy = []; // Xóa danh sách file
      _htmlContents = {}; // Xóa nội dung HTML
    });
  }

  void handlePickZipFiles(bool isBuy) async {
    // Sử dụng hàm từ file zip_file_utils.dart
    final files = await ZipFileUtils.pickZipFiles();

    if (files.isNotEmpty) {
      setState(() {
        isBuy ? uploadedFileBuy = files : uploadedFileSale = files;
      });

      // Duyệt qua từng file để xử lý nội dung
      for (var file in files) {
        final Uint8List? zipBytes = file['data'];
        if (zipBytes != null) {
          final invoiceContent =
              await ZipFileUtils.extractInvoiceHtml(zipBytes);

          if (invoiceContent != null) {
            setState(() {
              _htmlContents[file['name']] = {
                'content': invoiceContent,
                'isBuy': isBuy, // Thêm thông tin loại hóa đơn
              };
            });
          } else {
            ScaffoldMessenger.of(context).showSnackBar(
              SnackBar(
                  content: Text(
                      'Không tìm thấy file invoice.html trong: ${file['name']}')),
            );
            removeFile(
                isBuy
                    ? uploadedFileBuy
                        .indexWhere((f) => f['name'] == file['name'])
                    : uploadedFileSale
                        .indexWhere((f) => f['name'] == file['name']),
                isBuy);
          }
        }
      }
    }
  }

  void removeFile(int index, bool isBuy) {
    setState(() {
      isBuy
          ? uploadedFileBuy.removeAt(index)
          : uploadedFileSale.removeAt(index);
    });
  }

  void convertToExcel() {
    // dataExcel.addAll(listTitle);
    _htmlContents.forEach((fileName, data) {
      bool isBuy = (data as Map<String, dynamic>)['isBuy']; // Lấy giá trị isBuy
      String content =
          (data as Map<String, dynamic>)['content']; // Lấy nội dung hóa đơn
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
          var customerName = document // Tên người bán
              .querySelectorAll('.list-fill-out .data-item')[0]
              .querySelector('.di-value div')!
              .text;

          ///////////////////////////////////////////////
          var customerCode = document // Mã số thuế người bán
              .querySelectorAll('.list-fill-out .data-item')[1]
              .querySelector('.di-value div')!
              .text;
          ///////////////////////////////////////////////
          var customerAddress = document // Địa chỉ người bán
              .querySelectorAll('.list-fill-out .data-item')[2]
              .querySelector('.di-value div')!
              .text;
          ///////////////////////////////////////////////
          var customerPhone = document //Số điện thoại người bán
              .querySelectorAll('.list-fill-out .data-item')[3]
              .querySelector('.di-value div')!
              .text;
          /////////////////////////////////////////////////
          var buyerCode = document // Mã số thuế người mua
              .querySelectorAll('.list-fill-out .data-item')[7]
              .querySelector('.di-value div')!
              .text;

          ///////////////////////////////////////////////
          var customerNames = document // Tên người mua
              .querySelectorAll('.list-fill-out .data-item')[5]
              .querySelector('.di-value div')!
              .text;

          ///////////// tổng tiền thuế
          var taxMoney = document
              .querySelectorAll('.table-horizontal-wrapper .res-tb')[
                  1] // Chọn bảng thứ 2 trong .table-horizontal-wrapper
              .querySelectorAll(
                  'tbody tr')[1] // Chọn dòng thứ 2 (dòng Tổng tiền thuế)
              .querySelectorAll('td')[1] // Chọn cột thứ 2 (chứa số 37.840)
              .text
              .trim();
          ///////////// tổng tiền thuế
          var sumMoney = document
              .querySelectorAll('.table-horizontal-wrapper .res-tb')[
                  1] // Chọn bảng thứ 2 trong .table-horizontal-wrapper
              .querySelectorAll('tbody tr')[0] //
              .querySelectorAll('td')[1]
              .text
              .trim();

          var dataRows = document
              .querySelectorAll('.content-info .res-tb')
              .first
              .querySelectorAll('tbody tr');
          List<List<String>> rows = [];
          for (var row in dataRows) {
            var cells = row.querySelectorAll('td');
            rows.add(cells.map((e) => e.text.trim()).toList());
          }
          // Các index cần lấy
          List<List<String>> dataExcelDetail;
          // Hoá đơn giá trị gia tăng
          if (titleHeading.trim().toUpperCase() == "HOÁ ĐƠN GIÁ TRỊ GIA TĂNG") {
            dataExcelDetail = rows.map((row) {
              // Tạo một danh sách mới với chuỗi rỗng ở các vị trí cụ thể
              return [
                "",
                "",
                cleanedInvoiceNumber, // Ghi chú
                dayElement, // Ngày
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
                                .length)),
                isBuy
                    ? CodeGenerator.generateItemCode(
                        SafeValueHandler.safeValue(row, 2).toString())
                    : "", // Mã mặt hàng
                isBuy
                    ? ""
                    : CodeGenerator.generateItemCode(
                        SafeValueHandler.safeValue(row, 2).toString()),
                SafeValueHandler.safeValue(row, 3)
                    .toString()
                    .toUpperCase(), // Thông số
                SafeValueHandler.safeValue(row, 4)
                    .toString()
                    .replaceAll(".", ""), // Số lượng
                isBuy
                    ? SafeValueHandler.safeValue(row, 8)
                        .toString()
                        .replaceAll(".", "")
                    : "", // Số tiền trước thuế
                isBuy ? "" : "632",
                "",
                isBuy ? "331" : "",
                customerCode,
///////////////////////////
                dayElement, // Ngày
                sttNumber, // Thứ tự
                customerCode, // Mã KH/NCC
                customerName, // Tên KH/NCC
                buyerCode, // Người phụ trách
                "KN", // Kho nhập HH/NVL
                "", // Đơn vị tiền tệ
                "", // Tỷ giá
                "CN", // Hình thức thanh toán
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
                                .length)),
                SafeValueHandler.safeValue(row, 3)
                    .toString()
                    .toUpperCase(), // Thông số
                SafeValueHandler.safeValue(row, 4)
                    .toString()
                    .replaceAll(".", ""), // Số lượng
                SafeValueHandler.safeValue(row, 5)
                    .toString()
                    .replaceAll(".", ""), // Đơn giá
                "", // Số tiền ngoại tệ
                SafeValueHandler.safeValue(row, 8)
                    .toString()
                    .replaceAll(".", ""), // Số tiền trước thuế
                SafeValueHandler.safeValue(row, 6)
                    .toString()
                    .replaceAll(".", ""), // Số tiền chiết khấu
                SafeValueHandler.safeValue(row, 7)
                    .toString()
                    .replaceAll("%", ""), // % Thuế suất
                "", // Tiền thuế
                "", // Số tiền thanh toán
                cleanedInvoiceNumber, // Ghi chú
                "", // S/L phụ
                customerAddress, // Địa chỉ
                customerPhone, // Số điện thoại
              ];
            }).toList();
          } else {
            ///HOÁ ĐƠN BÁN HÀNG
            dataExcelDetail = rows.map((row) {
              // Tạo một danh sách mới với chuỗi rỗng ở các vị trí cụ thể
              return [
                "",
                "",
                cleanedInvoiceNumber, // Ghi chú
                dayElement, // Ngày
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
                                .length)),
                CodeGenerator.generateItemCode(
                    SafeValueHandler.safeValue(row, 2)
                        .toString()), // Mã mặt hàng
                "",
                SafeValueHandler.safeValue(row, 3)
                    .toString()
                    .toUpperCase(), // Thông số
                SafeValueHandler.safeValue(row, 4)
                    .toString()
                    .replaceAll(".", ""), // Số lượng
                isBuy
                    ? SafeValueHandler.safeValue(row, 7)
                        .toString()
                        .replaceAll(".", "")
                    : "", // Số tiền trước thuế
                isBuy ? "" : "632",
                "",
                isBuy ? "331" : "",
                customerCode,

                /////////
                dayElement, // Ngày
                sttNumber, // Thứ tự
                customerCode, // Mã KH/NCC
                customerName, // Tên KH/NCC
                buyerCode, // Người phụ trách
                "KN", // Kho nhập HH/NVL
                "", // Đơn vị tiền tệ
                "", // Tỷ giá
                "CN", // Hình thức thanh toán
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
                                .length)),
                SafeValueHandler.safeValue(row, 3)
                    .toString()
                    .toUpperCase(), // Thông số
                SafeValueHandler.safeValue(row, 4)
                    .toString()
                    .replaceAll(".", ""), // Số lượng
                SafeValueHandler.safeValue(row, 5)
                    .toString()
                    .replaceAll(".", ""), // Đơn giá
                "", // Số tiền ngoại tệ
                SafeValueHandler.safeValue(row, 7)
                    .toString()
                    .replaceAll(".", ""), // Số tiền trước thuế
                SafeValueHandler.safeValue(row, 6)
                    .toString()
                    .replaceAll(".", ""), // Số tiền chiết khấu
                "0", // % Thuế suất
                "", // Tiền thuế
                "", // Số tiền thanh toán
                cleanedInvoiceNumber, // Ghi chú
                "", // S/L phụ
                customerAddress, // Địa chỉ
                customerPhone, // Số điện thoại
              ];
            }).toList();
          }

          List<List<String>> listPercentBuy = [
            [
              "",
              "",
              cleanedInvoiceNumber,
              dayElement,
              "Tiền thuế",
              "",
              "",
              "",
              "",
              taxMoney.replaceAll(".", ""),
              isBuy ? "1331" : "131",
              isBuy ? "" : buyerCode,
              isBuy ? "" : "3331",
              "",
            ], // Thanh tieu de
          ];

          List<List<String>> listPercent = [
            [
              "",
              "",
              cleanedInvoiceNumber,
              dayElement,
              customerNames,
              "",
              "",
              "",
              "",
              sumMoney.replaceAll(".", ""),
              "131",
              buyerCode,
              "511",
              "",
            ], // Thanh tieu de
          ];

          // Thêm data vào bảng excel có tiêu đề
          dataExcel.addAll(dataExcelDetail);
          if (!isBuy) {
            dataExcel.addAll(listPercent);
          }
          dataExcel.addAll(listPercentBuy);
        } else {
          print('Không tìm thấy giá trị mẫu số.');
        }
      } catch (e) {
        print('Đã xảy ra lỗi: $e');
      }
    });

    List<List<String>> sortedData = List.from(dataExcel);

// Bước 1: Sắp xếp theo ngày và loại dữ liệu (giữ logic cũ)
    sortedData.sort((a, b) {
      DateTime dateA = parseDate(a[3]);
      DateTime dateB = parseDate(b[3]);

      // Nếu ngày khác nhau, sắp xếp theo ngày
      if (dateA.compareTo(dateB) != 0) {
        return dateA.compareTo(dateB);
      }

      // Xác định giá trị ưu tiên từ cột a[12], nếu rỗng thì lấy a[10]
      String valueA = a[12].isNotEmpty ? a[12] : a[10];
      String valueB = b[12].isNotEmpty ? b[12] : b[10];

      bool isACompany = valueA.contains("511");
      bool isBCompany = valueB.contains("511");
      bool isATax = valueA.contains("3331") || valueA.contains("1331");
      bool isBTax = valueB.contains("3331") || valueB.contains("1331");

      if (isATax && !isBTax) return 1; // "Tiền thuế" xuống cuối
      if (!isATax && isBTax) return -1; // "Tiền thuế" xuống cuối

      if (isACompany && !isBCompany) return 1; // "CÔNG TY" xuống dưới sản phẩm
      if (!isACompany && isBCompany) return -1; // "CÔNG TY" xuống dưới sản phẩm

      return 0; // Giữ nguyên thứ tự nếu cùng loại
    });

// Bước 2: Gom nhóm các phần tử có cùng a[2] lại gần nhau
    Map<String, List<List<String>>> groupedData = {};

// Nhóm các phần tử dựa trên giá trị của a[2]
    for (var row in sortedData) {
      String key = row[2]; // Lấy giá trị của cột a[2]
      if (!groupedData.containsKey(key)) {
        groupedData[key] = [];
      }
      groupedData[key]!.add(row);
    }

// Bước 3: Ghép các nhóm lại thành danh sách cuối cùng
    sortedData = groupedData.values.expand((element) => element).toList();

    List<List<String>> convertedData = sortedData.map((row) {
      String rawDate = row[3]; // Lấy ngày dạng ddMMyyyy
      if (rawDate.length == 8) {
        String formattedDate =
            "${rawDate.substring(0, 2)}/${rawDate.substring(2, 4)}/${rawDate.substring(4, 8)}";
        row[3] = formattedDate; // Cập nhật ngày mới vào danh sách
      }
      return row;
    }).toList();

    ExcelUtils.createExcel(convertedData, "FileTổngHợp .xlsx");
    resetState();
    // firework();

    /// Pháo hoa
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
        appBar: AppBar(
          title: const Text('Tạo mẫu File tổng hợp'),
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
                          Row(
                              mainAxisAlignment: MainAxisAlignment.center,
                              children: [
                                ElevatedButton.icon(
                                  onPressed: () => handlePickZipFiles(true),
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
                                  label: const Text(
                                      'Tải lên file .zip Hoá Đơn Mua Hàng'),
                                ),
                                const VerticalDivider(thickness: 2, width: 40),
                                ElevatedButton.icon(
                                  onPressed: () => handlePickZipFiles(false),
                                  style: ElevatedButton.styleFrom(
                                    padding: const EdgeInsets.symmetric(
                                        vertical: 16, horizontal: 32),
                                    textStyle: const TextStyle(fontSize: 18),
                                    elevation: 10,
                                    shape: RoundedRectangleBorder(
                                      borderRadius: BorderRadius.circular(10),
                                    ),
                                    backgroundColor: const Color.fromARGB(
                                        255, 252, 219, 237),
                                    shadowColor: const Color.fromARGB(
                                        255, 250, 226, 235),
                                  ).copyWith(
                                    elevation:
                                        WidgetStateProperty.all<double>(12),
                                    shadowColor: WidgetStateProperty.all<Color>(
                                        Colors.pink[800]!),
                                  ),
                                  icon: const Icon(Icons.upload_file),
                                  label: const Text(
                                      'Tải lên file .zip Hoá Đơn Bán Hàng'),
                                ),
                              ]),
                          const SizedBox(height: 20),
                          Row(
                              mainAxisAlignment: MainAxisAlignment.center,
                              children: [
                                if (_fileName != null)
                                  const SizedBox(height: 20),
                                Expanded(
                                  child: SizedBox(
                                    height: 400,
                                    child: ListView.builder(
                                      shrinkWrap: true,
                                      itemCount: uploadedFileBuy.length,
                                      itemBuilder: (context, index) {
                                        final file = uploadedFileBuy[index];
                                        return Card(
                                          elevation: 10,
                                          margin: const EdgeInsets.symmetric(
                                              vertical: 8, horizontal: 16),
                                          child: ListTile(
                                            leading: const Icon(
                                                Icons.file_present,
                                                size: 40),
                                            title: Text(file['name']!),
                                            trailing: IconButton(
                                              icon: const Icon(Icons.delete,
                                                  color: Colors.red),
                                              onPressed: () => removeFile(
                                                  index, true), // Xóa file
                                            ),
                                          ),
                                        );
                                      },
                                    ),
                                  ),
                                ),
                                const SizedBox(width: 20),
                                Expanded(
                                  child: SizedBox(
                                    height: 400,
                                    child: ListView.builder(
                                      shrinkWrap: true,
                                      itemCount: uploadedFileSale.length,
                                      itemBuilder: (context, index) {
                                        final file = uploadedFileSale[index];
                                        return Card(
                                          elevation: 10,
                                          margin: const EdgeInsets.symmetric(
                                              vertical: 8, horizontal: 16),
                                          child: ListTile(
                                            leading: const Icon(
                                                Icons.file_present,
                                                size: 40),
                                            title: Text(file['name']!),
                                            trailing: IconButton(
                                              icon: const Icon(Icons.delete,
                                                  color: Colors.red),
                                              onPressed: () => removeFile(
                                                  index, false), // Xóa file
                                            ),
                                          ),
                                        );
                                      },
                                    ),
                                  ),
                                )
                              ]),
                          const SizedBox(height: 20),
                          ElevatedButton(
                            onPressed: uploadedFileSale.isNotEmpty ||
                                    uploadedFileBuy.isNotEmpty
                                ? convertToExcel
                                : null, // Vô hiệu hóa nút nếu không có file
                            style: ElevatedButton.styleFrom(
                              padding: const EdgeInsets.symmetric(
                                  vertical: 16, horizontal: 32),
                              textStyle: const TextStyle(fontSize: 18),
                              elevation: 13,
                              backgroundColor: uploadedFileSale.isNotEmpty ||
                                      uploadedFileBuy.isNotEmpty
                                  ? const Color.fromARGB(255, 219, 237,
                                      252) // Màu khi nút được kích hoạt
                                  : Colors.grey, // Màu khi nút bị vô hiệu hóa
                            ),
                            child: const Text('Tạo File tổng hợp Excel'),
                          ),
                          const SizedBox(height: 20),
                        ],
                      ),
                    ),
                  ),
                ],
              ),
            ),
          ),
          // Container(
          //   child: FireworksDisplay(controller: fireworksController),
          // ),
        ]));
  }

  // Hàm chuyển đổi ngày từ "dd/MM/yyyy" thành DateTime
  DateTime parseDate(String date) {
    return DateTime(
      int.parse(date.substring(4, 8)), // yyyy
      int.parse(date.substring(2, 4)), // MM
      int.parse(date.substring(0, 2)), // dd
    );
  }
}
