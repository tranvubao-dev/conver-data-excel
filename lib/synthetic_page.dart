import 'dart:convert';
import 'dart:math';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:html/parser.dart';
import 'package:html_to_excel/common/common_file_list.dart';
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
  // List<Map<String, dynamic>> uploadedFileSale =
  //     []; // List để lưu các file hoá đơn bán hàng
  List<Map<String, dynamic>> uploadedFileSale1 = [];
  List<Map<String, dynamic>> uploadedFileSale2 = [];
  List<Map<String, dynamic>> uploadedFileSale3 = [];
  List<Map<String, dynamic>> uploadedFileBuy1 = [];
  List<Map<String, dynamic>> uploadedFileBuy2 = [];
  List<Map<String, dynamic>> uploadedFileBuy3 = [];
  List<Map<String, dynamic>> uploadedFileBuy4 = [];
  List<Map<String, dynamic>> uploadedFileBuy5 = [];
  List<Map<String, dynamic>> uploadedFileBuy6 = [];
  List<List<String>> dataExcel = [];
  String? _fileName;

  bool isAuxiliaryTemplateEnabled = false;
  Map<String, Object> _htmlContents = {};
  //////////////
  List<Map<String, String>> uploadedFiles = [];

  Map<String, int> invoiceCounterSale = {};
  Map<String, int> invoiceCounterBuy = {};

  void resetState() {
    setState(() {
      dataExcel = []; // Đặt lại thanh tiêu đề
      uploadedFileSale1 = []; // Xóa danh sách file
      uploadedFileSale2 = [];
      uploadedFileSale3 = [];
      uploadedFileBuy1 = []; // Xóa danh sách file
      uploadedFileBuy2 = [];
      uploadedFileBuy3 = [];
      uploadedFileBuy4 = [];
      uploadedFileBuy5 = [];
      uploadedFileBuy6 = [];
      _htmlContents = {}; // Xóa nội dung HTML
    });
  }

  void handlePickZipFiles(bool isBuy, String value) async {
    // Sử dụng hàm từ file zip_file_utils.dart
    final files = await ZipFileUtils.pickZipFiles();

    if (files.isNotEmpty) {
      setState(() {
        if (!isBuy) {
          if (value == "Co156") {
            uploadedFileSale1 = files;
          } else if (value == "Co155") {
            uploadedFileSale2 = files;
          } else if (value == "Co154") {
            uploadedFileSale3 = files;
          }
        } else {
          if (value == "No152") {
            uploadedFileBuy1 = files;
          } else if (value == "No156") {
            uploadedFileBuy2 = files;
          } else if (value == "No153") {
            uploadedFileBuy3 = files;
          } else if (value == "No211") {
            uploadedFileBuy4 = files;
          } else if (value == "No154") {
            uploadedFileBuy5 = files;
          } else if (value == "No642") {
            uploadedFileBuy6 = files;
          }
        }
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
                'value': value,
              };
            });
          } else {
            ScaffoldMessenger.of(context).showSnackBar(
              SnackBar(
                  content: Text(
                      'Không tìm thấy file invoice.html trong: ${file['name']}')),
            );
            int index = 0;
            if (!isBuy) {
              if (value == "Co156") {
                index = uploadedFileSale1
                    .indexWhere((f) => f['name'] == file['name']);
              } else if (value == "Co155") {
                index = uploadedFileSale2
                    .indexWhere((f) => f['name'] == file['name']);
              } else if (value == "Co154") {
                index = uploadedFileSale3
                    .indexWhere((f) => f['name'] == file['name']);
              }
            } else {
              if (value == "No152") {
                index = uploadedFileBuy1
                    .indexWhere((f) => f['name'] == file['name']);
              } else if (value == "No156") {
                index = uploadedFileBuy2
                    .indexWhere((f) => f['name'] == file['name']);
              } else if (value == "No153") {
                index = uploadedFileBuy3
                    .indexWhere((f) => f['name'] == file['name']);
              } else if (value == "No211") {
                index = uploadedFileBuy4
                    .indexWhere((f) => f['name'] == file['name']);
              } else if (value == "No154") {
                index = uploadedFileBuy5
                    .indexWhere((f) => f['name'] == file['name']);
              } else if (value == "No642") {
                index = uploadedFileBuy6
                    .indexWhere((f) => f['name'] == file['name']);
              }
            }
            removeFile(index, isBuy, value);
          }
        }
      }
    }
  }

  void removeFile(int index, bool isBuy, String value) {
    setState(() {
      if (isBuy) {
        if (value == "No152") {
          uploadedFileBuy1.removeAt(index);
        } else if (value == "No156") {
          uploadedFileBuy2.removeAt(index);
        } else if (value == "No153") {
          uploadedFileBuy3.removeAt(index);
        } else if (value == "No211") {
          uploadedFileBuy4.removeAt(index);
        } else if (value == "No154") {
          uploadedFileBuy5.removeAt(index);
        } else if (value == "No642") {
          uploadedFileBuy6.removeAt(index);
        }
      } else {
        if (value == "Co156") {
          uploadedFileSale1.removeAt(index);
        } else if (value == "Co155") {
          uploadedFileSale2.removeAt(index);
        } else if (value == "Co154") {
          uploadedFileSale3.removeAt(index);
        } else {
          uploadedFiles.removeAt(index);
        }
      }
    });
  }

  void convertToExcel() {
    // dataExcel.addAll(listTitle);
    _htmlContents.forEach((fileName, data) {
      bool isBuy = (data as Map<String, dynamic>)['isBuy']; // Lấy giá trị isBuy
      String value =
          (data as Map<String, dynamic>)['value']; // Lấy giá trị value
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

          String detail;
          if (value == "Co156" || value == "Co155" || value == "Co154") {
            detail = "";
          } else {
            detail = customerCode;
          }
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
          var customerNameBuy = document // Tên người mua
              .querySelectorAll('.list-fill-out .data-item')[5]
              .querySelector('.di-value div')!
              .text;
          ///////////////////////////////////////////////
          var customerAddressBuy = document // Địa chỉ người mua
              .querySelectorAll('.list-fill-out .data-item')[8]
              .querySelector('.di-value div')!
              .text;

          ///
          bool checkTitle =
              (titleHeading.trim().toUpperCase() == "HOÁ ĐƠN GIÁ TRỊ GIA TĂNG");
          var taxMoney;

          if (checkTitle) {
            var titletaxMoney = document
                .querySelectorAll('.table-horizontal-wrapper .res-tb')[
                    1] // Chọn bảng thứ 2 trong .table-horizontal-wrapper
                .querySelectorAll(
                    'tbody tr')[1] // Chọn dòng thứ 2 (dòng Tổng tiền thuế)
                .querySelectorAll('td')[0] // Chọn cột thứ 2 (chứa số 37.840)
                .text
                .trim();
            if (titletaxMoney == "Tổng tiền thuế (Tổng cộng tiền thuế)") {
              taxMoney = document
                  .querySelectorAll('.table-horizontal-wrapper .res-tb')[
                      1] // Chọn bảng thứ 2 trong .table-horizontal-wrapper
                  .querySelectorAll(
                      'tbody tr')[1] // Chọn dòng thứ 2 (dòng Tổng tiền thuế)
                  .querySelectorAll('td')[1] // Chọn cột thứ 2 (chứa số 37.840)
                  .text
                  .trim();
            } else {
              taxMoney = document
                  .querySelectorAll('.table-horizontal-wrapper .res-tb')[
                      1] // Chọn bảng thứ 2 trong .table-horizontal-wrapper
                  .querySelectorAll(
                      'tbody tr')[2] // Chọn dòng thứ 2 (dòng Tổng tiền thuế)
                  .querySelectorAll('td')[1] // Chọn cột thứ 2 (chứa số 37.840)
                  .text
                  .trim();
            }
          } else {
            taxMoney = document
                .querySelectorAll('.table-horizontal-wrapper .res-tb')[0]
                .querySelectorAll('tbody tr')[1]
                .querySelectorAll('td')[1]
                .text
                .trim();
          }

          ///////////// tổng tiền chưa thuế
          var sumMoney = checkTitle
              ? document
                  .querySelectorAll('.table-horizontal-wrapper .res-tb')[
                      1] // Chọn bảng thứ 2 trong .table-horizontal-wrapper
                  .querySelectorAll('tbody tr')[0] //
                  .querySelectorAll('td')[1]
                  .text
                  .trim()
              : document
                  .querySelectorAll('.table-horizontal-wrapper .res-tb')[0]
                  .querySelectorAll('tbody tr')[2] //
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
              String key = "$cleanedInvoiceNumber - $dayElement";
              String counterFormatted = isBuy
                  ? getInvoiceCounter(invoiceCounterBuy, key)
                  : getInvoiceCounter(invoiceCounterSale, key);
              return [
                "",
                isBuy ? "PN$counterFormatted" : "PX$counterFormatted",
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
                isBuy ? value.substring(2) : "632",
                "",
                isBuy ? "331" : value.substring(2),
                detail,
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
                customerPhone, // Số điện thoại người bán
                customerNameBuy, // Tên người mua
                customerAddressBuy, // Địa chỉ người mua
              ];
            }).toList();
          } else {
            ///HOÁ ĐƠN BÁN HÀNG
            dataExcelDetail = rows.map((row) {
              String key = "$cleanedInvoiceNumber - $dayElement";
              String counterFormatted = isBuy
                  ? getInvoiceCounter(invoiceCounterBuy, key)
                  : getInvoiceCounter(invoiceCounterSale, key);
              // Tạo một danh sách mới với chuỗi rỗng ở các vị trí cụ thể
              return [
                "",
                isBuy ? "PN$counterFormatted" : "PX$counterFormatted",
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
                isBuy ? value.substring(2) : "632",
                "",
                isBuy ? "331" : value.substring(2),
                detail,

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
                customerPhone, // Số điện thoại người bán
                customerNameBuy, // Tên người mua
                customerAddressBuy, // Địa chỉ người mua
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
              isBuy ? "331" : "3331",
              isBuy ? customerCode : "",
            ], // Thanh tieu de
          ];

          List<List<String>> listPercent = [
            [
              "",
              "",
              cleanedInvoiceNumber,
              dayElement,
              customerNameBuy,
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

    Map<String, List<List<String>>> groupedData = {};
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
    List<List<String>> sortedDataDay = [];
    for (var entry in sortedEntries) {
      sortedDataDay.addAll(
          entry.value); // Thêm tất cả dòng của nhóm vào danh sách kết quả
    }

    List<List<String>> convertedData = sortedDataDay.map((row) {
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
    // firework();

    /// Pháo hoa
  }

  String getInvoiceCounter(Map<String, int> invoiceCounter, String key) {
    if (!invoiceCounter.containsKey(key)) {
      invoiceCounter[key] = invoiceCounter.length + 1;
    }
    return invoiceCounter[key]!.toString().padLeft(3, '0');
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
              padding: const EdgeInsets.all(5.0),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Expanded(
                    child: Center(
                      child: Column(
                        children: [
                          Row(
                              mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                              children: [
                                PopupMenuButton<String>(
                                  onSelected: (String value) {
                                    handlePickZipFiles(
                                        true, value); // Gọi hàm xử lý tải file
                                  },
                                  itemBuilder: (BuildContext context) => [
                                    const PopupMenuItem(
                                      value: "No152",
                                      child: Text('Nhập kho nguyên vật liệu'),
                                    ),
                                    const PopupMenuItem(
                                      value: "No156",
                                      child: Text('Nhập kho hàng hóa'),
                                    ),
                                    const PopupMenuItem(
                                      value: "No153",
                                      child: Text('Nhâp kho CCDC'),
                                    ),
                                    const PopupMenuItem(
                                      value: "No211",
                                      child: Text('Nhâp kho TSCĐ'),
                                    ),
                                    const PopupMenuItem(
                                      value: "No154",
                                      child: Text('Chi phí SXDD'),
                                    ),
                                    const PopupMenuItem(
                                      value: "No642",
                                      child: Text('Chi phí QLKD'),
                                    ),
                                  ],
                                  child: ElevatedButton.icon(
                                    onPressed: null,
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
                                          255,
                                          226,
                                          235,
                                          250), // Màu bóng đổ// Giữ màu xanh ngay cả khi bị vô hiệu hóa
                                      foregroundColor:
                                          Colors.white, // Giữ màu chữ trắng
                                      disabledBackgroundColor: const Color
                                          .fromARGB(255, 219, 237,
                                          252), // Đảm bảo màu nền không thay đổi
                                      disabledForegroundColor:
                                          const Color.fromARGB(255, 40, 64, 99),
                                    ).copyWith(
                                      elevation:
                                          WidgetStateProperty.all<double>(
                                              12), // Tăng độ cao khi nhấn
                                      shadowColor:
                                          WidgetStateProperty.all<Color>(
                                              Colors.blue[800]!),
                                    ),
                                    icon: const Icon(Icons.upload_file),
                                    label: const Text(
                                        'Tải lên file .zip Hoá Đơn Mua Hàng'),
                                  ),
                                ),
                                // const VerticalDivider(thickness: 2, width: 80),
                                PopupMenuButton<String>(
                                  onSelected: (String value) {
                                    handlePickZipFiles(
                                        false, value); // Gọi hàm xử lý tải file
                                  },
                                  itemBuilder: (BuildContext context) => [
                                    const PopupMenuItem(
                                      value: "Co156",
                                      child: Text('Doanh thu bán hàng hoá'),
                                    ),
                                    const PopupMenuItem(
                                      value: "Co155",
                                      child: Text('Doanh thu bán thành phẩm'),
                                    ),
                                    const PopupMenuItem(
                                      value: "Co154",
                                      child: Text('Doanh thu cung cấp dịch vụ'),
                                    ),
                                  ],
                                  child: ElevatedButton.icon(
                                    onPressed: null,
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
                                      foregroundColor:
                                          Colors.white, // Giữ màu chữ trắng
                                      disabledBackgroundColor: const Color
                                          .fromARGB(255, 240, 128,
                                          133), // Đảm bảo màu nền không thay đổi
                                      disabledForegroundColor:
                                          const Color.fromARGB(255, 40, 64, 99),
                                    ).copyWith(
                                      elevation:
                                          WidgetStateProperty.all<double>(12),
                                      shadowColor:
                                          WidgetStateProperty.all<Color>(
                                              Colors.pink[800]!),
                                    ),
                                    icon: const Icon(Icons.upload_file),
                                    label: const Text(
                                        'Tải lên file .zip Hoá Đơn Bán Hàng'),
                                  ),
                                ),
                                ElevatedButton.icon(
                                  onPressed: _pickExcelFile,
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
                                    shadowColor: const Color.fromARGB(
                                        255, 226, 235, 250),
                                  ).copyWith(
                                    elevation:
                                        WidgetStateProperty.all<double>(12),
                                    shadowColor: WidgetStateProperty.all<Color>(
                                        Colors.blue[800]!),
                                  ),
                                  icon: const Icon(Icons.upload_file),
                                  label: const Text(
                                      'Tải lên file Excel Ngân Hàng'),
                                ),
                              ]),
                          const SizedBox(height: 30),
                          Container(
                            decoration: BoxDecoration(
                              border: Border.all(
                                  color: Colors.black,
                                  width: 2), // Viền đen bao cả khối
                              borderRadius: BorderRadius.circular(
                                  10), // Bo góc (tùy chọn)
                            ),
                            child: Row(
                                mainAxisAlignment: MainAxisAlignment.center,
                                children: [
                                  if (_fileName != null)
                                    const SizedBox(height: 20),
                                  CommonFileList(
                                    title: "Nợ 152",
                                    files: uploadedFileBuy1,
                                    isBuy: true,
                                    removeFile: removeFile,
                                    debtCode: "No152",
                                  ),
                                  CommonFileList(
                                    title: "Nợ 156",
                                    files: uploadedFileBuy2,
                                    isBuy: true,
                                    removeFile: removeFile,
                                    debtCode: "No156",
                                  ),
                                  CommonFileList(
                                    title: "Nợ 153",
                                    files: uploadedFileBuy3,
                                    isBuy: true,
                                    removeFile: removeFile,
                                    debtCode: "No153",
                                  ),
                                  CommonFileList(
                                    title: "Nợ 211",
                                    files: uploadedFileBuy4,
                                    isBuy: true,
                                    removeFile: removeFile,
                                    debtCode: "No211",
                                  ),
                                  CommonFileList(
                                    title: "Nợ 154",
                                    files: uploadedFileBuy5,
                                    isBuy: true,
                                    removeFile: removeFile,
                                    debtCode: "No154",
                                  ),
                                  CommonFileList(
                                    title: "Nợ 642",
                                    files: uploadedFileBuy6,
                                    isBuy: true,
                                    removeFile: removeFile,
                                    debtCode: "No642",
                                  ),
                                  CommonFileList(
                                    title: "Có 156",
                                    files: uploadedFileSale1,
                                    isBuy: false,
                                    removeFile: removeFile,
                                    debtCode: "Co156",
                                  ),
                                  CommonFileList(
                                    title: "Có 155",
                                    files: uploadedFileSale2,
                                    isBuy: false,
                                    removeFile: removeFile,
                                    debtCode: "Co155",
                                  ),
                                  CommonFileList(
                                    title: "Có 154",
                                    files: uploadedFileSale3,
                                    isBuy: false,
                                    removeFile: removeFile,
                                    debtCode: "Co154",
                                  ),
                                  CommonFileList(
                                    title: "File Excel",
                                    files: uploadedFiles,
                                    isBuy: false,
                                    removeFile: removeFile,
                                    debtCode: "Excel",
                                  ),
                                ]),
                          ),
                          const SizedBox(height: 20),
                          ElevatedButton(
                            onPressed: uploadedFileSale1.isNotEmpty ||
                                    uploadedFileSale2.isNotEmpty ||
                                    uploadedFileSale3.isNotEmpty ||
                                    uploadedFileBuy1.isNotEmpty ||
                                    uploadedFileBuy2.isNotEmpty ||
                                    uploadedFileBuy3.isNotEmpty ||
                                    uploadedFileBuy4.isNotEmpty ||
                                    uploadedFileBuy5.isNotEmpty ||
                                    uploadedFileBuy6.isNotEmpty
                                ? convertToExcel
                                : null, // Vô hiệu hóa nút nếu không có file
                            style: ElevatedButton.styleFrom(
                              padding: const EdgeInsets.symmetric(
                                  vertical: 16, horizontal: 32),
                              textStyle: const TextStyle(fontSize: 18),
                              elevation: 13,
                              backgroundColor: uploadedFileSale1.isNotEmpty ||
                                      uploadedFileSale2.isNotEmpty ||
                                      uploadedFileSale3.isNotEmpty ||
                                      uploadedFileBuy1.isNotEmpty ||
                                      uploadedFileBuy2.isNotEmpty ||
                                      uploadedFileBuy3.isNotEmpty ||
                                      uploadedFileBuy4.isNotEmpty ||
                                      uploadedFileBuy5.isNotEmpty ||
                                      uploadedFileBuy6.isNotEmpty
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
