import 'dart:async';
import 'dart:convert';
import 'package:excel/excel.dart' show Excel;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_fireworks/fireworks_display.dart';
import 'package:html/parser.dart';
import 'package:html_to_excel/common/common_file_list.dart';
import 'package:html_to_excel/common_fireworks.dart';
import 'package:html_to_excel/logic/excel_utils.dart';
import 'package:html_to_excel/logic/safe_value.dart';
import 'package:html_to_excel/logic/zip_file_utils.dart';
import 'logic/code_generator.dart';
import 'package:lottie/lottie.dart';
import 'dart:html' as html; // Dành cho tải file trong Flutter web

class SyntheticPage extends StatefulWidget {
  const SyntheticPage({super.key});

  @override
  _SyntheticPageState createState() => _SyntheticPageState();
}

class _SyntheticPageState extends State<SyntheticPage> {
  Uint8List? barcodePng;
  String? barcodeSvg;
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
  List<List<dynamic>> dataExcel = [];
  List<List<dynamic>> dataBank = [];

  bool isAuxiliaryTemplateEnabled = false;
  Map<String, Object> _htmlContents = {};
  //////////////
  List<Map<String, String>> uploadedFiles = [];

  Map<String, int> invoiceCounterSale = {};
  Map<String, int> invoiceCounterBuy = {};
  final StreamController<double> _progressController =
      StreamController<double>.broadcast();
  Stream<double> get progressStream => _progressController.stream;
  bool isProcessing = false; // Biến kiểm soát trạng thái của nút

  final CommonFireworks commonFireworks = CommonFireworks();
  // Danh sách file theo cấu trúc chung
  List<Map<String, dynamic>> getFileListData() {
    return [
      {
        "title": "Nợ 152",
        "files": uploadedFileBuy1,
        "isBuy": true,
        "debtCode": "No152"
      },
      {
        "title": "Nợ 156",
        "files": uploadedFileBuy2,
        "isBuy": true,
        "debtCode": "No156"
      },
      {
        "title": "Nợ 153",
        "files": uploadedFileBuy3,
        "isBuy": true,
        "debtCode": "No153"
      },
      {
        "title": "Nợ 211",
        "files": uploadedFileBuy4,
        "isBuy": true,
        "debtCode": "No211"
      },
      {
        "title": "Nợ 154",
        "files": uploadedFileBuy5,
        "isBuy": true,
        "debtCode": "No154"
      },
      {
        "title": "Nợ 642",
        "files": uploadedFileBuy6,
        "isBuy": true,
        "debtCode": "No642"
      },
      {
        "title": "Có 156",
        "files": uploadedFileSale1,
        "isBuy": false,
        "debtCode": "Co156"
      },
      {
        "title": "Có 155",
        "files": uploadedFileSale2,
        "isBuy": false,
        "debtCode": "Co155"
      },
      {
        "title": "Có 152",
        "files": uploadedFileSale3,
        "isBuy": false,
        "debtCode": "Co152"
      },
      {
        "title": "File Excel",
        "files": uploadedFiles,
        "isBuy": false,
        "debtCode": "Excel"
      },
    ];
  }

// Hàm tạo danh sách CommonFileList
  List<Widget> buildFileLists() {
    return getFileListData().map((file) {
      return CommonFileList(
        title: file["title"],
        files: file["files"],
        isBuy: file["isBuy"],
        removeFile: removeFile,
        debtCode: file["debtCode"],
      );
    }).toList();
  }

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
      uploadedFiles = [];
      dataBank = [];
      _htmlContents = {}; // Xóa nội dung HTML
    });
  }

  void handlePickZipFiles(bool isBuy, String value) async {
    _progressController.add(0);
    // Sử dụng hàm từ file zip_file_utils.dart
    final files = await ZipFileUtils.pickZipFiles();

    if (files.isNotEmpty) {
      setState(() {
        if (!isBuy) {
          if (value == "Co156") {
            uploadedFileSale1 = files;
          } else if (value == "Co155") {
            uploadedFileSale2 = files;
          } else if (value == "Co152") {
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
              } else if (value == "Co152") {
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
        } else if (value == "Co152") {
          uploadedFileSale3.removeAt(index);
        } else {
          uploadedFiles.removeAt(index);
        }
      }
    });
  }

  void convertToExcel() async {
    if (isProcessing) return; // Chặn nhấp liên tiếp
    setState(() {
      isProcessing = true; // Đánh dấu đang chạy tiến trình
    });

    pickAndProcessExcelFile();
    int totalFiles = _htmlContents.length;
    int processedFiles = 0;

    for (var fileName in _htmlContents.keys) {
      var data = _htmlContents[fileName] as Map<String, dynamic>;

      bool isBuy = (data)['isBuy']; // Lấy giá trị isBuy
      String value = (data)['value']; // Lấy giá trị value
      String content = (data)['content']; // Lấy nội dung hóa đơn
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
          var dataItems =
              document.querySelectorAll('.list-fill-out .data-item');
          ///////////////////////////////////////////////
          var customerName =
              dataItems[0].querySelector('.di-value div')?.text ??
                  ''; // Tên người bán
          var customerCode =
              dataItems[1].querySelector('.di-value div')?.text ??
                  ''; // Mã số thuế người bán
          var customerAddress =
              dataItems[2].querySelector('.di-value div')?.text ??
                  ''; // Địa chỉ người bán
          var customerPhone =
              dataItems[3].querySelector('.di-value div')?.text ??
                  ''; //Số điện thoại người bán
          var buyerCode = dataItems[7].querySelector('.di-value div')?.text ??
              ''; // Mã số thuế người mua
          var customerNameBuy =
              dataItems[5].querySelector('.di-value div')?.text ??
                  ''; // Tên người mua
          var customerAddressBuy =
              dataItems[8].querySelector('.di-value div')?.text ??
                  ''; // Địa chỉ người mua

          String detail;
          if (value == "Co156" || value == "Co155" || value == "Co152") {
            detail = "";
          } else {
            detail = customerCode;
          }

          ///
          bool checkTitle =
              (titleHeading.trim().toUpperCase() == "HOÁ ĐƠN GIÁ TRỊ GIA TĂNG");
          String taxMoney;

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
          List<List<dynamic>> rows = [];
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
              String numberVotes = "CTKT";
              if (value == "No642") {
                numberVotes = "CTKT";
              } else {
                numberVotes =
                    isBuy ? "PN$counterFormatted" : "PX$counterFormatted";
              }

              String debitOut;
              if (isBuy) {
                debitOut = value.substring(2);
              } else {
                if (value == "Co152") {
                  debitOut = "1541";
                } else {
                  debitOut = "632";
                }
              }
              // String note = SafeValueHandler.safeValue(row, 2)
              //     .toString()
              //     .replaceAll(
              //         RegExp(r'[^a-zA-Z0-9À-ỹ ]'), '') // Loại bỏ ký tự đặc biệt
              //     .toUpperCase() // Chuyển thành chữ in hoa
              //     .substring(
              //         0,
              //         min(
              //             95,
              //             SafeValueHandler.safeValue(row, 2)
              //                 .toString()
              //                 .replaceAll(RegExp(r'[^a-zA-Z0-9À-ỹ ]'), '')
              //                 .length));
              String note = SafeValueHandler.safeValue(row, 2).toString();
              String noteDetail;
              if (value == "No152" ||
                  value == "No156" ||
                  value == "No153" ||
                  value == "No211") {
                noteDetail = "Nhập kho chưa thanh toán - $note - $customerName";
              } else if (value == "No154" || value == "No642") {
                noteDetail =
                    "Chi phí mua vật tư, dịch vụ - $note - $customerName";
              } else if (value == "Co156" || value == "Co155") {
                noteDetail = "Xuất kho bán - $note";
              } else if (value == "Co152") {
                noteDetail = "Xuất kho NVL sản xuất - $note";
              } else {
                noteDetail = note;
              }

              return [
                "",
                numberVotes,
                cleanedInvoiceNumber, // Ghi chú
                dayElement, // Ngày
                noteDetail,
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
                debitOut,
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
                note,
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
              String numberVotes = "CTKT";
              if (value == "No642") {
                numberVotes = "CTKT";
              } else {
                numberVotes =
                    isBuy ? "PN$counterFormatted" : "PX$counterFormatted";
              }
              String debitOut;
              if (isBuy) {
                debitOut = value.substring(2);
              } else {
                if (value == "Co152") {
                  debitOut = "1541";
                } else {
                  debitOut = "632";
                }
              }
              // String note = SafeValueHandler.safeValue(row, 2)
              //     .toString()
              //     .replaceAll(
              //         RegExp(r'[^a-zA-Z0-9À-ỹ ]'), '') // Loại bỏ ký tự đặc biệt
              //     .toUpperCase() // Chuyển thành chữ in hoa
              //     .substring(
              //         0,
              //         min(
              //             95,
              //             SafeValueHandler.safeValue(row, 2)
              //                 .toString()
              //                 .replaceAll(RegExp(r'[^a-zA-Z0-9À-ỹ ]'), '')
              //                 .length));
              String note = SafeValueHandler.safeValue(row, 2).toString();
              String noteDetail;
              if (value == "No152" ||
                  value == "No156" ||
                  value == "No153" ||
                  value == "No211") {
                noteDetail = "Nhập kho chưa thanh toán - $note - $customerName";
              } else if (value == "No154" || value == "No642") {
                noteDetail =
                    "Chi phí mua vật tư, dịch vụ - $note - $customerName";
              } else if (value == "Co156" || value == "Co155") {
                noteDetail = "Xuất kho bán - $note";
              } else if (value == "Co152") {
                noteDetail = "Xuất kho NVL sản xuất - $note";
              } else {
                noteDetail = note;
              }

              // Tạo một danh sách mới với chuỗi rỗng ở các vị trí cụ thể
              return [
                "",
                numberVotes,
                cleanedInvoiceNumber, // Ghi chú
                dayElement, // Ngày
                noteDetail,
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
                debitOut,
                "",
                isBuy ? "331" : value.substring(2),
                detail,

                ///////////////////////////////
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
                note,
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
          List<List<dynamic>> listPercentBuy = [
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
          String data511;
          String nameDetail;
          if (value == "Co156") {
            data511 = "5111";
            nameDetail = "Doanh thu bán hàng hoá - $customerNameBuy";
          } else if (value == "Co155") {
            data511 = "5112";
            nameDetail = "Doanh thu bán thành phẩm - $customerNameBuy";
          } else if (value == "Co152") {
            data511 = "5113";
            nameDetail = "Doanh thu cung cấp dịch vụ - $customerNameBuy";
          } else {
            data511 = "";
            nameDetail = customerNameBuy;
          }
          List<List<dynamic>> listPercent = [
            [
              "",
              "",
              cleanedInvoiceNumber,
              dayElement,
              nameDetail,
              "",
              "",
              "",
              "",
              sumMoney.replaceAll(".", ""),
              "131",
              buyerCode,
              data511,
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
      processedFiles++;
      double progress = (processedFiles / totalFiles) * 100;
      await Future.delayed(const Duration(milliseconds: 5), () {
        _progressController.add(progress); // Cập nhật tiến trình
      });
    }
// Thêm dữ liệu ngân hàng
    dataExcel.addAll(dataBank);

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
      // row[9] = formatMoney(row[9]);
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

  String getInvoiceCounter(Map<String, int> invoiceCounter, String key) {
    if (!invoiceCounter.containsKey(key)) {
      invoiceCounter[key] = invoiceCounter.length + 1;
    }
    return invoiceCounter[key]!.toString().padLeft(3, '0');
  }

  void _pickExcelFile() async {
    _progressController.add(0);
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

  void pickAndProcessExcelFile() async {
    if (uploadedFiles.isEmpty) {
      print("Không có file nào được chọn!");
      return;
    }

    try {
      // Lấy bytes từ uploadedFiles
      String base64String = uploadedFiles.first['bytes']!;
      var bytes = base64Decode(base64String);
      var excel = Excel.decodeBytes(bytes);

      // Lấy sheet đầu tiên
      var sheetName = excel.tables.keys.first;
      var table = excel.tables[sheetName];

      if (table == null || table.rows.length <= 2) {
        print("Lỗi: File không có đủ dữ liệu!");
        return;
      }

      print("3 - Đang xử lý dữ liệu");

      // Xử lý từng dòng dữ liệu (bắt đầu từ dòng thứ 3)
      for (var i = 1; i < table.rows.length; i++) {
        if (table.rows[i].isEmpty) {
          print("Bỏ qua dòng $i vì không có dữ liệu.");
          continue;
        }

        // Chuyển đổi tất cả giá trị trong hàng thành String
        List<dynamic> rowData =
            table.rows[i].map((cell) => cell?.value?.toString() ?? "").toList();

        // Đảm bảo danh sách có ít nhất 3 cột trước khi truy xuất
        String column1 = rowData.length > 0 ? rowData[0] : "N/A";
        String column2 = rowData.length > 1 ? rowData[1] : "N/A";
        String column3 = rowData.length > 2 ? rowData[2] : "N/A";
        String column4 = rowData.length > 3 ? rowData[3] : "N/A";
        String column5 = rowData.length > 4 ? rowData[4] : "N/A";
        String column7 = rowData.length > 6 ? rowData[6] : "N/A";
        String money;
        bool isDebit;
        if (column3 == "0") {
          isDebit = false;
          money = column4;
        } else {
          isDebit = true;
          money = column3;
        }
        var dayElement = column1;
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

        List<dynamic> sortedRow = [
          "",
          "",
          column2,
          dayElement,
          column5,
          "",
          "",
          "",
          "",
          money,
          isDebit ? "" : "1121",
          isDebit ? "" : column7,
          isDebit ? "1121" : "",
          isDebit ? column7 : "",
        ];

        // Thêm vào dataBank
        dataBank.add(sortedRow);
      }
      print("4 - Dữ liệu đã xử lý xong, chuẩn bị ghi file.");
    } catch (e) {
      print("Lỗi xử lý file: $e");
    }
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
                crossAxisAlignment: CrossAxisAlignment.center,
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  Expanded(
                    child: Center(
                      child: Column(
                        mainAxisAlignment: MainAxisAlignment.spaceBetween,
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
                                      child: Text('Nhập kho nguyên vật liệu',
                                          style: TextStyle(
                                              fontSize: 15,
                                              fontWeight: FontWeight.bold)),
                                    ),
                                    const PopupMenuDivider(),
                                    const PopupMenuItem(
                                      value: "No156",
                                      child: Text('Nhập kho hàng hóa',
                                          style: TextStyle(
                                              fontSize: 15,
                                              fontWeight: FontWeight.bold)),
                                    ),
                                    const PopupMenuDivider(),
                                    const PopupMenuItem(
                                      value: "No153",
                                      child: Text('Nhâp kho CCDC',
                                          style: TextStyle(
                                              fontSize: 15,
                                              fontWeight: FontWeight.bold)),
                                    ),
                                    const PopupMenuDivider(),
                                    const PopupMenuItem(
                                      value: "No211",
                                      child: Text('Nhâp kho TSCĐ',
                                          style: TextStyle(
                                              fontSize: 15,
                                              fontWeight: FontWeight.bold)),
                                    ),
                                    const PopupMenuDivider(),
                                    const PopupMenuItem(
                                      value: "No154",
                                      child: Text('Chi phí SXDD',
                                          style: TextStyle(
                                              fontSize: 15,
                                              fontWeight: FontWeight.bold)),
                                    ),
                                    const PopupMenuDivider(),
                                    const PopupMenuItem(
                                      value: "No642",
                                      child: Text('Chi phí QLKD',
                                          style: TextStyle(
                                              fontSize: 15,
                                              fontWeight: FontWeight.bold)),
                                    ),
                                  ],
                                  shape: RoundedRectangleBorder(
                                    borderRadius:
                                        BorderRadius.circular(12), // Bo góc
                                    side: const BorderSide(
                                        color:
                                            Color.fromARGB(255, 155, 156, 147),
                                        width: 2), // Đường viền xanh
                                  ),
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
                                      child: Text('Doanh thu bán hàng hoá',
                                          style: TextStyle(
                                              fontSize: 15,
                                              fontWeight: FontWeight.bold)),
                                    ),
                                    const PopupMenuDivider(),
                                    const PopupMenuItem(
                                      value: "Co155",
                                      child: Text('Doanh thu bán thành phẩm',
                                          style: TextStyle(
                                              fontSize: 15,
                                              fontWeight: FontWeight.bold)),
                                    ),
                                    const PopupMenuDivider(),
                                    const PopupMenuItem(
                                      value: "Co152",
                                      child: Text('Doanh thu cung cấp dịch vụ',
                                          style: TextStyle(
                                              fontSize: 15,
                                              fontWeight: FontWeight.bold)),
                                    ),
                                  ],
                                  shape: RoundedRectangleBorder(
                                    borderRadius:
                                        BorderRadius.circular(12), // Bo góc
                                    side: const BorderSide(
                                        color:
                                            Color.fromARGB(255, 155, 156, 147),
                                        width: 2), // Đường viền xanh
                                  ),
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
                                          .fromARGB(255, 242, 138,
                                          144), // Đảm bảo màu nền không thay đổi
                                      disabledForegroundColor:
                                          const Color.fromARGB(255, 40, 64, 99),
                                    ).copyWith(
                                      elevation:
                                          WidgetStateProperty.all<double>(12),
                                      shadowColor:
                                          WidgetStateProperty.all<Color>(
                                              const Color.fromARGB(
                                                  255, 187, 44, 106)),
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
                                  color: Color.fromARGB(255, 155, 156, 147),
                                  width: 2), // Viền đen bao cả khối
                              borderRadius: BorderRadius.circular(
                                  10), // Bo góc (tùy chọn)
                            ),
                            child: Row(
                                mainAxisAlignment: MainAxisAlignment.center,
                                children: buildFileLists()),
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
                                    uploadedFileBuy6.isNotEmpty ||
                                    uploadedFiles.isNotEmpty
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
                                      uploadedFileBuy6.isNotEmpty ||
                                      uploadedFiles.isNotEmpty
                                  ? const Color.fromARGB(255, 219, 237,
                                      252) // Màu khi nút được kích hoạt
                                  : Colors.grey, // Màu khi nút bị vô hiệu hóa
                            ),
                            child: const Text('Tạo File tổng hợp Excel'),
                          ),
                          const SizedBox(height: 50),
                          StreamBuilder<double>(
                            stream: progressStream,
                            builder: (context, snapshot) {
                              double progress = (snapshot.data ?? 0) / 100;
                              double maxWidth =
                                  MediaQuery.of(context).size.width /
                                      2; // Chiều dài thanh tiến trình
                              return SizedBox(
                                width: maxWidth,
                                child: Padding(
                                  padding: const EdgeInsets.all(0),
                                  child: Column(
                                    children: [
                                      Stack(
                                        children: [
                                          // Thanh tiến trình bo tròn
                                          ClipRRect(
                                            borderRadius:
                                                BorderRadius.circular(10),
                                            child: SizedBox(
                                              height: 14,
                                              child: LinearProgressIndicator(
                                                value: progress,
                                                minHeight: 14,
                                                backgroundColor:
                                                    Colors.grey[300],
                                                valueColor:
                                                    const AlwaysStoppedAnimation<
                                                        Color>(Colors.blue),
                                              ),
                                            ),
                                          ),
                                          // Icon con chó chạy trên thanh tiến trình
                                          AnimatedAlign(
                                            duration: const Duration(
                                                milliseconds:
                                                    2), // Cập nhật ngay khi progress thay đổi
                                            alignment: Alignment(
                                                (progress * 2) -
                                                    1, // Chuyển progress thành vị trí trên thanh
                                                0), // Giữ icon nằm trên thanh tiến trình
                                            child: Transform.translate(
                                              offset: const Offset(0,
                                                  -50), // Dịch chuyển icon lên trên 12 pixel
                                              child: Lottie.asset(
                                                'assets/run.json', // Animation từ ảnh cá nhân của bạn
                                                width: 50,
                                                height: 50,
                                                repeat: true,
                                              ),
                                            ),
                                          ),
                                        ],
                                      ),
                                      Text(
                                        '${(progress * 100).toStringAsFixed(1)}%',
                                        style: const TextStyle(
                                          fontSize: 15, // Tăng kích thước chữ
                                          fontWeight: FontWeight
                                              .bold, // Làm chữ đậm hơn
                                          color:
                                              Colors.red, // Đổi màu chữ sang đỏ
                                        ),
                                      ),
                                    ],
                                  ),
                                ),
                              );
                            },
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
            child: FireworksDisplay(
                controller: commonFireworks.fireworksController),
          ),
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

  String formatMoney(String money) {
    int moneyInt = int.tryParse(money) ?? 0;
    return moneyInt.toString().replaceAllMapped(
        RegExp(r'(\d)(?=(\d{3})+(?!\d))'), (Match m) => '${m[1]}.');
  }
}
