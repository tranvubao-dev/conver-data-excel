import 'dart:convert';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as excel;
import 'dart:html' as html;

class ExcelUtils {
  /// Tạo và tải xuống file Excel
  static void createExcelFile(List<List<String>> dataExcel, String fileName,
      bool createManySheets, bool isShopping) {
    // Tạo một workbook
    final excel.Workbook workbook = excel.Workbook();

    // Accessing worksheet via index.
    final excel.Worksheet sheet = workbook.worksheets[0];
    sheet.name = fileName; // Đổi tên sheet

    // Thêm dữ liệu vào sheet
    for (int rowIndex = 0; rowIndex < dataExcel.length; rowIndex++) {
      for (int colIndex = 0;
          colIndex < dataExcel[rowIndex].length - 2;
          colIndex++) {
        final cell = sheet.getRangeByIndex(rowIndex + 1, colIndex + 1);
        cell.setText(dataExcel[rowIndex][colIndex]);

        // Thêm bo viền cho ô
        cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;

        if (rowIndex == 0) {
          cell.cellStyle.bold = true; // Đặt chữ in đậm
          cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
        } else {
          cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
        }
      }
    }
    // Thêm dữ liệu vào sheet hàng hoá
    List<List<String>> transformedData = [];
    Set<String> seenCodeMH = {}; // Set lưu các mã mặt hàng đã xuất hiện
    for (int i = 0; i < dataExcel.length; i++) {
      List<String> row = dataExcel[i];
      // for (var row in dataExcel) {
      if (row.length >= 4) {
        // Lấy mã khách hàng và tên khách hàng
        String codeMH; // Mã mặt hàng
        String nameMH; // Tên mặt hàng
        String unit; // Đơn vị tính
        String price; // Đơn giá
        if (isShopping) {
          codeMH = row[9]; // Mã mặt hàng
          nameMH = row[10]; // Tên mặt hàng
          unit = row[11]; // Đơn vị tính
          price = row[13]; // Đơn giá
        } else {
          codeMH = row[11]; // Mã mặt hàng
          nameMH = row[12]; // Tên mặt hàng
          unit = row[13]; // Đơn vị tính
          price = row[15]; // Đơn giá
        }

        // Nếu không phải hàng header và codeMH đã tồn tại thì bỏ qua
        if (i != 0 && seenCodeMH.contains(codeMH)) {
          continue;
        }
        // Thêm codeMH vào set để kiểm tra cho các hàng sau
        seenCodeMH.add(codeMH);

        List<String> newRow;
        if (i == 0) {
          newRow = [
            codeMH,
            nameMH,
            "Loại quy cách",
            "Thông số đặc tả",
            unit,
            "Danh mục mặt hàng",
            "Hàng đóng kiện",
            "Quản lý số lượng",
            "Quy trình",
            price,
            "Giá mua Tình trạng thuế",
            "Giá bán",
            "Giá bán Tình trạng thuế"
          ];
        } else {
          newRow = [
            codeMH,
            nameMH,
            "",
            "",
            unit,
            "3",
            "",
            "",
            "",
            price,
            "",
            "",
            ""
          ];
        }

        transformedData.add(newRow);
      }
    }

    // Thêm dữ liệu vào sheet nhà cung cấp
    List<List<String>> customerData = [];
    Set<String> seenCodeCustomer = {}; // Set lưu các mã mặt hàng đã xuất hiện
    for (int i = 0; i < dataExcel.length; i++) {
      List<String> row = dataExcel[i];
      // for (var row in dataExcel) {
      if (row.length >= 4) {
        // Lấy mã khách hàng và tên khách hàng
        String codeCustomer = row[2]; // Mã KH/ NCC
        String nameCustomer = row[3]; // Tên công ty
        String person = row[4]; // Người phụ trách
        String customerAddress; // Địa chỉ công ty
        String customerPhone; // Điện thoại
        if (isShopping) {
          customerAddress = row[22]; // Địa chỉ công ty
          customerPhone = row[23]; // Điện thoại
        } else {
          customerAddress = row[26]; // Địa chỉ công ty
          customerPhone = row[27]; // Điện thoại
        }

        // Nếu không phải hàng header và codeMH đã tồn tại thì bỏ qua
        if (i != 0 && seenCodeCustomer.contains(codeCustomer)) {
          continue;
        }
        // Thêm codeMH vào set để kiểm tra cho các hàng sau
        seenCodeCustomer.add(codeCustomer);

        List<String> newRow;
        if (i == 0) {
          newRow = [
            codeCustomer,
            nameCustomer,
            "Điện thoại",
            "Mã 1",
            "Địa chỉ công ty",
            "Email",
            person,
            "Mã số thuế"
          ];
        } else {
          newRow = [
            codeCustomer,
            nameCustomer,
            customerPhone,
            "",
            customerAddress,
            "",
            person,
            codeCustomer
          ];
        }

        customerData.add(newRow);
      }
    }

    if (createManySheets) {
      final excel.Worksheet sheet2 = workbook.worksheets.add();
      sheet2.name = "Mặt hàng"; // Đặt tên sheet phụ
      // Thêm dữ liệu vào sheet
      for (int rowIndex = 0; rowIndex < transformedData.length; rowIndex++) {
        for (int colIndex = 0;
            colIndex < transformedData[rowIndex].length;
            colIndex++) {
          final cell = sheet2.getRangeByIndex(rowIndex + 1, colIndex + 1);
          cell.setText(transformedData[rowIndex][colIndex]);
          // Thêm bo viền cho ô
          cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;
          if (rowIndex == 0) {
            cell.cellStyle.bold = true; // Đặt chữ in đậm
            cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
          } else {
            cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
          }
        }
      }

      final excel.Worksheet sheet3 = workbook.worksheets.add();
      sheet3.name = "Khách hàng"; // Đặt tên sheet phụ

      // Thêm dữ liệu vào sheet
      for (int rowIndex = 0; rowIndex < customerData.length; rowIndex++) {
        for (int colIndex = 0;
            colIndex < customerData[rowIndex].length;
            colIndex++) {
          final cell = sheet3.getRangeByIndex(rowIndex + 1, colIndex + 1);
          cell.setText(customerData[rowIndex][colIndex]);

          // Thêm bo viền cho ô
          cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;

          if (rowIndex == 0) {
            cell.cellStyle.bold = true; // Đặt chữ in đậm
            cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
          } else {
            cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
          }
        }
      }
    }

    // Lưu file dưới dạng stream
    final List<int> bytes = workbook.saveAsStream();

    // Tải xuống file Excel trên trình duyệt
    html.AnchorElement(
        href:
            "data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}")
      ..setAttribute("download", fileName)
      ..click();
  }

  /// Tạo và tải xuống file Excel
  static void createExcelFileCode(
      List<List<String>> dataExcel, String fileName) {
    // Tạo một workbook
    final excel.Workbook workbook = excel.Workbook();

    // Accessing worksheet via index.
    final excel.Worksheet sheet = workbook.worksheets[0];
    sheet.name = fileName; // Đổi tên sheet

    // Thêm dữ liệu vào sheet
    for (int rowIndex = 0; rowIndex < dataExcel.length; rowIndex++) {
      for (int colIndex = 0;
          colIndex < dataExcel[rowIndex].length;
          colIndex++) {
        final cell = sheet.getRangeByIndex(rowIndex + 1, colIndex + 1);
        cell.setText(dataExcel[rowIndex][colIndex]);
        // Thêm bo viền cho ô
        cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;
        if (rowIndex == 0) {
          cell.cellStyle.bold = true; // Đặt chữ in đậm
          cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
        } else {
          cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
        }
      }
    }

    // Thêm dữ liệu vào sheet kiểm kho 2
    List<List<String>> transformedData = [];
    // Set<String> seenCodeMH = {}; // Set lưu các mã mặt hàng đã xuất hiện
    for (int i = 0; i < dataExcel.length; i++) {
      List<String> row = dataExcel[i];
      // for (var row in dataExcel) {
      if (row.length >= 4) {
        // Lấy mã khách hàng và tên khách hàng
        String codeMH; // Mã mặt hàng
        String nameMH; // Tên mặt hàng
        String numberHH; // Số lượng
        String specifications;
        String price;

        codeMH = row[1]; // Mã mặt hàng
        nameMH = row[0]; // Tên mặt hàng
        numberHH = row[2]; // Số lượng
        specifications = row[3]; // Thông số đặc tả
        price = row[4]; // Đơn giá

        List<String> newRow;
        if (i == 0) {
          newRow = [
            "Ngày",
            "Thứ tự",
            "Người phụ trách",
            "Kho - Địa điểm",
            "Mã mặt hàng",
            "Tên mặt hàng",
            "Số lượng",
            "Ghi chú",
            "S/L phụ",
          ];
        } else {
          newRow = [
            "",
            "",
            "",
            "",
            codeMH,
            nameMH,
            numberHH,
            "",
            "",
          ];
        }
        transformedData.add(newRow);
      }
    }

// Thêm dữ liệu vào sheet hàng hoá
    List<List<String>> transformedData2 = [];
    for (int i = 0; i < dataExcel.length; i++) {
      List<String> row = dataExcel[i];
      // for (var row in dataExcel) {
      if (row.length >= 4) {
        String codeMH; // Mã mặt hàng
        String nameMH; // Tên mặt hàng
        String specifications;
        String price;

        codeMH = row[1]; // Mã mặt hàng
        nameMH = row[0]; // Tên mặt hàng
        specifications = row[3]; // Thông số đặc tả
        price = row[4]; // Đơn giá

        List<String> newRow;
        if (i == 0) {
          newRow = [
            codeMH,
            nameMH,
            "Loại quy cách",
            "Thông số đặc tả",
            "Thông số",
            "Danh mục mặt hàng",
            "Hàng đóng kiện",
            "Quản lý số lượng",
            "Quy trình",
            price,
            "Giá mua Tình trạng thuế",
            "Giá bán",
            "Giá bán Tình trạng thuế"
          ];
        } else {
          newRow = [
            codeMH,
            nameMH,
            "",
            "",
            specifications,
            "3",
            "",
            "",
            "",
            price,
            "",
            "",
            ""
          ];
        }
        transformedData2.add(newRow);
      }
    }

    final excel.Worksheet sheet2 = workbook.worksheets.add();
    sheet2.name = "Template_kiểm kho 2"; // Đặt tên sheet phụ
    // Thêm dữ liệu vào sheet
    for (int rowIndex = 0; rowIndex < transformedData.length; rowIndex++) {
      for (int colIndex = 0;
          colIndex < transformedData[rowIndex].length;
          colIndex++) {
        final cell = sheet2.getRangeByIndex(rowIndex + 1, colIndex + 1);
        cell.setText(transformedData[rowIndex][colIndex]);
        // Thêm bo viền cho ô
        cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;

        if (rowIndex == 0) {
          cell.cellStyle.bold = true; // Đặt chữ in đậm
          cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
        } else {
          cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
        }
      }
    }

    final excel.Worksheet sheet3 = workbook.worksheets.add();
    sheet3.name = "Template mặt hàng"; // Đặt tên sheet phụ

    // Thêm dữ liệu vào sheet
    for (int rowIndex = 0; rowIndex < transformedData2.length; rowIndex++) {
      for (int colIndex = 0;
          colIndex < transformedData2[rowIndex].length;
          colIndex++) {
        final cell = sheet3.getRangeByIndex(rowIndex + 1, colIndex + 1);
        cell.setText(transformedData2[rowIndex][colIndex]);

        // Thêm bo viền cho ô
        cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;

        if (rowIndex == 0) {
          cell.cellStyle.bold = true; // Đặt chữ in đậm
          cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
        } else {
          cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
        }
      }
    }

    // Lưu file dưới dạng stream
    final List<int> bytes = workbook.saveAsStream();
    // Tải xuống file Excel trên trình duyệt
    html.AnchorElement(
        href:
            "data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}")
      ..setAttribute("download", fileName)
      ..click();
  }

///////////////////
  static void createExcel(List<List<String>> dataExcel, String fileName) {
    final excel.Workbook workbook = excel.Workbook();
    final excel.Worksheet sheet = workbook.worksheets[0];

    // Đặt tên sheet
    sheet.name = 'Bảng Kế Toán';

    // Danh sách tiêu đề hàng 1
    List<String> header1 = [
      'STT',
      'Số phiếu',
      'Chứng từ',
      '',
      'Diễn giải',
      'Mã hiệu hàng',
      '',
      'Đơn vị tính',
      'Số lượng',
      'Số tiền phát sinh',
      'Số hiệu TK đối ứng',
      '',
      '',
      '',
    ];

    // Danh sách tiêu đề hàng 2
    List<String> header2 = [
      '',
      '',
      'Số hiệu',
      'Ngày tháng',
      '',
      'Nhập',
      'Xuất',
      '',
      '',
      '',
      'Nợ',
      'Chi tiết',
      'Có',
      'Chi tiết'
    ];

    // Hợp nhất các ô cần thiết
    sheet.getRangeByIndex(1, 1, 2, 1).merge(); // STT
    sheet.getRangeByIndex(1, 2, 2, 2).merge(); // Ngày tháng ghi sổ
    sheet
        .getRangeByIndex(1, 3, 1, 4)
        .merge(); // Chứng từ (Số hiệu + Ngày tháng)
    sheet.getRangeByIndex(1, 5, 2, 5).merge(); // Diễn giải
    sheet.getRangeByIndex(1, 6, 1, 7).merge(); // Mã hiệu hàng (Nhập + Xuất)
    sheet.getRangeByIndex(1, 8, 2, 8).merge(); // Số lượng
    sheet.getRangeByIndex(1, 9, 2, 9).merge(); // Số lượng
    sheet.getRangeByIndex(1, 10, 2, 10).merge(); // Số tiền phát sinh
    sheet.getRangeByIndex(1, 11, 1, 14).merge(); // Số hiệu TK đối ứng (Nợ + Có)

    // Ghi tiêu đề hàng 1
    for (int col = 0; col < header1.length; col++) {
      sheet.getRangeByIndex(1, col + 1).setText(header1[col]);
    }

    // Ghi tiêu đề hàng 2
    for (int col = 0; col < header2.length; col++) {
      sheet.getRangeByIndex(2, col + 1).setText(header2[col]);
    }

    // Căn giữa và định dạng in đậm
    final excel.Style headerStyle = workbook.styles.add('HeaderStyle');
    headerStyle.bold = true;
    headerStyle.hAlign = excel.HAlignType.center;
    headerStyle.vAlign = excel.VAlignType.center;
    headerStyle.backColor = '#FFF9C4'; // Màu nền cho header
    headerStyle.fontSize = 12; // Cỡ chữ
    headerStyle.borders.all.lineStyle = excel.LineStyle.thin; // Bo viền

    // Áp dụng style cho tiêu đề
    sheet.getRangeByIndex(1, 1, 2, header1.length).cellStyle = headerStyle;

    // Thiết lập độ rộng cột
    sheet.getRangeByIndex(1, 1).columnWidth = 10; // STT
    sheet.getRangeByIndex(1, 2).columnWidth = 15; // Ngày tháng ghi sổ
    sheet.getRangeByIndex(1, 3).columnWidth = 15; // Số hiệu
    sheet.getRangeByIndex(1, 4).columnWidth = 12; // Ngày tháng
    sheet.getRangeByIndex(1, 5).columnWidth = 30; // Diễn giải
    sheet.getRangeByIndex(1, 6).columnWidth = 30; // Nhập
    sheet.getRangeByIndex(1, 7).columnWidth = 30; // Xuất
    sheet.getRangeByIndex(1, 8).columnWidth = 10; // Đơn vị tính
    sheet.getRangeByIndex(1, 9).columnWidth = 10; // Số lượng
    sheet.getRangeByIndex(1, 10).columnWidth = 15; // Số tiền phát sinh
    sheet.getRangeByIndex(1, 11).columnWidth = 12; // Nợ
    sheet.getRangeByIndex(1, 12).columnWidth = 12; // Chi tiết
    sheet.getRangeByIndex(1, 13).columnWidth = 12; // Có
    sheet.getRangeByIndex(1, 14).columnWidth = 15; // Chi tiết
////////////////////////////////
    // 👉 **Sắp xếp dữ liệu theo cột "Ngày tháng ghi sổ"**
    // sortByDate(dataExcel, 1); // Cột "Ngày tháng ghi sổ" là cột thứ 2 (index 1)

    // Thêm dữ liệu vào sheet
    for (int rowIndex = 0; rowIndex < dataExcel.length; rowIndex++) {
      String isMissingSerial = dataExcel[rowIndex][10].toString().trim();
      for (int colIndex = 0;
          colIndex < dataExcel[rowIndex].length;
          colIndex++) {
        final cell = sheet.getRangeByIndex(rowIndex + 3, colIndex + 1);
        if (colIndex < 14) {
          cell.setText(dataExcel[rowIndex][colIndex]);

          // Thêm bo viền cho ô
          cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;
          cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu

          // Kiểm tra nếu số hiệu trống thì đổi màu nền thành đỏ
          cell.cellStyle.backColor =
              (isMissingSerial == "131" || isMissingSerial == "1331")
                  ? '#b7fcc4'
                  : '#DFEBF5';
        }
      }
    }
// // Thêm dữ liệu vào sheet hàng hoá
    List<List<String>> transformedData = [];
    Set<String> seenCodeMH = {}; // Set lưu các mã mặt hàng đã xuất hiện
    List<String> newTitle = [
      "Mã mặt hàng",
      "Tên mặt hàng",
      "Loại quy cách",
      "Thông số đặc tả",
      'Thông số',
      "Danh mục mặt hàng",
      "Hàng đóng kiện",
      "Quản lý số lượng",
      "Quy trình",
      "Đơn giá",
      "Giá mua Tình trạng thuế",
      "Giá bán",
      "Giá bán Tình trạng thuế"
    ];
    transformedData.add(newTitle);

    for (int i = 0; i < dataExcel.length; i++) {
      List<String> row = dataExcel[i];
      // for (var row in dataExcel) {
      if (row.length > 14) {
        // Lấy mã khách hàng và tên khách hàng
        String codeMH; // Mã mặt hàng
        String nameMH; // Tên mặt hàng
        String unit; // Đơn vị tính
        String price; // Đơn giá
        codeMH = row[23]; // Mã mặt hàng
        nameMH = row[24]; // Tên mặt hàng
        unit = row[25]; // Đơn vị tính
        price = row[27]; // Đơn giá

        // Nếu không phải hàng header và codeMH đã tồn tại thì bỏ qua
        if (i != 0 && seenCodeMH.contains(codeMH)) {
          continue;
        }
        // Thêm codeMH vào set để kiểm tra cho các hàng sau
        seenCodeMH.add(codeMH);

        List<String> newRow;
        newRow = [
          codeMH,
          nameMH,
          "",
          "",
          unit,
          "3",
          "",
          "",
          "",
          price,
          "",
          "",
          ""
        ];
        transformedData.add(newRow);
      }
    }

//     // Thêm dữ liệu vào sheet nhà cung cấp
    List<List<String>> customerData = [];
    Set<String> seenCodeCustomer = {}; // Set lưu các mã mặt hàng đã xuất hiện
    List<String> newTitle2 = [
      "Mã KH/NCC",
      "Tên KH/NCC",
      "Điện thoại",
      "Mã 1",
      "Địa chỉ công ty",
      "Email",
      "Người phụ trách",
      "Mã số thuế"
    ];
    customerData.add(newTitle2);
    for (int i = 0; i < dataExcel.length; i++) {
      List<String> row = dataExcel[i];
      // for (var row in dataExcel) {
      if (row.length > 14) {
        // Lấy mã khách hàng và tên khách hàng
        String codeCustomer = row[16]; // Mã KH/ NCC
        String nameCustomer = row[17]; // Tên công ty
        String person = row[18]; // Người phụ trách
        String customerAddress; // Địa chỉ công ty
        String customerPhone; // Điện thoại
        customerAddress = row[36]; // Địa chỉ công ty
        customerPhone = row[37]; // Điện thoại
        String customerNameBuy = row[38]; // Tên người mua
        String customerAddressBuy = row[39]; // Địa chỉ người mua

        // Nếu codeCustomer chưa tồn tại, thêm vào danh sách
        if (!seenCodeCustomer.contains(codeCustomer)) {
          seenCodeCustomer.add(codeCustomer);

          List<String> newRowBuy = [
            codeCustomer,
            nameCustomer,
            customerPhone,
            "",
            customerAddress,
            "",
            person,
            codeCustomer
          ];
          customerData.add(newRowBuy);
        }

        // Kiểm tra xem người phụ trách đã xuất hiện chưa
        if (!seenCodeCustomer.contains(person) && person.isNotEmpty) {
          seenCodeCustomer.add(person);

          List<String> newRowSale = [
            person,
            customerNameBuy,
            "",
            "",
            customerAddressBuy,
            "",
            codeCustomer,
            person
          ];
          customerData.add(newRowSale);
        }
      } else {
        print("Dòng $i không đủ dữ liệu, bỏ qua.");
      }
    }

    ///
    final excel.Worksheet sheet2 = workbook.worksheets.add();
    sheet2.name = "Mặt hàng"; // Đặt tên sheet phụ
    // Thêm dữ liệu vào sheet
    for (int rowIndex = 0; rowIndex < transformedData.length; rowIndex++) {
      for (int colIndex = 0;
          colIndex < transformedData[rowIndex].length;
          colIndex++) {
        final cell = sheet2.getRangeByIndex(rowIndex + 1, colIndex + 1);
        cell.setText(transformedData[rowIndex][colIndex]);
        // Thêm bo viền cho ô
        cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;
        if (rowIndex == 0) {
          cell.cellStyle.bold = true; // Đặt chữ in đậm
          cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
        } else {
          cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
        }
      }
    }

    final excel.Worksheet sheet3 = workbook.worksheets.add();
    sheet3.name = "Khách hàng"; // Đặt tên sheet phụ

    // Thêm dữ liệu vào sheet
    for (int rowIndex = 0; rowIndex < customerData.length; rowIndex++) {
      for (int colIndex = 0;
          colIndex < customerData[rowIndex].length;
          colIndex++) {
        final cell = sheet3.getRangeByIndex(rowIndex + 1, colIndex + 1);
        cell.setText(customerData[rowIndex][colIndex]);

        // Thêm bo viền cho ô
        cell.cellStyle.borders.all.lineStyle = excel.LineStyle.thin;

        if (rowIndex == 0) {
          cell.cellStyle.bold = true; // Đặt chữ in đậm
          cell.cellStyle.backColor = '#FFF9C4'; // Màu nền cho header
        } else {
          cell.cellStyle.backColor = '#DFEBF5'; // Màu nền cho dữ liệu
        }
      }
    }

    // Lưu file dưới dạng stream
    final List<int> bytes = workbook.saveAsStream();
    // Tải xuống file Excel trên trình duyệt
    html.AnchorElement(
        href:
            "data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}")
      ..setAttribute("download", fileName)
      ..click();
  }

  /// 🏷 Hàm sắp xếp theo cột ngày tháng
  static void sortByDate(List<List<String>> data, int dateColumnIndex) {
    data.sort((a, b) {
      DateTime? dateA = _parseDate(a[dateColumnIndex]);
      DateTime? dateB = _parseDate(b[dateColumnIndex]);
      if (dateA == null || dateB == null) return 0;
      return dateA.compareTo(dateB);
    });
  }

  /// 🏷 Chuyển đổi chuỗi ngày tháng thành DateTime
  static DateTime? _parseDate(String dateString) {
    try {
      List<String> parts = dateString.split('/');
      if (parts.length == 3) {
        int day = int.parse(parts[0]);
        int month = int.parse(parts[1]);
        int year = int.parse(parts[2]);
        return DateTime(year, month, day);
      }
    } catch (e) {
      return null;
    }
    return null;
  }
}
