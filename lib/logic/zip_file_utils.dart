import 'dart:convert';
import 'dart:typed_data';
import 'package:file_picker/file_picker.dart';
import 'package:archive/archive.dart';

class ZipFileUtils {
  // Hàm chọn file zip và trả về danh sách các file đã chọn cùng nội dung
  static Future<List<Map<String, dynamic>>> pickZipFiles() async {
    try {
      // Chọn file .zip
      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['zip'],
        allowMultiple: true,
      );

      if (result != null && result.files.isNotEmpty) {
        // Trả về danh sách các file đã chọn
        return result.files
            .map((file) => {
                  'name': file.name,
                  'data': file.bytes, // Lưu data của file
                })
            .toList();
      } else {
        print('Người dùng không chọn file.');
        return [];
      }
    } catch (e) {
      print('Lỗi khi tải file: $e');
      return [];
    }
  }

  // Hàm giải nén và tìm file invoice.html trong một file zip
  static Future<String?> extractInvoiceHtml(Uint8List zipBytes) async {
    try {
      // Giải nén file zip
      final archive = ZipDecoder().decodeBytes(zipBytes);

      // Tìm và trả về nội dung file invoice.html
      for (final file in archive) {
        if (file.isFile && file.name == 'invoice.html') {
          return utf8.decode(file.content as List<int>);
        }
      }
      return null; // Không tìm thấy file
    } catch (e) {
      print('Lỗi khi giải nén: $e');
      return null;
    }
  }
}
