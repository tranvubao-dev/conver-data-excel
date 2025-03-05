import 'dart:convert';
import 'dart:typed_data';
import 'package:file_picker/file_picker.dart';
import 'package:archive/archive.dart';

class ZipFileHandler {
  // Chọn file .zip
  static Future<Uint8List?> pickZipFile() async {
    try {
      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['zip'],
      );
      if (result != null && result.files.single.bytes != null) {
        return result.files.single.bytes;
      } else {
        print('Người dùng không chọn file.');
        return null;
      }
    } catch (e) {
      print('Lỗi khi tải file: $e');
      return null;
    }
  }

  // Giải nén file zip và trả về nội dung file invoice.html
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
