class SafeValueHandler {
  static String safeValue(List row, int index, {String defaultValue = ""}) {
    // Nếu index nằm ngoài phạm vi, trả về defaultValue
    if (index < 0 || index >= row.length) {
      return defaultValue;
    }
    // row[index] có thể là null, nên dùng toString() an toàn
    return row[index]?.toString() ?? defaultValue;
  }
}
