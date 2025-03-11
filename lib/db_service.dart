import 'package:mysql1/mysql1.dart';

class DatabaseService {
  static Future<MySqlConnection> connect() async {
    final settings = ConnectionSettings(
      host: '103.200.23.120', // Thay bằng IP MySQL Server
      port: 3306,
      user: 'converco_admin', // Thay bằng user MySQL
      password: 'q{XaSibqn%dA', // Thay bằng mật khẩu
      db: 'converco_database', // Thay bằng tên database
    );

    return await MySqlConnection.connect(settings);
  }
}
