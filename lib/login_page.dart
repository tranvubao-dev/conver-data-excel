import 'package:flutter/material.dart';
import 'package:html_to_excel/db_service.dart';
import 'package:html_to_excel/main.dart';
import 'package:shared_preferences/shared_preferences.dart';

class LoginPage extends StatefulWidget {
  const LoginPage({super.key});

  @override
  State<LoginPage> createState() => _LoginPageState();
}

class _LoginPageState extends State<LoginPage> {
  final TextEditingController _usernameController = TextEditingController();
  final TextEditingController _passwordController = TextEditingController();
  String? _errorText;

  Future<void> _login() async {
    final username = _usernameController.text;
    final password = _passwordController.text;

    if (username == 'Admin' && password == 'Admin') {
      final prefs = await SharedPreferences.getInstance();
      await prefs.setInt('lastLogin', DateTime.now().millisecondsSinceEpoch);

      // Chuyển sang trang chính
      Navigator.pushReplacement(
        context,
        MaterialPageRoute(builder: (context) => const MainTabPage()),
      );
    } else {
      setState(() {
        _errorText = 'Sai tài khoản hoặc mật khẩu!';
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    double screenWidth = MediaQuery.of(context).size.width;
    double textFieldWidth = screenWidth / 4; // TextField bằng 1/4 màn hình

    return Scaffold(
      body: LayoutBuilder(
        builder: (context, constraints) {
          return Container(
            width: constraints.maxWidth,
            height: constraints.maxHeight,
            decoration: const BoxDecoration(
              color: Colors.black, // Màu nền phòng trường hợp ảnh không đủ lớn
              image: DecorationImage(
                image: AssetImage('assets/login.webp'),
                fit: BoxFit.fill, // Giữ tỷ lệ ảnh mà vẫn phủ toàn màn hình
                alignment: Alignment.center, // Đảm bảo ảnh căn giữa
              ),
            ),
            child: Center(
              child: Padding(
                padding: const EdgeInsets.all(24.0),
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    const SizedBox(height: 20),
                    SizedBox(
                      width: textFieldWidth,
                      child: TextField(
                        controller: _usernameController,
                        decoration: InputDecoration(
                          labelText: 'Tài khoản',
                          border: const OutlineInputBorder(),
                          errorText: _errorText,
                        ),
                      ),
                    ),
                    const SizedBox(height: 10),
                    SizedBox(
                      width: textFieldWidth,
                      child: TextField(
                        controller: _passwordController,
                        decoration: InputDecoration(
                          labelText: 'Mật khẩu',
                          border: const OutlineInputBorder(),
                          errorText: _errorText,
                        ),
                        obscureText: true,
                      ),
                    ),
                    const SizedBox(height: 40),
                    ElevatedButton(
                      onPressed: _login,
                      child: const Text(
                        'Đăng nhập',
                        style: TextStyle(
                            fontSize: 20, fontWeight: FontWeight.bold),
                      ),
                    ),
                  ],
                ),
              ),
            ),
          );
        },
      ),
    );
  }
}
