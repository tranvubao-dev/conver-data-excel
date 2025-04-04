import 'dart:async';
import 'package:flutter/material.dart';
import 'package:html_to_excel/login_page.dart';
import 'package:lottie/lottie.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:url_launcher/url_launcher.dart';
import 'sales_page.dart';
import 'shopping_page.dart';
import 'create_code_page.dart';
import 'filter_page.dart';
import 'synthetic_page.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Van Minh',
      theme: ThemeData(
        primarySwatch: Colors.blue,
        scaffoldBackgroundColor: Colors.transparent,
        textTheme: const TextTheme(
          // Tùy chỉnh tiêu đề (title) trong app
          titleLarge: TextStyle(
            fontSize: 22, // Kích thước chữ tiêu đề
            fontWeight: FontWeight.bold, // Làm chữ đậm
            fontFamily: 'Roboto', // Font đẹp cho tiêu đề
            color: Colors.black, // Màu chữ
          ),
        ),
      ),
      home: const AuthCheck(),
    );
  }
}

// Kiểm tra trạng thái đăng nhập trước khi vào app
class AuthCheck extends StatefulWidget {
  const AuthCheck({super.key});

  @override
  State<AuthCheck> createState() => _AuthCheckState();
}

class _AuthCheckState extends State<AuthCheck> {
  Future<bool> _isLoggedIn() async {
    final prefs = await SharedPreferences.getInstance();
    final lastLogin = prefs.getInt('lastLogin') ?? 0;
    final currentTime = DateTime.now().millisecondsSinceEpoch;

    // Kiểm tra nếu đã hơn 1 giờ (3600000 ms)
    if (currentTime - lastLogin > 3600000) {
      return false; // Hết hạn, cần đăng nhập lại
    }
    return true; // Vẫn trong thời gian hợp lệ
  }

  @override
  Widget build(BuildContext context) {
    return FutureBuilder<bool>(
      future: _isLoggedIn(),
      builder: (context, snapshot) {
        if (snapshot.connectionState == ConnectionState.waiting) {
          return const Scaffold(
              body: Center(child: CircularProgressIndicator()));
        } else if (snapshot.data == true) {
          return const MainTabPage();
        } else {
          return const LoginPage();
        }
      },
    );
  }
}

class MainTabPage extends StatefulWidget {
  const MainTabPage({super.key});

  @override
  State<MainTabPage> createState() => _MainTabPageState();
}

class _MainTabPageState extends State<MainTabPage> {
  int _currentIndex = 0; // Tab hiện tại
  int _currentImageIndex = 0;
  final List<String> _imageList = [
    'assets/1.webp',
    'assets/2.webp',
    'assets/3.webp',
    'assets/4.webp',
    'assets/5.webp',
  ];
  Timer? _timer;
  bool _showThought = false;
  bool _isDisposed = false;

  final GlobalKey<ScaffoldState> scaffoldKey = GlobalKey<ScaffoldState>();

  final List<Widget> _pages = const [
    // ShoppingPage(),
    // SalesPage(),
    CreateCodePage(),
    FilterDataPage(),
    SyntheticPage(),
  ];

  @override
  void initState() {
    super.initState();
    _startImageSlider();
    _startThoughtLoop();
  }

  @override
  void dispose() {
    _timer?.cancel(); // Hủy bộ đếm khi thoát màn hình
    _isDisposed = true;
    super.dispose();
  }

  Future<void> _startThoughtLoop() async {
    while (!_isDisposed) {
      // Hiện suy nghĩ
      setState(() => _showThought = true);
      await Future.delayed(const Duration(seconds: 3));

      // Ẩn suy nghĩ
      setState(() => _showThought = false);
      await Future.delayed(const Duration(seconds: 1));
    }
  }

  void _startImageSlider() {
    _timer = Timer.periodic(const Duration(seconds: 5), (timer) {
      setState(() {
        _currentImageIndex =
            (_currentImageIndex + 1) % _imageList.length; // Vòng lặp ảnh
      });
    });
  }

  // Hàm tạo hiệu ứng chuyển trang mượt mà
  Route _createRoute(Widget page) {
    return PageRouteBuilder(
      transitionDuration:
          const Duration(milliseconds: 300), // Chuyển trang mượt
      reverseTransitionDuration: Duration.zero, // Ẩn ngay khi back
      pageBuilder: (context, animation, secondaryAnimation) => page,
      transitionsBuilder: (context, animation, secondaryAnimation, child) {
        const begin = Offset(1.0, 0.0); // Hiệu ứng từ phải sang trái
        const end = Offset.zero;
        const curve = Curves.easeInOut;

        var tween =
            Tween(begin: begin, end: end).chain(CurveTween(curve: curve));
        var offsetAnimation = animation.drive(tween);

        return SlideTransition(
          position: offsetAnimation,
          child: child,
        );
      },
    );
  }

  void _showZaloPopup(BuildContext context) {
    showDialog(
      context: context,
      builder: (BuildContext context) {
        return AlertDialog(
          title: const Text(
            'Hỗ trợ qua Zalo',
            textAlign: TextAlign.center,
          ),
          content: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Image.asset(
                'assets/zalo.webp',
                width: 150,
                height: 150,
              ),
              const SizedBox(height: 10),
              const Text(
                'Quét mã QR hoặc nhấn nút bên dưới để liên hệ hỗ trợ qua Zalo.',
                textAlign: TextAlign.center,
              ),
            ],
          ),
          actions: [
            TextButton(
              onPressed: () {
                Navigator.pop(context); // Đóng popup
              },
              child: const Text('Đóng'),
            ),
            TextButton(
              onPressed: () async {
                final Uri zaloUrl = Uri.parse(
                    'https://chat.zalo.me/login'); // Thay số Zalo của bạn
                if (await canLaunchUrl(zaloUrl)) {
                  await launchUrl(zaloUrl,
                      mode: LaunchMode.externalApplication);
                } else {
                  ScaffoldMessenger.of(context).showSnackBar(
                    const SnackBar(content: Text('Không thể mở Zalo')),
                  );
                }
              },
              child: const Text('Mở Zalo'),
            ),
          ],
        );
      },
    );
  }

  @override
  Widget build(BuildContext context) {
    return DefaultTabController(
      length: 3, // Số lượng tab
      initialIndex: _currentIndex, // Tab mặc định
      child: Scaffold(
        backgroundColor: Colors.transparent,
        appBar: AppBar(
          title: const Text('CREATE TEMPLATE'),
          leading: Builder(
            builder: (context) {
              return IconButton(
                icon: Lottie.asset(
                  'assets/setting.json', // Animation từ ảnh cá nhân của bạn
                  width: 50,
                  height: 50,
                  repeat: true,
                ),
                onPressed: () {
                  Scaffold.of(context).openDrawer(); // Mở menu trái
                },
              );
            },
          ),
          bottom: TabBar(
            onTap: (index) {
              setState(() {
                _currentIndex = index;
              });
            },
            tabs: const [
              // Tab(icon: Icon(Icons.shopping_cart), text: 'Template Mua Hàng'),
              // Tab(icon: Icon(Icons.storefront), text: 'Template Bán Hàng'),
              Tab(icon: Icon(Icons.barcode_reader), text: 'Tạo Mã'),
              Tab(icon: Icon(Icons.filter_alt_sharp), text: 'Gộp File'),
              Tab(icon: Icon(Icons.summarize), text: 'File tổng hợp'),
            ],
          ),
          backgroundColor: Colors.transparent, // Làm trong suốt AppBar
          elevation: 0, // Xóa bóng AppBar
          flexibleSpace: Container(
            width: double.infinity,
            height: double.infinity,
            child: AnimatedSwitcher(
              duration: const Duration(seconds: 4), // Thời gian chuyển đổi
              transitionBuilder: (child, animation) {
                return FadeTransition(
                  opacity: animation,
                  child: child,
                );
              },
              child: Image.asset(
                _imageList[_currentImageIndex],
                key: ValueKey<int>(_currentImageIndex), // Đảm bảo đổi ảnh
                fit: BoxFit.cover, // Đảm bảo phủ kín
                width: double.infinity,
                height: double.infinity,
              ),
            ),
          ),
        ),
        drawer: Drawer(
          child: ListView(
            padding: EdgeInsets.zero,
            children: [
              const DrawerHeader(
                decoration: BoxDecoration(color: Colors.blue),
                child: Text(
                  'Menu',
                  style: TextStyle(fontSize: 24, color: Colors.white),
                ),
              ),
              ListTile(
                leading: const Icon(Icons.settings),
                title: const Text('Cài đặt'),
                onTap: () async {
                  Navigator.pop(context); // Đóng Drawer trước khi chuyển trang
                  await Navigator.of(context)
                      .push(_createRoute(SettingsPage()));
                },
              ),
              ListTile(
                leading: const Icon(Icons.book),
                title: const Text('Hướng dẫn'),
                onTap: () async {
                  const String url =
                      'https://docs.google.com/spreadsheets/d/1ZIaLAD_MGnD95HDe9k9ajl_1bXDR5RSbjEbL3t1KnXw/edit?gid=10907254#gid=10907254';
                  final Uri uri = Uri.parse(url);
                  if (await canLaunchUrl(uri)) {
                    await launchUrl(uri,
                        mode: LaunchMode
                            .platformDefault); // Hoặc LaunchMode.inAppBrowserView
                  } else {
                    print('Không thể mở đường link: $url');
                  }
                },
              ),
              ListTile(
                leading: const Icon(Icons.support_agent,
                    color: Colors.blue), // Icon hỗ trợ kỹ thuật
                title: const Text('Hỗ trợ kỹ thuật'),
                trailing: Builder(
                  builder: (context) {
                    double screenWidth = MediaQuery.of(context).size.width;
                    if (screenWidth > 400) {
                      // Nếu màn hình đủ lớn thì hiển thị QR Code
                      return const Row(
                        mainAxisSize: MainAxisSize.min,
                        children: [
                          SizedBox(width: 8),
                          Icon(Icons.qr_code, color: Colors.black54), // Icon QR
                        ],
                      );
                    } else {
                      return const Icon(Icons.qr_code, color: Colors.black54);
                    }
                  },
                ),
                onTap: () {
                  _showZaloPopup(context);
                },
              ),
              ListTile(
                leading: const Icon(Icons.logout),
                title: const Text('Logout'),
                onTap: () async {
                  Navigator.pop(context); // Đóng Drawer trước khi chuyển trang
                  Navigator.of(context).pushReplacement(
                    MaterialPageRoute(builder: (context) => const LoginPage()),
                  );
                },
              ),
            ],
          ),
        ),
        body: Stack(
          children: [
            // Sử dụng AnimatedSwitcher để thay ảnh nền
            AnimatedSwitcher(
              duration: const Duration(seconds: 4), // Thời gian chuyển đổi
              transitionBuilder: (child, animation) {
                return FadeTransition(
                  opacity: animation,
                  child: child,
                );
              },
              child: Image.asset(
                _imageList[_currentImageIndex],
                key: ValueKey<int>(_currentImageIndex), // Đảm bảo đổi ảnh
                fit: BoxFit.cover, // Đảm bảo phủ kín
                width: double.infinity,
                height: double.infinity,
              ),
            ),
            // Nội dung chính
            Container(
              child: Column(
                children: [
                  Expanded(
                      child: Stack(
                    children: _pages.asMap().entries.map((entry) {
                      int index = entry.key;
                      Widget page = entry.value;

                      return AnimatedOpacity(
                        opacity: _currentIndex == index
                            ? 1.0
                            : 0.0, // Hiển thị trang hiện tại
                        duration: const Duration(
                            milliseconds: 1200), // Thời gian chuyển động
                        child: Visibility(
                          visible: _currentIndex ==
                              index, // Chỉ hiển thị trang hiện tại
                          child: page,
                        ),
                      );
                    }).toList(),
                  )),
                  // Thêm dòng chữ và favicon ở cuối màn hình
                  const Divider(thickness: 1.0, color: Colors.grey),
                  Padding(
                    padding: const EdgeInsets.all(3.0),
                    child: Row(
                      mainAxisAlignment:
                          MainAxisAlignment.end, // Căn giữa theo chiều ngang
                      crossAxisAlignment:
                          CrossAxisAlignment.center, // Căn giữa theo chiều dọc
                      children: [
                        Image.asset(
                          'assets/40.png', // Đường dẫn tới favicon
                          width: 24, // Kích thước favicon
                          height: 24,
                        ),
                        const SizedBox(width: 8),
                      ],
                    ),
                  ),
                ],
              ),
            ),
            Positioned(
              bottom: 16,
              left: 16,
              child: Lottie.asset(
                'assets/cat.json', // Animation từ ảnh cá nhân của bạn
                width: 150,
                height: 150,
                repeat: true,
              ),
            ),
            Positioned(
              top: -10,
              left: -10,
              child: Lottie.asset(
                'assets/mouse.json', // Animation từ ảnh cá nhân của bạn
                width: 150,
                height: 150,
                repeat: true,
              ),
            ),
            Positioned(
              bottom: 50,
              left: 46,
              child: Container(
                width: 400,
                height: 300,
                child: AnimatedOpacity(
                  opacity: _showThought ? 1.0 : 0.0,
                  duration: const Duration(milliseconds: 800),
                  child: const CharacterThoughtBox(
                    thoughtText: "Hmm... Mình là Liênn đây",
                  ),
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}

// Trang Cài đặt
class SettingsPage extends StatelessWidget {
  const SettingsPage({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Setting')),
      body: const Center(
          child: Text('Developing', style: TextStyle(fontSize: 24))),
    );
  }
}

class CharacterThoughtBox extends StatelessWidget {
  final String thoughtText;

  const CharacterThoughtBox({super.key, required this.thoughtText});

  @override
  Widget build(BuildContext context) {
    return Stack(
      children: [
        // Nội dung chính của web/app
        Positioned(
          left: 50, // vị trí cố định bên trái
          bottom: 100, // cách đáy 100px
          child: ThoughtBubble(text: thoughtText),
        ),
      ],
    );
  }
}

class ThoughtBubble extends StatelessWidget {
  final String text;

  const ThoughtBubble({super.key, required this.text});

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.all(12),
      constraints: const BoxConstraints(maxWidth: 250),
      decoration: BoxDecoration(
        color: Colors.white,
        border: Border.all(color: Colors.grey),
        borderRadius: const BorderRadius.only(
          topLeft: Radius.circular(20),
          topRight: Radius.circular(20),
          bottomRight: Radius.circular(20),
        ),
        boxShadow: const [
          BoxShadow(
            color: Colors.black26,
            blurRadius: 4,
            offset: Offset(2, 2),
          )
        ],
      ),
      child: Text(
        text,
        style: const TextStyle(
          fontStyle: FontStyle.italic,
          fontSize: 16,
        ),
      ),
    );
  }
}
