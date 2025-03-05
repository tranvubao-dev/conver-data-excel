import 'package:flutter/material.dart';
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
      title: 'TEMPLATE',
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
      home: const MainTabPage(),
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
  final GlobalKey<ScaffoldState> scaffoldKey = GlobalKey<ScaffoldState>();

  final List<Widget> _pages = const [
    ShoppingPage(),
    SalesPage(),
    CreateCodePage(),
    FilterDataPage(),
    SyntheticPage(),
  ];

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
        length: 5, // Số lượng tab
        initialIndex: _currentIndex, // Tab mặc định
        child: Scaffold(
          backgroundColor: Colors.transparent,
          appBar: AppBar(
            title: const Text('CREATE TEMPLATE'),
            leading: Builder(
              builder: (context) {
                return IconButton(
                  icon: const Icon(Icons.settings),
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
                Tab(icon: Icon(Icons.shopping_cart), text: 'Template Mua Hàng'),
                Tab(icon: Icon(Icons.storefront), text: 'Template Bán Hàng'),
                Tab(icon: Icon(Icons.barcode_reader), text: 'Tạo Mã'),
                Tab(
                    icon: Icon(Icons.filter_alt_sharp),
                    text: 'Lọc mã sản phẩm'),
                Tab(icon: Icon(Icons.summarize), text: 'File tổng hợp'),
              ],
            ),
            backgroundColor: Colors.transparent, // Làm trong suốt AppBar
            elevation: 0, // Xóa bóng AppBar
            flexibleSpace: Container(
              decoration: const BoxDecoration(
                image: DecorationImage(
                  image: AssetImage('assets/tabbar.webp'), // Ảnh nền AppBar
                  fit: BoxFit.cover, // Đảm bảo ảnh phủ kín AppBar
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
                    Navigator.pop(
                        context); // Đóng Drawer trước khi chuyển trang
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
                        return Row(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            // Image.asset(
                            //   'assets/tabbar.webp', // Thay bằng đường dẫn mã QR
                            //   width: 40,
                            //   height: 40,
                            // ),
                            const SizedBox(width: 8),
                            const Icon(Icons.qr_code,
                                color: Colors.black54), // Icon QR
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
              ],
            ),
          ),
          body: Container(
            decoration: const BoxDecoration(
              image: DecorationImage(
                image:
                    AssetImage('assets/background.webp'), // Đường dẫn ảnh nền
                fit: BoxFit.cover, // Cách ảnh nền phủ kín màn hình
              ),
            ),
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
        ));
  }
}

// Trang Cài đặt
class SettingsPage extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Setting')),
      body: const Center(
          child: Text('Developing', style: TextStyle(fontSize: 24))),
    );
  }
}
