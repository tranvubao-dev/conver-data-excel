import 'package:flutter/material.dart';

class CommonButton extends StatelessWidget {
  final VoidCallback? onPressed;
  final String label;
  final IconData? icon;
  final bool isDisabled;
  final double elevation;
  final Color enabledColor;
  final Color disabledColor;

  const CommonButton({
    Key? key,
    required this.onPressed,
    required this.label,
    this.icon,
    this.isDisabled = false,
    this.elevation = 10,
    this.enabledColor = const Color.fromARGB(255, 219, 237, 252),
    this.disabledColor = Colors.grey,
  }) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return icon != null
        ? ElevatedButton.icon(
            onPressed: isDisabled ? null : onPressed,
            icon: Icon(icon),
            label: Text(label),
            style: ElevatedButton.styleFrom(
              padding: const EdgeInsets.symmetric(vertical: 16, horizontal: 32),
              textStyle: const TextStyle(fontSize: 18),
              elevation: elevation,
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(10),
              ),
              backgroundColor: isDisabled ? disabledColor : enabledColor,
              shadowColor: Colors.blue[800],
            ),
          )
        : ElevatedButton(
            onPressed: isDisabled ? null : onPressed,
            style: ElevatedButton.styleFrom(
              padding: const EdgeInsets.symmetric(vertical: 16, horizontal: 32),
              textStyle: const TextStyle(fontSize: 18),
              elevation: elevation,
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(10),
              ),
              backgroundColor: isDisabled ? disabledColor : enabledColor,
              shadowColor: Colors.blue[800],
            ),
            child: Text(label),
          );
  }
}
