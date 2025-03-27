import 'package:flutter/material.dart';
import 'package:flutter_fireworks/fireworks_controller.dart';

class CommonFireworks {
  final FireworksController fireworksController;

  CommonFireworks()
      : fireworksController = FireworksController(
          colors: [
            const Color(0xFFFF4C40), // Coral
            const Color(0xFF6347A6), // Purple Haze
            const Color(0xFF7FB13B), // Greenery
            const Color(0xFF82A0D1), // Serenity Blue
            const Color(0xFFF7B3B2), // Rose Quartz
            const Color(0xFF864542), // Marsala
            const Color(0xFFB04A98), // Orchid
            const Color(0xFF008F6C), // Sea Green
            const Color(0xFFFFD033), // Pastel Yellow
            const Color(0xFFFF6F7C), // Pink Grapefruit
          ],
          minExplosionDuration: 0.5,
          maxExplosionDuration: 3.5,
          minParticleCount: 125,
          maxParticleCount: 275,
          fadeOutDuration: 0.4,
        );

  void firework() {
    fireworksController.fireMultipleRockets(
      minRockets: 20,
      maxRockets: 50,
      launchWindow: const Duration(milliseconds: 600),
    );
  }
}
