package jp.co.pm.ai.desktop.io.win32;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.awt.Color;
import java.awt.image.BufferedImage;

import org.junit.jupiter.api.Test;

class MstscWindowCaptureTest {

    @Test
    void isLikelyBlank_detectsMostlyBlack() {
        BufferedImage black = new BufferedImage(320, 240, BufferedImage.TYPE_INT_RGB);
        assertTrue(MstscWindowCapture.isLikelyBlank(black));
        assertTrue(MstscWindowCapture.isLikelyBlank(null));
    }

    @Test
    void isLikelyBlank_acceptsNormalFrame() {
        BufferedImage image = new BufferedImage(320, 240, BufferedImage.TYPE_INT_RGB);
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                image.setRGB(x, y, Color.LIGHT_GRAY.getRGB());
            }
        }
        assertFalse(MstscWindowCapture.isLikelyBlank(image));
    }
}
