package jp.co.pm.ai.desktop.io.win32;

import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.image.BufferedImage;
import java.util.Optional;

import com.sun.jna.Memory;
import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.GDI32;
import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef.HBITMAP;
import com.sun.jna.platform.win32.WinDef.HDC;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.platform.win32.WinDef.RECT;
import com.sun.jna.platform.win32.WinNT.HANDLE;
import com.sun.jna.platform.win32.WinGDI;
import com.sun.jna.platform.win32.WinGDI.BITMAPINFO;
import com.sun.jna.platform.win32.WinGDI.BITMAPINFOHEADER;

import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/** mstsc ウィンドウの読み取り専用キャプチャ（PrintWindow → 画面矩形フォールバック）。 */
public final class MstscWindowCapture {

    private static final int PW_RENDERFULLCONTENT = 0x00000002;
    private static final int BI_RGB = 0;

    private MstscWindowCapture() {}

    public static double brightPixelRatio(BufferedImage image) {
        if (image == null) {
            return 0.0;
        }
        int w = image.getWidth();
        int h = image.getHeight();
        if (w < 1 || h < 1) {
            return 0.0;
        }
        int stepX = Math.max(1, w / 32);
        int stepY = Math.max(1, h / 32);
        int samples = 0;
        int bright = 0;
        for (int y = 0; y < h; y += stepY) {
            for (int x = 0; x < w; x += stepX) {
                samples++;
                int rgb = image.getRGB(x, y);
                int r = (rgb >> 16) & 0xFF;
                int g = (rgb >> 8) & 0xFF;
                int b = rgb & 0xFF;
                if (r + g + b > 48) {
                    bright++;
                }
            }
        }
        return samples == 0 ? 0.0 : (bright * 100.0 / samples);
    }

    public static boolean isSupported() {
        return RemoteDesktopLauncher.isSupportedPlatform();
    }

    public static Optional<BufferedImage> captureWindow(long hwndNative) {
        if (!isSupported() || hwndNative == 0L) {
            return Optional.empty();
        }
        HWND hwnd = toHwnd(hwndNative);
        User32 user32 = User32.INSTANCE;
        if (!user32.IsWindow(hwnd)) {
            return Optional.empty();
        }
        Optional<BufferedImage> print = captureViaPrintWindow(hwnd);
        if (print.isPresent() && !isLikelyBlank(print.get())) {
            return print;
        }
        Optional<BufferedImage> screen = captureViaScreenRect(hwnd);
        if (screen.isPresent() && !isLikelyBlank(screen.get())) {
            return screen;
        }
        if (print.isPresent()) {
            return print;
        }
        if (screen.isPresent()) {
            return screen;
        }
        return Optional.empty();
    }

    /** 外枠→クライアントの順でキャプチャを試す。 */
    public static Optional<BufferedImage> captureTarget(MstscCaptureTarget target) {
        if (target == null || !target.isValid()) {
            return Optional.empty();
        }
        for (long hwnd : target.handlesToTry()) {
            Optional<BufferedImage> frame = captureWindow(hwnd);
            if (frame.isPresent() && !isLikelyBlank(frame.get())) {
                return frame;
            }
        }
        for (long hwnd : target.handlesToTry()) {
            Optional<BufferedImage> frame = captureWindow(hwnd);
            if (frame.isPresent()) {
                return frame;
            }
        }
        return Optional.empty();
    }

    /** 有効画素が極端に少ない（黒画面等）場合 {@code true}。 */
    public static boolean isLikelyBlank(BufferedImage image) {
        if (image == null || image.getWidth() < 8 || image.getHeight() < 8) {
            return true;
        }
        return brightPixelRatio(image) < 2.0;
    }

    private static Optional<BufferedImage> captureViaPrintWindow(HWND hwnd) {
        User32 user32 = User32.INSTANCE;
        GDI32 gdi32 = GDI32.INSTANCE;
        RECT rect = new RECT();
        if (!user32.GetWindowRect(hwnd, rect)) {
            return Optional.empty();
        }
        int width = rect.right - rect.left;
        int height = rect.bottom - rect.top;
        if (width < 1 || height < 1) {
            return Optional.empty();
        }

        HDC hdcWindow = user32.GetDC(hwnd);
        if (hdcWindow == null) {
            return Optional.empty();
        }
        HDC hdcMem = null;
        HBITMAP hBitmap = null;
        HANDLE oldBitmap = null;
        try {
            hdcMem = gdi32.CreateCompatibleDC(hdcWindow);
            if (hdcMem == null) {
                return Optional.empty();
            }
            hBitmap = gdi32.CreateCompatibleBitmap(hdcWindow, width, height);
            if (hBitmap == null) {
                return Optional.empty();
            }
            oldBitmap = gdi32.SelectObject(hdcMem, hBitmap);
            boolean printed =
                    user32.PrintWindow(hwnd, hdcMem, PW_RENDERFULLCONTENT)
                            || user32.PrintWindow(hwnd, hdcMem, 0);
            if (!printed) {
                return Optional.empty();
            }
            return Optional.ofNullable(bitmapToImage(hBitmap, width, height));
        } finally {
            if (oldBitmap != null && hdcMem != null) {
                gdi32.SelectObject(hdcMem, oldBitmap);
            }
            if (hBitmap != null) {
                gdi32.DeleteObject(hBitmap);
            }
            if (hdcMem != null) {
                gdi32.DeleteDC(hdcMem);
            }
            user32.ReleaseDC(hwnd, hdcWindow);
        }
    }

    private static Optional<BufferedImage> captureViaScreenRect(HWND hwnd) {
        User32 user32 = User32.INSTANCE;
        RECT rect = new RECT();
        if (!user32.GetWindowRect(hwnd, rect)) {
            return Optional.empty();
        }
        int width = rect.right - rect.left;
        int height = rect.bottom - rect.top;
        if (width < 1 || height < 1) {
            return Optional.empty();
        }
        try {
            Robot robot = new Robot();
            BufferedImage image =
                    robot.createScreenCapture(new Rectangle(rect.left, rect.top, width, height));
            return Optional.ofNullable(image);
        } catch (Exception ex) {
            return Optional.empty();
        }
    }

    private static BufferedImage bitmapToImage(HBITMAP hBitmap, int width, int height) {
        GDI32 gdi32 = GDI32.INSTANCE;
        BITMAPINFO bmi = new BITMAPINFO();
        bmi.bmiHeader.biSize = bmi.bmiHeader.size();
        bmi.bmiHeader.biWidth = width;
        bmi.bmiHeader.biHeight = -height;
        bmi.bmiHeader.biPlanes = 1;
        bmi.bmiHeader.biBitCount = 32;
        bmi.bmiHeader.biCompression = BI_RGB;

        Memory buffer = new Memory((long) width * height * 4);
        HDC hdc = gdi32.CreateCompatibleDC(null);
        if (hdc == null) {
            return null;
        }
        try {
            int lines =
                    gdi32.GetDIBits(
                            hdc,
                            hBitmap,
                            0,
                            height,
                            buffer,
                            bmi,
                            WinGDI.DIB_RGB_COLORS);
            if (lines == 0) {
                return null;
            }
            BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
            int[] pixels = new int[width * height];
            buffer.read(0, pixels, 0, pixels.length);
            for (int i = 0; i < pixels.length; i++) {
                int bgr = pixels[i];
                int b = bgr & 0xFF;
                int g = (bgr >> 8) & 0xFF;
                int r = (bgr >> 16) & 0xFF;
                pixels[i] = (r << 16) | (g << 8) | b;
            }
            image.setRGB(0, 0, width, height, pixels, 0, width);
            return image;
        } finally {
            gdi32.DeleteDC(hdc);
        }
    }

    private static HWND toHwnd(long handle) {
        return new HWND(Pointer.createConstant(handle));
    }
}
