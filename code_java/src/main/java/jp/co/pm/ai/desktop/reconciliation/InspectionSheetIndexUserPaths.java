package jp.co.pm.ai.desktop.reconciliation;

import java.nio.file.Path;

/**
 * 索引 CSV の検査表パスのうち、ユーザーごとに変わる {@code C:\Users\<名前>\Box} を
 * 現在の user.home 配下へ付け替える。UNC 等は変更しない。
 */
public final class InspectionSheetIndexUserPaths {

    private InspectionSheetIndexUserPaths() {}

    public static String relocateBoxHome(String filePath) {
        String home = System.getProperty("user.home", "");
        if (home == null || home.isBlank()) {
            return filePath == null ? "" : filePath;
        }
        return relocateBoxHome(filePath, Path.of(home));
    }

    public static String relocateBoxHome(String filePath, Path userHome) {
        if (filePath == null || filePath.isBlank() || userHome == null) {
            return filePath == null ? "" : filePath;
        }
        Path src;
        try {
            src = Path.of(filePath.strip());
        } catch (RuntimeException ex) {
            return filePath;
        }
        if (!isDriveRoot(src.getRoot())) {
            return filePath;
        }
        int boxIdx = boxSegmentIndex(src);
        if (boxIdx < 0) {
            return filePath;
        }
        Path remainder = src.subpath(boxIdx, src.getNameCount());
        return userHome.resolve(remainder).normalize().toString();
    }

    private static boolean isDriveRoot(Path root) {
        if (root == null) {
            return false;
        }
        String text = root.toString();
        return text.length() >= 2
                && Character.isLetter(text.charAt(0))
                && text.charAt(1) == ':';
    }

    private static int boxSegmentIndex(Path src) {
        int n = src.getNameCount();
        for (int i = 0; i + 2 < n; i++) {
            if (equalsIgnoreCase(src.getName(i), "Users")
                    && equalsIgnoreCase(src.getName(i + 2), "Box")) {
                return i + 2;
            }
        }
        return -1;
    }

    private static boolean equalsIgnoreCase(Path name, String expected) {
        return name != null
                && expected != null
                && expected.equalsIgnoreCase(name.toString());
    }
}
