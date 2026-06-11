package jp.co.pm.ai.desktop.io;

/**
 * UNC パス内の既知セグメント名を実フォルダ表記へ揃える。
 *
 * <p>トークン化で {@code 002␠␠加工G} が {@code 002␠加工G} に潰れると RPA がファイルを見つけられない。
 */
public final class UncPathSegmentRepair {

    /** 湖南共有の実フォルダ名（002 と 加工G の間はスペース2つ）。 */
    public static final String KONAN_002_KAKOG_SEGMENT = "002  加工G";

    private static final String KONAN_002_KAKOG_REGEX = "\\\\002\\s+加工G";

    private UncPathSegmentRepair() {}

    /** 既知の誤表記（主に空白潰れ）を修復する。 */
    public static String repair(String path) {
        if (path == null || path.isBlank()) {
            return path != null ? path : "";
        }
        return path.replaceAll(KONAN_002_KAKOG_REGEX, "\\\\" + KONAN_002_KAKOG_SEGMENT);
    }
}
