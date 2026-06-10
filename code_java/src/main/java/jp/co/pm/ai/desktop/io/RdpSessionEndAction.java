package jp.co.pm.ai.desktop.io;

import java.util.Locale;

/** 子プロセス終了後の RDP セッション操作（{@link RdpRemoteLauncherIni} / C# ランチャーと共有）。 */
public enum RdpSessionEndAction {
    NONE("なし"),
    DISCONNECT("切断"),
    SIGN_OUT("サインアウト");

    private final String iniValue;

    RdpSessionEndAction(String iniValue) {
        this.iniValue = iniValue;
    }

    public String iniValue() {
        return iniValue;
    }

    public String displayLabel() {
        return iniValue;
    }

    public boolean enabled() {
        return this != NONE;
    }

    public static RdpSessionEndAction fromIniValue(String raw, RdpSessionEndAction defaultValue) {
        if (raw == null || raw.isBlank()) {
            return defaultValue;
        }
        String normalized = raw.trim().toLowerCase(Locale.ROOT);
        return switch (normalized) {
            case "なし", "none", "off", "0" -> NONE;
            case "切断", "disconnect", "rdp切断", "rdp_disconnect" -> DISCONNECT;
            case "サインアウト", "signout", "sign_out", "logoff", "ログオフ" -> SIGN_OUT;
            default -> defaultValue;
        };
    }

    public static RdpSessionEndAction fromProfileJson(String raw) {
        return fromIniValue(raw, null);
    }
}
