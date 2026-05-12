package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Locale;
import java.util.Objects;

import javafx.scene.Scene;

/**
 * {@link PushButtonDesignPrefs} から CSS を生成し、ユーザー設定ディレクトリに書き出してシーンへ適用する。
 */
public final class PushButtonCssEmitter {

    private static final String FILE_NAME = "push-button-overrides.css";

    private PushButtonCssEmitter() {}

    /** 工場出荷 UI リセット時にユーザー CSS を削除する。 */
    public static void deleteUserOverridesFileSilently() {
        try {
            Files.deleteIfExists(resolveStoreFile());
        } catch (IOException ignored) {
        }
    }

    /**
     * 既存の上書きシートを外し、必要なら生成した CSS を最後に追加する（{@code pm-ai-desktop.css} より優先）。
     */
    public static void applyToScene(Scene scene, PushButtonDesignPrefs prefs) {
        if (scene == null) {
            return;
        }
        PushButtonDesignPrefs p = prefs != null ? prefs : PushButtonDesignPrefs.inactiveDefaults();
        scene.getStylesheets().removeIf(PushButtonCssEmitter::isOurOverrideUrl);
        if (!p.anyCustomizationEnabled()) {
            return;
        }
        String css = emitCss(p);
        if (css.isBlank()) {
            return;
        }
        try {
            Path file = resolveStoreFile();
            Files.createDirectories(file.getParent());
            Files.writeString(file, css, StandardCharsets.UTF_8);
            scene.getStylesheets().add(file.toUri().toString());
        } catch (IOException ignored) {
        }
    }

    static boolean isOurOverrideUrl(String url) {
        return url != null && url.contains(FILE_NAME);
    }

    static Path resolveStoreFile() {
        return Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", FILE_NAME);
    }

    static String emitCss(PushButtonDesignPrefs p) {
        Objects.requireNonNull(p);
        StringBuilder sb = new StringBuilder();
        sb.append("/* User push-button overrides (generated; do not edit by hand) */\n");
        if (p.customizeGeneralRunTab()) {
            appendGeneral(sb, p);
        }
        if (p.customizeStageRunButtons()) {
            appendStageShared(sb, p);
            appendStageVariant(sb, 1, p.stage1BgHex(), p.stage1BorderHex(), p.stage1HoverBgHex(), p.stage1PressedBgHex());
            appendStageVariant(sb, 2, p.stage2BgHex(), p.stage2BorderHex(), p.stage2HoverBgHex(), p.stage2PressedBgHex());
            appendStageVariant(sb, 3, p.stage3BgHex(), p.stage3BorderHex(), p.stage3HoverBgHex(), p.stage3PressedBgHex());
            sb.append(".button.pm-stage-run-button:disabled {\n");
            sb.append("    -fx-opacity: 0.5;\n");
            sb.append("}\n");
        }
        if (p.customizeDialogButtons()) {
            appendDialogButtons(sb, p);
        }
        return sb.toString();
    }

    /**
     * {@link javafx.scene.control.DialogPane} 内のボタンに加え、メインシェル「タブの並びとグループ」画面（{@code
     * pm-main-shell-tab-organizer}）フッターの「構成を適用」（既定）なども同一の主／副スタイルで上書きする。
     */
    private static void appendDialogButtons(StringBuilder sb, PushButtonDesignPrefs p) {
        String[] scopes = {".dialog-pane", ".pm-main-shell-tab-organizer"};
        for (String scope : scopes) {
            appendDialogVariant(
                    sb,
                    scope + " .button:default",
                    p.dialogPrimaryBorderRadius(),
                    p.dialogPrimaryPaddingV(),
                    p.dialogPrimaryPaddingH(),
                    p.dialogPrimaryFontPx(),
                    p.dialogPrimaryBgHex(),
                    p.dialogPrimaryBorderHex(),
                    p.dialogPrimaryTextHex(),
                    p.dialogPrimaryHoverBgHex(),
                    p.dialogPrimaryPressedBgHex());
            appendDialogVariant(
                    sb,
                    scope + " .button:not(:default)",
                    p.dialogSecondaryBorderRadius(),
                    p.dialogSecondaryPaddingV(),
                    p.dialogSecondaryPaddingH(),
                    p.dialogSecondaryFontPx(),
                    p.dialogSecondaryBgHex(),
                    p.dialogSecondaryBorderHex(),
                    p.dialogSecondaryTextHex(),
                    p.dialogSecondaryHoverBgHex(),
                    p.dialogSecondaryPressedBgHex());
        }
        for (int i = 0; i < scopes.length; i++) {
            String scope = scopes[i];
            sb.append(scope).append(" .button:default:focused,\n");
            sb.append(scope).append(" .button:not(:default):focused");
            if (i < scopes.length - 1) {
                sb.append(",\n");
            } else {
                sb.append(" {\n");
            }
        }
        sb.append("    -fx-focus-color: transparent;\n");
        sb.append("    -fx-faint-focus-color: transparent;\n");
        sb.append("}\n");
    }

    private static void appendDialogVariant(
            StringBuilder sb,
            String selector,
            double radius,
            double padV,
            double padH,
            double fontPx,
            String bg,
            String border,
            String text,
            String hover,
            String pressed) {
        sb.append(selector)
                .append(" {\n")
                .append("    -fx-cursor: hand;\n")
                .append("    -fx-background-color: ")
                .append(hexOrFallback(bg, "#808080"))
                .append(";\n")
                .append("    -fx-border-color: ")
                .append(hexOrFallback(border, "#606060"))
                .append(";\n")
                .append("    -fx-border-width: 1;\n")
                .append("    -fx-border-radius: ")
                .append(fmtPx(radius, 6))
                .append(";\n")
                .append("    -fx-background-radius: ")
                .append(fmtPx(radius, 6))
                .append(";\n")
                .append("    -fx-padding: ")
                .append(fmtPx(padV, 8))
                .append(" ")
                .append(fmtPx(padH, 14))
                .append(";\n")
                .append("    -fx-font-size: ")
                .append(fmtPx(fontPx, 12))
                .append(";\n")
                .append("    -fx-text-fill: ")
                .append(hexOrFallback(text, "#202020"))
                .append(";\n")
                .append("}\n");
        sb.append(selector)
                .append(":hover {\n")
                .append("    -fx-background-color: ")
                .append(hexOrFallback(hover, hexOrFallback(bg, "#808080")))
                .append(";\n")
                .append("}\n");
        sb.append(selector)
                .append(":pressed {\n")
                .append("    -fx-background-color: ")
                .append(hexOrFallback(pressed, hexOrFallback(hover, hexOrFallback(bg, "#808080"))))
                .append(";\n")
                .append("}\n");
    }

    private static void appendGeneral(StringBuilder sb, PushButtonDesignPrefs p) {
        sb.append(".pm-run-tab .button:not(.pm-stage-run-button) {\n");
        sb.append("    -fx-cursor: hand;\n");
        sb.append("    -fx-background-color: ")
                .append(hexOrFallback(p.generalBgHex(), "#f4f4f4"))
                .append(";\n");
        sb.append("    -fx-border-color: ")
                .append(hexOrFallback(p.generalBorderHex(), "#c8c8c8"))
                .append(";\n");
        sb.append("    -fx-border-width: 1;\n");
        sb.append("    -fx-border-radius: ")
                .append(fmtPx(p.generalBorderRadius(), 6))
                .append(";\n");
        sb.append("    -fx-background-radius: ")
                .append(fmtPx(p.generalBorderRadius(), 6))
                .append(";\n");
        sb.append("    -fx-padding: ")
                .append(fmtPx(p.generalPaddingV(), 8))
                .append(" ")
                .append(fmtPx(p.generalPaddingH(), 14))
                .append(";\n");
        sb.append("    -fx-font-size: ")
                .append(fmtPx(p.generalFontPx(), 12))
                .append(";\n");
        sb.append("    -fx-text-fill: ")
                .append(hexOrFallback(p.generalTextHex(), "#2b2b2b"))
                .append(";\n");
        sb.append("}\n");
        sb.append(".pm-run-tab .button:not(.pm-stage-run-button):hover {\n");
        sb.append("    -fx-background-color: ")
                .append(hexOrFallback(p.generalHoverBgHex(), "#e8e8e8"))
                .append(";\n");
        sb.append("}\n");
        sb.append(".pm-run-tab .button:not(.pm-stage-run-button):pressed {\n");
        sb.append("    -fx-background-color: ")
                .append(hexOrFallback(p.generalPressedBgHex(), "#dedede"))
                .append(";\n");
        sb.append("}\n");
    }

    private static void appendStageShared(StringBuilder sb, PushButtonDesignPrefs p) {
        sb.append(".button.pm-stage-run-button {\n");
        sb.append("    -fx-cursor: hand;\n");
        sb.append("    -fx-font-size: ")
                .append(fmtPx(p.stageFontPx(), 15))
                .append(";\n");
        sb.append("    -fx-font-weight: bold;\n");
        sb.append("    -fx-padding: ")
                .append(fmtPx(p.stagePaddingV(), 14))
                .append(" ")
                .append(fmtPx(p.stagePaddingH(), 28))
                .append(";\n");
        sb.append("    -fx-background-radius: ")
                .append(fmtPx(p.stageBorderRadius(), 8))
                .append(";\n");
        sb.append("    -fx-border-radius: ")
                .append(fmtPx(p.stageBorderRadius(), 8))
                .append(";\n");
        sb.append("    -fx-border-width: 1;\n");
        sb.append("    -fx-min-height: ")
                .append(fmtPx(p.stageMinHeight(), 48))
                .append(";\n");
        sb.append("    -fx-min-width: ")
                .append(fmtPx(p.stageMinWidth(), 200))
                .append(";\n");
        sb.append("    -fx-text-fill: #ffffff;\n");
        sb.append("}\n");
        sb.append(".button.pm-stage-run-button:focused {\n");
        sb.append("    -fx-focus-color: transparent;\n");
        sb.append("    -fx-faint-focus-color: transparent;\n");
        sb.append("}\n");
    }

    private static void appendStageVariant(
            StringBuilder sb,
            int stage,
            String bg,
            String border,
            String hover,
            String pressed) {
        String cls = ".button.pm-stage-run-button.pm-stage-run-" + stage;
        sb.append(cls)
                .append(" {\n")
                .append("    -fx-background-color: ")
                .append(hexOrFallback(bg, "#808080"))
                .append(";\n")
                .append("    -fx-border-color: ")
                .append(hexOrFallback(border, "#606060"))
                .append(";\n")
                .append("}\n");
        sb.append(cls)
                .append(":hover {\n")
                .append("    -fx-background-color: ")
                .append(hexOrFallback(hover, hexOrFallback(bg, "#808080")))
                .append(";\n")
                .append("}\n");
        sb.append(cls)
                .append(":pressed {\n")
                .append("    -fx-background-color: ")
                .append(hexOrFallback(pressed, hexOrFallback(hover, hexOrFallback(bg, "#808080"))))
                .append(";\n")
                .append("}\n");
    }

    private static String hexOrFallback(String hex, String fallback) {
        String h = normalizeHex(hex);
        return h != null ? h : fallback;
    }

    private static String normalizeHex(String hex) {
        if (hex == null) {
            return null;
        }
        String t = hex.strip();
        if (t.isEmpty()) {
            return null;
        }
        if (!t.startsWith("#")) {
            t = "#" + t;
        }
        if (t.length() == 9 && t.startsWith("#")) {
            return t.substring(0, 7);
        }
        if (t.length() == 7) {
            return t;
        }
        return null;
    }

    private static String fmtPx(double v, double fallback) {
        double x = v > 0 && Double.isFinite(v) ? v : fallback;
        if (x <= 0 || !Double.isFinite(x)) {
            x = fallback;
        }
        return String.format(Locale.US, "%.1fpx", x);
    }
}
