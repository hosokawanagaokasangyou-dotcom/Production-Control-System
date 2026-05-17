package jp.co.pm.ai.desktop.io;

import jp.co.pm.ai.desktop.config.DesktopTheme;

/** Excel セルスタイル用のテーマ別 RGB（サマリ AI 配台出力）。 */
record SummaryAiDispatchExportThemePalette(
        byte[] headerFillRgb,
        byte[] headerFontRgb,
        byte[] dataFillRgb,
        byte[] dataFontRgb) {

    private static final byte[] LIGHT_HEADER = new byte[] {(byte) 0xFF, (byte) 0xF2, (byte) 0xCC};
    private static final byte[] LIGHT_DATA = new byte[] {(byte) 0xFF, (byte) 0xFF, (byte) 0xFF};
    private static final byte[] DARK_HEADER = new byte[] {(byte) 0x2D, (byte) 0x3A, (byte) 0x4F};
    private static final byte[] DARK_DATA = new byte[] {(byte) 0x1E, (byte) 0x26, (byte) 0x32};
    private static final byte[] BLACK = new byte[] {0, 0, 0};
    private static final byte[] WHITE = new byte[] {(byte) 0xFF, (byte) 0xFF, (byte) 0xFF};
    private static final byte[] WARM_HEADER = new byte[] {(byte) 0xF5, (byte) 0xE6, (byte) 0xD3};
    private static final byte[] BLUE_HEADER = new byte[] {(byte) 0xD6, (byte) 0xE4, (byte) 0xF0};

    static SummaryAiDispatchExportThemePalette forTheme(DesktopTheme theme) {
        DesktopTheme t = theme != null ? theme : DesktopTheme.LIGHT;
        return switch (t) {
            case DARK,
                    MIDNIGHT,
                    SLATE,
                    MIDNIGHT_BLUE,
                    DEEP_NAVY,
                    SPACE_GRAY,
                    CHARCOAL,
                    INK_BLACK,
                    PRUSSIAN_BLUE,
                    GUNMETAL,
                    NORD,
                    DRACULA,
                    AUBERGINE,
                    OCEANIC,
                    DEEP_ASTRAL,
                    MATERIAL_DARK,
                    OLED_BLACK,
                    NIGHT_MOSS,
                    STEEL_GRAY,
                    OXFORD_BLUE,
                    EMBER,
                    OCEAN -> new SummaryAiDispatchExportThemePalette(DARK_HEADER, WHITE, DARK_DATA, WHITE);
            case SEPIA -> new SummaryAiDispatchExportThemePalette(WARM_HEADER, BLACK, LIGHT_DATA, BLACK);
            case BLUE, CONTRAST -> new SummaryAiDispatchExportThemePalette(BLUE_HEADER, BLACK, LIGHT_DATA, BLACK);
            default -> new SummaryAiDispatchExportThemePalette(LIGHT_HEADER, BLACK, LIGHT_DATA, BLACK);
        };
    }
}
