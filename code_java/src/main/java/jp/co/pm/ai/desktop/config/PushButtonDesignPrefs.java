package jp.co.pm.ai.desktop.config;

/**
 * 実行タブ・ダイアログなどのプッシュボタン見た目のユーザー上書き（セッション保存用）。
 *
 * <p>各 {@code customize*} が {@code true} のときだけ該当ブロックの生成 CSS がシーンに適用される。
 * プッシュボタン編集タブではスライダー／色を変えると該当グループのカスタムが自動でオンになる。
 */
public record PushButtonDesignPrefs(
        boolean customizeGeneralRunTab,
        double generalBorderRadius,
        double generalPaddingV,
        double generalPaddingH,
        double generalFontPx,
        String generalBgHex,
        String generalBorderHex,
        String generalTextHex,
        String generalHoverBgHex,
        String generalPressedBgHex,
        boolean customizeStageRunButtons,
        double stageFontPx,
        double stageMinWidth,
        double stageMinHeight,
        double stagePaddingV,
        double stagePaddingH,
        double stageBorderRadius,
        String stage1BgHex,
        String stage1BorderHex,
        String stage1HoverBgHex,
        String stage1PressedBgHex,
        String stage2BgHex,
        String stage2BorderHex,
        String stage2HoverBgHex,
        String stage2PressedBgHex,
        String stage3BgHex,
        String stage3BorderHex,
        String stage3HoverBgHex,
        String stage3PressedBgHex,
        boolean customizeDialogButtons,
        double dialogPrimaryBorderRadius,
        double dialogPrimaryPaddingV,
        double dialogPrimaryPaddingH,
        double dialogPrimaryFontPx,
        String dialogPrimaryBgHex,
        String dialogPrimaryBorderHex,
        String dialogPrimaryTextHex,
        String dialogPrimaryHoverBgHex,
        String dialogPrimaryPressedBgHex,
        double dialogSecondaryBorderRadius,
        double dialogSecondaryPaddingV,
        double dialogSecondaryPaddingH,
        double dialogSecondaryFontPx,
        String dialogSecondaryBgHex,
        String dialogSecondaryBorderHex,
        String dialogSecondaryTextHex,
        String dialogSecondaryHoverBgHex,
        String dialogSecondaryPressedBgHex) {

    /** 組み込み CSS に任せる（上書きスタイルシートを載せない）。 */
    public static PushButtonDesignPrefs inactiveDefaults() {
        BuiltIn b = BuiltIn.INSTANCE;
        return new PushButtonDesignPrefs(
                false,
                b.generalBorderRadius,
                b.generalPaddingV,
                b.generalPaddingH,
                b.generalFontPx,
                b.generalBgHex,
                b.generalBorderHex,
                b.generalTextHex,
                b.generalHoverBgHex,
                b.generalPressedBgHex,
                false,
                b.stageFontPx,
                b.stageMinWidth,
                b.stageMinHeight,
                b.stagePaddingV,
                b.stagePaddingH,
                b.stageBorderRadius,
                b.stage1BgHex,
                b.stage1BorderHex,
                b.stage1HoverBgHex,
                b.stage1PressedBgHex,
                b.stage2BgHex,
                b.stage2BorderHex,
                b.stage2HoverBgHex,
                b.stage2PressedBgHex,
                b.stage3BgHex,
                b.stage3BorderHex,
                b.stage3HoverBgHex,
                b.stage3PressedBgHex,
                false,
                b.dialogPrimaryBorderRadius,
                b.dialogPrimaryPaddingV,
                b.dialogPrimaryPaddingH,
                b.dialogPrimaryFontPx,
                b.dialogPrimaryBgHex,
                b.dialogPrimaryBorderHex,
                b.dialogPrimaryTextHex,
                b.dialogPrimaryHoverBgHex,
                b.dialogPrimaryPressedBgHex,
                b.dialogSecondaryBorderRadius,
                b.dialogSecondaryPaddingV,
                b.dialogSecondaryPaddingH,
                b.dialogSecondaryFontPx,
                b.dialogSecondaryBgHex,
                b.dialogSecondaryBorderHex,
                b.dialogSecondaryTextHex,
                b.dialogSecondaryHoverBgHex,
                b.dialogSecondaryPressedBgHex);
    }

    /** {@code pm-ai-desktop.css} に近い既定値（カスタム編集の初期値）。 */
    public static PushButtonDesignPrefs builtInSnapshot() {
        BuiltIn b = BuiltIn.INSTANCE;
        return new PushButtonDesignPrefs(
                true,
                b.generalBorderRadius,
                b.generalPaddingV,
                b.generalPaddingH,
                b.generalFontPx,
                b.generalBgHex,
                b.generalBorderHex,
                b.generalTextHex,
                b.generalHoverBgHex,
                b.generalPressedBgHex,
                true,
                b.stageFontPx,
                b.stageMinWidth,
                b.stageMinHeight,
                b.stagePaddingV,
                b.stagePaddingH,
                b.stageBorderRadius,
                b.stage1BgHex,
                b.stage1BorderHex,
                b.stage1HoverBgHex,
                b.stage1PressedBgHex,
                b.stage2BgHex,
                b.stage2BorderHex,
                b.stage2HoverBgHex,
                b.stage2PressedBgHex,
                b.stage3BgHex,
                b.stage3BorderHex,
                b.stage3HoverBgHex,
                b.stage3PressedBgHex,
                true,
                b.dialogPrimaryBorderRadius,
                b.dialogPrimaryPaddingV,
                b.dialogPrimaryPaddingH,
                b.dialogPrimaryFontPx,
                b.dialogPrimaryBgHex,
                b.dialogPrimaryBorderHex,
                b.dialogPrimaryTextHex,
                b.dialogPrimaryHoverBgHex,
                b.dialogPrimaryPressedBgHex,
                b.dialogSecondaryBorderRadius,
                b.dialogSecondaryPaddingV,
                b.dialogSecondaryPaddingH,
                b.dialogSecondaryFontPx,
                b.dialogSecondaryBgHex,
                b.dialogSecondaryBorderHex,
                b.dialogSecondaryTextHex,
                b.dialogSecondaryHoverBgHex,
                b.dialogSecondaryPressedBgHex);
    }

    public boolean anyCustomizationEnabled() {
        return customizeGeneralRunTab || customizeStageRunButtons || customizeDialogButtons;
    }

    private static final class BuiltIn {
        static final BuiltIn INSTANCE = new BuiltIn();

        final double generalBorderRadius = 6;
        final double generalPaddingV = 8;
        final double generalPaddingH = 14;
        final double generalFontPx = 12;
        final String generalBgHex = "#f4f4f4";
        final String generalBorderHex = "#c8c8c8";
        final String generalTextHex = "#2b2b2b";
        final String generalHoverBgHex = "#e8e8e8";
        final String generalPressedBgHex = "#dedede";

        final double stageFontPx = 15;
        final double stageMinWidth = 200;
        final double stageMinHeight = 48;
        final double stagePaddingV = 14;
        final double stagePaddingH = 28;
        final double stageBorderRadius = 8;

        final String stage1BgHex = "#0e7490";
        final String stage1BorderHex = "#155e75";
        final String stage1HoverBgHex = "#155e75";
        final String stage1PressedBgHex = "#164e63";

        final String stage2BgHex = "#c2410c";
        final String stage2BorderHex = "#9a3412";
        final String stage2HoverBgHex = "#9a3412";
        final String stage2PressedBgHex = "#7c2d12";

        final String stage3BgHex = "#15803d";
        final String stage3BorderHex = "#166534";
        final String stage3HoverBgHex = "#16a34a";
        final String stage3PressedBgHex = "#14532d";

        /** OK・既定（:default）向け。 */
        final double dialogPrimaryBorderRadius = 6;
        final double dialogPrimaryPaddingV = 8;
        final double dialogPrimaryPaddingH = 18;
        final double dialogPrimaryFontPx = 12;
        final String dialogPrimaryBgHex = "#2563eb";
        final String dialogPrimaryBorderHex = "#1d4ed8";
        final String dialogPrimaryTextHex = "#ffffff";
        final String dialogPrimaryHoverBgHex = "#3b82f6";
        final String dialogPrimaryPressedBgHex = "#1e40af";

        /** キャンセル等（:not(:default)）向け。 */
        final double dialogSecondaryBorderRadius = 6;
        final double dialogSecondaryPaddingV = 8;
        final double dialogSecondaryPaddingH = 14;
        final double dialogSecondaryFontPx = 12;
        final String dialogSecondaryBgHex = "#f4f4f4";
        final String dialogSecondaryBorderHex = "#c8c8c8";
        final String dialogSecondaryTextHex = "#2b2b2b";
        final String dialogSecondaryHoverBgHex = "#e8e8e8";
        final String dialogSecondaryPressedBgHex = "#dedede";

        private BuiltIn() {}
    }
}
