package jp.co.pm.ai.desktop.reconciliation;

import org.apache.poi.ss.usermodel.VerticalAlignment;

/** セルまたはシェイプ内テキスト 1 ラン分の書式。 */
record RequestFormPreviewCellStyle(
        double fontSizePx,
        String foreground,
        String background,
        boolean bold,
        boolean italic,
        boolean underline,
        boolean doubleUnderline,
        boolean strike,
        boolean doubleStrike,
        String fontFamily,
        VerticalAlignment verticalAlignment,
        boolean wrapText,
        String borderCss) {

    static RequestFormPreviewCellStyle defaults() {
        return new RequestFormPreviewCellStyle(
                11.0 * 96.0 / 72.0,
                "#000000",
                "#FFFFFF",
                false,
                false,
                false,
                false,
                false,
                false,
                null,
                VerticalAlignment.CENTER,
                false,
                "");
    }
}
