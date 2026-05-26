package jp.co.pm.ai.desktop.reconciliation;

/**
 * 受注ﾌｧｲﾙ 行3の見出しと {@link JuchuSheetColumnLayout.Col} 定義の不一致1件。
 */
public record JuchuHeaderMismatch(
        JuchuSheetColumnLayout.Col column,
        String expectedHeader,
        String actualHeader,
        boolean actualEmpty) {

    public String columnLetter() {
        return column.columnLetter();
    }

    /** 依頼書手修正フォーム上の対応項目。 */
    public String formItemDescription() {
        return column.formItemDescription();
    }

    public String summaryLine() {
        String prefix = formItemDescription() + "（" + columnLetter() + "列）";
        if (actualEmpty) {
            return prefix + ": 見出しが空です（期待: " + expectedHeader + "）";
        }
        return prefix
                + ": 期待「"
                + expectedHeader
                + "」だが実際「"
                + actualHeader
                + "」";
    }
}
