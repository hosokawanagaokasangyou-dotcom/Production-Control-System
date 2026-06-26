package jp.co.pm.ai.desktop.reconciliation;

/**
 * 受注ﾌｧｲﾙ 見出しと {@link JuchuSheetColumnLayout.Col} 定義の不一致1件。
 * {@code transferColumnLetter} は採用列（転記・読込の物理列）。未指定時は {@link Col} 既定。
 */
public record JuchuHeaderMismatch(
        JuchuSheetColumnLayout.Col column,
        String expectedHeader,
        String actualHeader,
        boolean actualEmpty,
        String transferColumnLetter) {

    public JuchuHeaderMismatch(
            JuchuSheetColumnLayout.Col column,
            String expectedHeader,
            String actualHeader,
            boolean actualEmpty) {
        this(column, expectedHeader, actualHeader, actualEmpty, column.columnLetter());
    }

    public String columnLetter() {
        return transferColumnLetter != null && !transferColumnLetter.isBlank()
                ? transferColumnLetter
                : column.columnLetter();
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
