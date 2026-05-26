package jp.co.pm.ai.desktop.reconciliation;

/**
 * 受注ﾌｧｲﾙ行3で、{@link JuchuSheetColumnLayout.Col} に無い列位置の見出し。
 */
public record JuchuUnknownExcelColumn(
        String columnLetter, int columnIndex, String headerText, boolean ignored) {

    public String displayLabel() {
        return columnLetter + "列: " + headerText;
    }
}
