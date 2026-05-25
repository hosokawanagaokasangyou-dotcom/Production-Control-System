package jp.co.pm.ai.desktop.print;

import java.util.ArrayList;
import java.util.List;
import java.util.function.IntPredicate;

/**
 * 設備ガント印刷: 1 暦日分の行が用紙 1 枚の縦幅に収まらないとき、行インデックス束を複数ページに分割する。
 *
 * <p>各ページは {@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane#buildDedicatedPrintSheet} で
 * 見出し（機械名・工程名・時刻軸＋暦日）付きの独立シートとして組み立てる。
 */
public final class EquipmentGanttPrintVerticalPagination {

    private EquipmentGanttPrintVerticalPagination() {}

    /**
     * {@code 0 .. rowCount-1} の行を、1 ページあたりの縦寸法上限に収まる連続チャンクに分ける。
     *
     * @param rowCount 行数（{@code spec.rows()} の size）
     * @param isSectionRow 行インデックスが日付バナー等のセクション行なら {@code true}
     * @param paperHeightPx 可印刷高さ（pt）
     * @param pad 上下パディング（片側）
     * @param headerH 列見出し＋時刻軸の高さ
     * @param sectionH セクション行 1 行の高さ
     * @param dataRowH データ行 1 行の高さ（固定読み取り高）
     */
    public static List<List<Integer>> paginateLocalRowIndices(
            int rowCount,
            IntPredicate isSectionRow,
            double paperHeightPx,
            double pad,
            double headerH,
            double sectionH,
            double dataRowH) {
        if (rowCount <= 0) {
            return List.of();
        }
        double available = paperHeightPx - 2 * pad - headerH;
        if (!Double.isFinite(available) || available < sectionH) {
            return List.of(allIndices(rowCount));
        }
        double secH = Math.max(1, sectionH);
        double dataH = Math.max(1, dataRowH);

        List<List<Integer>> pages = new ArrayList<>();
        List<Integer> page = new ArrayList<>();
        double used = 0;
        for (int i = 0; i < rowCount; i++) {
            double rowH = isSectionRow.test(i) ? secH : dataH;
            if (!page.isEmpty() && used + rowH > available + 0.5) {
                pages.add(page);
                page = new ArrayList<>();
                used = 0;
            }
            page.add(i);
            used += rowH;
        }
        if (!page.isEmpty()) {
            pages.add(page);
        }
        return pages.isEmpty() ? List.of(allIndices(rowCount)) : pages;
    }

    private static List<Integer> allIndices(int rowCount) {
        List<Integer> all = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            all.add(i);
        }
        return all;
    }
}
