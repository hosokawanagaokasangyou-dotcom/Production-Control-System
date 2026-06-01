package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

/** {@link TableColumnOrderPersistence#alignVisibilityToHeaders} の見出し名ベース写像。 */
class TableColumnOrderPersistenceVisibilityTest {

    @Test
    void alignVisibilityToHeaders_mapsByTitleNotFileIndex() {
        List<String> savedOrder = List.of("A", "B", "C");
        boolean[] stored = {true, false, true};
        List<String> target = List.of("A", "B", "C");

        boolean[] aligned =
                TableColumnOrderPersistence.alignVisibilityToHeaders(savedOrder, stored, target);

        assertArrayEquals(new boolean[] {true, false, true}, aligned);
    }

    @Test
    void alignVisibilityToHeaders_survivesFileOrderDifferentFromSavedLayout() {
        List<String> savedOrder = List.of("A", "B", "C");
        boolean[] stored = {true, false, true};
        // ファイル列順が C,A,B でも、保存済みレイアウト順 A,B,C へ写す
        List<String> targetAfterReorder = List.of("A", "B", "C");

        boolean[] aligned =
                TableColumnOrderPersistence.alignVisibilityToHeaders(
                        savedOrder, stored, targetAfterReorder);

        assertArrayEquals(new boolean[] {true, false, true}, aligned);
    }

    @Test
    void alignVisibilityToHeaders_defaultsUnknownColumnsToVisible() {
        List<String> savedOrder = List.of("A", "B");
        boolean[] stored = {true, false};
        List<String> target = List.of("A", "B", "NEW");

        boolean[] aligned =
                TableColumnOrderPersistence.alignVisibilityToHeaders(savedOrder, stored, target);

        assertArrayEquals(new boolean[] {true, false, true}, aligned);
    }

    @Test
    void alignVisibilityToHeaders_indexFallbackWhenNoSavedOrder() {
        boolean[] stored = {false, true};
        List<String> target = List.of("X", "Y");

        boolean[] aligned =
                TableColumnOrderPersistence.alignVisibilityToHeaders(List.of(), stored, target);

        assertArrayEquals(new boolean[] {false, true}, aligned);
    }

    @Test
    void alignVisibilityToHeaders_droppedColumnsIgnored() {
        List<String> savedOrder = List.of("A", "B", "DROP");
        boolean[] stored = {true, false, false};
        List<String> target = List.of("A", "B");

        boolean[] aligned =
                TableColumnOrderPersistence.alignVisibilityToHeaders(savedOrder, stored, target);

        assertArrayEquals(new boolean[] {true, false}, aligned);
    }

    @Test
    void alignVisibilityToHeaders_duplicateTitlesUseFirstEntry() {
        List<String> savedOrder = List.of("A", "A", "B");
        boolean[] stored = {false, true, true};
        List<String> target = List.of("A", "B");

        boolean[] aligned =
                TableColumnOrderPersistence.alignVisibilityToHeaders(savedOrder, stored, target);

        assertTrue(!aligned[0]);
        assertTrue(aligned[1]);
    }
}
