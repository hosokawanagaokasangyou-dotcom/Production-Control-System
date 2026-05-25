package jp.co.pm.ai.desktop.print;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class EquipmentGanttPrintVerticalPaginationTest {

    @Test
    void paginate_singlePageWhenRowsFit() {
        List<List<Integer>> pages =
                EquipmentGanttPrintVerticalPagination.paginateLocalRowIndices(
                        5, i -> i == 0, 800, 8, 40, 22, 50);
        assertEquals(1, pages.size());
        assertEquals(List.of(0, 1, 2, 3, 4), pages.get(0));
    }

    @Test
    void paginate_splitsWhenDataRowsOverflow() {
        // available = 800 - 16 - 40 = 744; section 22 + 15*50 = 772 > 744 → need 2 pages
        List<List<Integer>> pages =
                EquipmentGanttPrintVerticalPagination.paginateLocalRowIndices(
                        16, i -> i == 0, 800, 8, 40, 22, 50);
        assertTrue(pages.size() >= 2);
        assertEquals(0, pages.get(0).get(0).intValue());
        int total = pages.stream().mapToInt(List::size).sum();
        assertEquals(16, total);
    }

    @Test
    void paginate_emptyRowCount() {
        assertTrue(
                EquipmentGanttPrintVerticalPagination.paginateLocalRowIndices(
                                0, i -> false, 800, 8, 40, 22, 50)
                        .isEmpty());
    }
}
