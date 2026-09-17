package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class AladdinReaderWarehouseTest {

    @Test
    @DisplayName("湖南は同一依頼NOでも国分倉庫の加工金額を合算しない")
    void konanKeepsOnlyKonanWarehouse() {
        List<List<Object>> rows = header();
        rows.add(List.of("511101", "国分工場01本倉庫", "049006", "W8-11", "加工金額", 105000));
        rows.add(List.of("520201", "湖南工場01本倉庫", "049006", "W8-11", "加工金額", 15976));
        AladdinData both = AladdinReader.read(rows, "049006", "test");
        assertEquals(120976.0, both.byIrai().get("W8-11"), 0.001, "倉庫未指定なら合算するのが現行");

        AladdinData konan = AladdinReader.read(rows, "049006", "湖南", "test");
        assertEquals(15976.0, konan.byIrai().get("W8-11"), 0.001);
        assertFalse(Math.abs(konan.byIrai().get("W8-11") - 120976.0) < 0.5);
    }

    @Test
    @DisplayName("倉庫名列が無い旧形式は倉庫絞り込みをしない")
    void missingWarehouseColumnDoesNotFilter() {
        List<List<Object>> rows = new ArrayList<>();
        rows.add(List.of("対象年月 : 2026年08月"));
        rows.add(List.of("依頼NO", "項目", "--合計--", "得意先"));
        rows.add(List.of("W8-11", "加工金額", 15976, "049006"));
        AladdinData data = AladdinReader.read(rows, "049006", "湖南", "test");
        assertEquals(15976.0, data.byIrai().get("W8-11"), 0.001);
    }

    private static List<List<Object>> header() {
        List<List<Object>> rows = new ArrayList<>();
        rows.add(List.of("対象年月 : 2026年08月"));
        rows.add(List.of("倉庫", "倉庫名", "得意先", "依頼NO", "項目", "--合計--"));
        return rows;
    }
}
