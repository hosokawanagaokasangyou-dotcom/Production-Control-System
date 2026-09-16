package jp.co.pm.ai.kouchin.verify;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class VerifyScopeTest {

    @Test
    @DisplayName("TPIの依頼NOと得意先049052は対象外")
    void excludesTpiByIraiAndCustomer() {
        assertTrue(VerifyScope.outOfScope("TPI07-01", ""));
        assertTrue(VerifyScope.outOfScope("TPI 1-1", ""));
        assertTrue(VerifyScope.outOfScope("C8-9", "049052"));
        assertFalse(VerifyScope.outOfScope("C8-9", "049006"));
        assertFalse(VerifyScope.outOfScope("T8-1", ""));
    }

    @Test
    @DisplayName("依頼NO先頭が2の自社加工は対象外")
    void excludesInHouseProcessingByLeadingTwo() {
        assertTrue(VerifyScope.outOfScope("2-15", ""));
        assertTrue(VerifyScope.outOfScope("201-3", ""));
        assertTrue(VerifyScope.outOfScope("２-1", ""));
        assertFalse(VerifyScope.outOfScope("C8-1", ""));
        assertFalse(VerifyScope.outOfScope("Y7-4", ""));
        assertFalse(VerifyScope.outOfScope("W7-12", ""));
    }

    @Test
    @DisplayName("③はTPI得意先と自社加工の依頼NOを集計しない")
    void aladdinSkipsTpiAndInHouse() {
        List<List<Object>> rows = new ArrayList<>();
        rows.add(List.of("対象年月 : 2026年07月"));
        rows.add(List.of("依頼NO", "項目", "--合計--", "得意先"));
        rows.add(List.of("C8-1", "加工金額", 10000, "049006"));
        rows.add(List.of("C8-9", "加工金額", 777, "049052"));
        rows.add(List.of("2-15", "加工金額", 500, "049006"));
        rows.add(List.of("TPI07-01", "加工金額", 300, "049052"));
        AladdinData data = AladdinReader.read(rows, null, "test");
        assertEquals(1, data.byIrai().size());
        assertEquals(10000.0, data.byIrai().get("C8-1"), 0.001);
        assertFalse(data.byIrai().containsKey("C8-9"));
        assertFalse(data.byIrai().containsKey("2-15"));
        assertFalse(data.byIrai().containsKey("TPI07-01"));
    }

    @Test
    @DisplayName("②東レまとめはTPIと自社加工の行を取り込まない")
    void nagaokaSkipsTpiAndInHouse() {
        List<List<Object>> rows = blankNagaokaHeader();
        rows.add(List.of("C", 8, "191352R", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "", "", "", "", "", "", "", "", "", "", 10000));
        rows.add(List.of("TPI", "07-01", "191999T", "", "", "", "", "", "", "", "", "", "", "", "",
                "", "", "", "", "", "", "", "", "", "", "", 800));
        rows.add(List.of("2", 15, "191888S", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "", "", "", "", "", "", "", "", "", "", 500));
        MoneyMaps maps = NagaokaReader.read(rows, "test.xlsx");
        assertEquals(java.util.Set.of("C-8"), maps.byIrai().keySet());
        assertEquals(10000.0, maps.irai("C-8"), 0.001);
        assertEquals(10000.0, maps.keiyaku("191352R"), 0.001);
        assertFalse(maps.byKeiyaku().containsKey("191999T"));
        assertFalse(maps.byKeiyaku().containsKey("191888S"));
    }

    private static List<List<Object>> blankNagaokaHeader() {
        List<List<Object>> rows = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            List<Object> r = new ArrayList<>();
            for (int j = 0; j < 27; j++) {
                r.add("");
            }
            rows.add(r);
        }
        rows.get(2).set(2, "契約No.");
        rows.get(3).set(26, "合計");
        return rows;
    }
}
