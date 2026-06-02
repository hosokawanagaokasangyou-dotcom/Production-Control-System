package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Test;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

class PlanInputDeprecatedOverrideColumnSupportTest {

    @Test
    void detectsOverrideAndReferenceColumns() {
        assertTrue(
                PlanInputDeprecatedOverrideColumnSupport.isDeprecatedOverrideColumn(
                        "原反投入日_上書き"));
        assertTrue(
                PlanInputDeprecatedOverrideColumnSupport.isOriginalReferenceColumn(
                        "（元）原反投入日_上書き"));
        assertTrue(
                PlanInputDeprecatedOverrideColumnSupport.isOriginalReferenceColumn(
                        "（元）担当OP_指定"));
        assertTrue(
                PlanInputDeprecatedOverrideColumnSupport.isDeprecatedReferenceOverrideColumn(
                        "（元）原反投入日_上書き"));
        assertFalse(
                PlanInputDeprecatedOverrideColumnSupport.isDeprecatedOverrideColumn(
                        "担当OP_指定"));
    }

    @Test
    void migrateAndDrop_removesOriginalReferenceColumns() {
        List<String> headers =
                new ArrayList<>(
                        List.of("依頼NO", "（元）担当OP_指定", "担当OP_指定", "（元）特別指定_備考", "特別指定_備考"));
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("A", "（OP1）", "OP1", "（備考）", "備考"));
        int dropped =
                PlanInputDeprecatedOverrideColumnSupport.migrateAndDropDeprecatedOverrideColumns(
                        headers, rows);
        assertEquals(2, dropped);
        assertEquals(List.of("依頼NO", "担当OP_指定", "特別指定_備考"), headers);
        assertEquals("OP1", rows.get(0).get(1));
        assertEquals("備考", rows.get(0).get(2));
    }

    @Test
    void migrateAndDrop_mergesOverrideIntoEmptyBase() {
        List<String> headers =
                new ArrayList<>(
                        List.of("依頼NO", "原反投入日", "原反投入日_上書き", "（元）原反投入日_上書き"));
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("A", "", "2026/6/14", "（2026/6/15）"));
        int dropped =
                PlanInputDeprecatedOverrideColumnSupport.migrateAndDropDeprecatedOverrideColumns(
                        headers, rows);
        assertEquals(2, dropped);
        assertEquals(List.of("依頼NO", "原反投入日"), headers);
        assertEquals("2026/6/14", rows.get(0).get(1));
    }

    @Test
    void migrateAndDrop_doesNotOverwriteNonemptyBase() {
        List<String> headers = new ArrayList<>(List.of("加工速度", "加工速度_上書き"));
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        rows.add(FXCollections.observableArrayList("10", "20"));
        PlanInputDeprecatedOverrideColumnSupport.migrateAndDropDeprecatedOverrideColumns(
                headers, rows);
        assertEquals(List.of("加工速度"), headers);
        assertEquals("10", rows.get(0).get(0));
    }
}
