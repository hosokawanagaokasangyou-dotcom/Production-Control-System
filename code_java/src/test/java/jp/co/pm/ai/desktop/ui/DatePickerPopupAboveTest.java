package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;

class DatePickerPopupAboveTest {

    @Test
    void computePopupY_placesAboveWhenThereIsRoom() {
        double fieldMinY = 400;
        double fieldMaxY = 428;
        double popupHeight = 240;
        double screenMinY = 0;
        double screenMaxY = 1080;

        assertEquals(
                160,
                DatePickerPopupAbove.computePopupY(
                        fieldMinY, fieldMaxY, popupHeight, screenMinY, screenMaxY));
    }

    @Test
    void computePopupY_fallsBackBelowWhenAboveWouldClipScreenTop() {
        double fieldMinY = 40;
        double fieldMaxY = 68;
        double popupHeight = 240;
        double screenMinY = 0;
        double screenMaxY = 1080;

        assertEquals(
                68,
                DatePickerPopupAbove.computePopupY(
                        fieldMinY, fieldMaxY, popupHeight, screenMinY, screenMaxY));
    }

    @Test
    void computePopupY_clampsAboveWhenNeitherSideFits() {
        double fieldMinY = 100;
        double fieldMaxY = 128;
        double popupHeight = 240;
        double screenMinY = 80;
        double screenMaxY = 200;

        assertEquals(
                80,
                DatePickerPopupAbove.computePopupY(
                        fieldMinY, fieldMaxY, popupHeight, screenMinY, screenMaxY));
    }

    @Test
    void computePopupY_keepsDefaultBelowWhenHeightUnknown() {
        assertEquals(
                428,
                DatePickerPopupAbove.computePopupY(400, 428, 0, 0, 1080));
    }

    @Test
    void requestFormAddFormField_installsAbovePlacement() throws Exception {
        Path java =
                Path.of(
                        "src/main/java/jp/co/pm/ai/desktop/reconciliation/ReconciliationApp.java");
        String text = Files.readString(java, StandardCharsets.UTF_8);
        int start = text.indexOf("private static void addFormField(");
        assertTrue(start >= 0, "addFormField が見つからない");
        int next = text.indexOf("\n    /** 左ペイン", start);
        String body = next > start ? text.substring(start, next) : text.substring(start, start + 800);
        assertTrue(body.contains("DatePickerPopupAbove.install"), "依頼書入力の DatePicker に上開きを入れる");
    }
}
