package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TextField;

import org.controlsfx.control.spreadsheet.SpreadsheetCellEditor;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetView;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class MasterDispatchListPopupEditorTest {

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void dragBeyondThreshold_cancelsInlineEdit() {
        assertFalse(MasterDispatchListPopupEditor.isSelectionDrag(100, 100, 103, 101));
        assertTrue(MasterDispatchListPopupEditor.isSelectionDrag(100, 100, 100, 112));
        assertTrue(MasterDispatchListPopupEditor.isSelectionDrag(100, 100, 80, 100));
    }

    @Test
    void createEditor_usesCompactTextFieldNotComboBox() throws Exception {
        CountDownLatch done = new CountDownLatch(1);
        AtomicReference<SpreadsheetCellEditor> editorRef = new AtomicReference<>();
        Platform.runLater(
                () -> {
                    SpreadsheetView view = new SpreadsheetView();
                    MasterDispatchListCellType type =
                            new MasterDispatchListCellType(List.of("", "OP", "AS"));
                    editorRef.set(type.createEditor(view));
                    done.countDown();
                });
        assertTrue(done.await(5, TimeUnit.SECONDS));
        SpreadsheetCellEditor editor = editorRef.get();
        assertTrue(editor instanceof MasterDispatchListPopupEditor);
        assertTrue(editor.getEditor() instanceof TextField);
        assertFalse(editor.getEditor() instanceof ComboBox);
        assertEquals("", ((TextField) editor.getEditor()).getText());
    }

    @Test
    void popupEditorKeepsExactlyFixedHeightToAvoidRelayout() throws Exception {
        CountDownLatch done = new CountDownLatch(1);
        AtomicReference<TextField> fieldRef = new AtomicReference<>();
        Platform.runLater(
                () -> {
                    SpreadsheetCellEditor editor =
                            new MasterDispatchListCellType(List.of("", "OP", "AS"))
                                    .createEditor(new SpreadsheetView());
                    fieldRef.set((TextField) editor.getEditor());
                    done.countDown();
                });
        assertTrue(done.await(5, TimeUnit.SECONDS));
        TextField field = fieldRef.get();
        assertEquals(field.getPrefHeight(), field.getMinHeight());
        assertEquals(field.getPrefHeight(), field.getMaxHeight());
    }

    @Test
    void popupListCellCanBeRecognizedForSingleClickEditing() {
        SpreadsheetCell listCell =
                new MasterDispatchListCellType(List.of("", "OP", "AS"))
                        .createCell(1, 1, 1, 1, "");
        SpreadsheetCell plainCell = SpreadsheetCellType.STRING.createCell(1, 1, 1, 1, "");

        assertTrue(MasterDispatchListCellType.isPopupListCell(listCell));
        assertFalse(MasterDispatchListCellType.isPopupListCell(plainCell));
    }

    @Test
    void shouldStartListEditOnClick_onlyWhenSingleCellAndStillSincePress() {
        assertTrue(MasterDispatchSheetGridSupport.shouldStartListEditOnClick(1, true));
        assertFalse(MasterDispatchSheetGridSupport.shouldStartListEditOnClick(2, true));
        assertFalse(MasterDispatchSheetGridSupport.shouldStartListEditOnClick(9, true));
        assertFalse(MasterDispatchSheetGridSupport.shouldStartListEditOnClick(1, false));
        assertFalse(MasterDispatchSheetGridSupport.shouldStartListEditOnClick(0, true));
    }
}
