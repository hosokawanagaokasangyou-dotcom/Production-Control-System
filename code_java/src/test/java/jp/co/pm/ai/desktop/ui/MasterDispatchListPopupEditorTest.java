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
}
