package jp.co.pm.ai.desktop.ui;

import javafx.application.Platform;
import javafx.geometry.Bounds;
import javafx.geometry.Insets;
import javafx.scene.control.Button;
import javafx.scene.control.Control;
import javafx.scene.control.TextField;
import javafx.scene.control.TextFormatter;
import javafx.scene.input.KeyCode;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Popup;

import org.controlsfx.control.spreadsheet.SpreadsheetCellEditor;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * セル内はキーボード入力用 TextField。テンキーは Popup（セル高さにクリップされない）。
 */
public final class MasterDispatchDecimalKeypadEditor extends SpreadsheetCellEditor {

    private final double min;
    private final double max;
    private final int fractionDigits;
    private final TextField field = new TextField();
    private final VBox keypadBox = new VBox(6);
    private final Popup popup = new Popup();
    private boolean editing;

    public MasterDispatchDecimalKeypadEditor(
            SpreadsheetView view, double min, double max, int fractionDigits) {
        super(view);
        this.min = min;
        this.max = max;
        this.fractionDigits = Math.max(0, fractionDigits);
        field.getStyleClass().add("pm-ai-decimal-editor-field");
        field.setPrefColumnCount(8);
        field.setTextFormatter(
                new TextFormatter<>(
                        change -> {
                            String next = change.getControlNewText();
                            if (MasterDispatchSheetEditRules.isDecimalTypingAllowed(
                                    next, min, max, this.fractionDigits)) {
                                return change;
                            }
                            return null;
                        }));
        field.setOnAction(e -> tryCommit());
        field.setOnKeyPressed(
                e -> {
                    if (e.getCode() == KeyCode.ESCAPE) {
                        endEdit(false);
                        e.consume();
                    }
                });
        GridPane pad = buildPad();
        keypadBox.getStyleClass().add("pm-ai-decimal-keypad");
        keypadBox.setPadding(new Insets(6));
        keypadBox.getChildren().add(pad);
        var cssUrl =
                SpreadsheetTabularSupport.class.getResource(
                        "/jp/co/pm/ai/desktop/css/delivery-calendar-spreadsheet.css");
        if (cssUrl != null) {
            String css = cssUrl.toExternalForm();
            if (!keypadBox.getStylesheets().contains(css)) {
                keypadBox.getStylesheets().add(css);
            }
        }
        popup.getContent().add(keypadBox);
        popup.setAutoHide(true);
        popup.setHideOnEscape(true);
        popup.setAutoFix(true);
        popup.setOnAutoHide(
                e -> {
                    if (editing) {
                        endEdit(true);
                    }
                });
    }

    private GridPane buildPad() {
        GridPane pad = new GridPane();
        pad.setHgap(4);
        pad.setVgap(4);
        pad.getStyleClass().add("pm-ai-decimal-keypad-grid");
        for (int i = 0; i < 4; i++) {
            ColumnConstraints cc = new ColumnConstraints();
            cc.setPercentWidth(25);
            cc.setHgrow(Priority.ALWAYS);
            pad.getColumnConstraints().add(cc);
        }
        String[][] keys = {
            {"7", "8", "9", "←"},
            {"4", "5", "6", "C"},
            {"1", "2", "3", "確定"},
            {"0", fractionDigits > 0 ? "." : "", "", ""}
        };
        for (int r = 0; r < keys.length; r++) {
            for (int c = 0; c < keys[r].length; c++) {
                String label = keys[r][c];
                if (label.isEmpty()) {
                    continue;
                }
                Button b = new Button(label);
                b.setMaxWidth(Double.MAX_VALUE);
                b.setFocusTraversable(false);
                b.getStyleClass().add("pm-ai-decimal-keypad-key");
                b.addEventFilter(MouseEvent.MOUSE_PRESSED, e -> field.requestFocus());
                if ("確定".equals(label)) {
                    b.getStyleClass().add("pm-ai-decimal-keypad-ok");
                    b.setMaxHeight(Double.MAX_VALUE);
                    pad.add(b, c, r, 1, 2);
                } else {
                    pad.add(b, c, r);
                }
                b.setOnAction(e -> onKey(label));
            }
        }
        return pad;
    }

    private void onKey(String label) {
        switch (label) {
            case "←" -> {
                String t = field.getText() != null ? field.getText() : "";
                if (!t.isEmpty()) {
                    field.setText(t.substring(0, t.length() - 1));
                }
            }
            case "C" -> field.clear();
            case "確定" -> tryCommit();
            default -> {
                String cur = field.getText() != null ? field.getText() : "";
                String next = cur + label;
                if (MasterDispatchSheetEditRules.isDecimalTypingAllowed(
                        next, min, max, fractionDigits)) {
                    field.setText(next);
                    markValid();
                } else {
                    markInvalid();
                }
            }
        }
        field.requestFocus();
        field.end();
    }

    private void tryCommit() {
        String t = field.getText() != null ? field.getText().strip() : "";
        if (t.isEmpty()
                || MasterDispatchSheetEditRules.isDecimalInRange(t, min, max, fractionDigits)) {
            markValid();
            endEdit(true);
            return;
        }
        markInvalid();
    }

    private void markInvalid() {
        if (!field.getStyleClass().contains("pm-ai-decimal-editor-invalid")) {
            field.getStyleClass().add("pm-ai-decimal-editor-invalid");
        }
    }

    private void markValid() {
        field.getStyleClass().remove("pm-ai-decimal-editor-invalid");
    }

    private void showKeypad() {
        if (!editing || popup.isShowing()) {
            return;
        }
        Bounds b = field.localToScreen(field.getBoundsInLocal());
        if (b == null) {
            return;
        }
        popup.show(field, b.getMinX(), b.getMaxY());
        field.requestFocus();
        field.end();
    }

    @Override
    public void startEdit(Object value, String format, Object... options) {
        editing = true;
        markValid();
        String s = value == null ? "" : value.toString();
        field.setText(s);
        field.requestFocus();
        field.end();
        Platform.runLater(this::showKeypad);
    }

    @Override
    public String getControlValue() {
        String t = field.getText() != null ? field.getText().strip() : "";
        if (t.isEmpty()) {
            return "";
        }
        if (t.endsWith(".")) {
            t = t.substring(0, t.length() - 1);
        }
        if (MasterDispatchSheetEditRules.isDecimalInRange(t, min, max, fractionDigits)) {
            return MasterDispatchSheetEditRules.formatDecimal(t, fractionDigits);
        }
        return t;
    }

    @Override
    public void end() {
        editing = false;
        popup.hide();
        markValid();
    }

    @Override
    public Control getEditor() {
        return field;
    }
}
