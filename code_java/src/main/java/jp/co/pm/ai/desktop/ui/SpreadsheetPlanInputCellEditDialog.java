package jp.co.pm.ai.desktop.ui;

import java.util.Optional;

import javafx.application.Platform;
import javafx.geometry.Rectangle2D;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextInputControl;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Screen;
import javafx.stage.Window;

/**
 * Modal editor for a single plan-input spreadsheet cell (near-click placement, width from column).
 */
public final class SpreadsheetPlanInputCellEditDialog {

    private SpreadsheetPlanInputCellEditDialog() {}

    /**
     * Shows a small dialog near {@code anchorScreenX/Y} for editing one cell value.
     *
     * @param columnWidthHint column width in px (from {@code TableColumn#getWidth()}), or 0 to use default width
     */
    public static Optional<String> edit(
            Window owner,
            String columnTitle,
            String initialValue,
            double columnWidthHint,
            double anchorScreenX,
            double anchorScreenY) {
        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.WINDOW_MODAL);
        String title =
                columnTitle != null && !columnTitle.isBlank()
                        ? columnTitle.strip()
                        : "セルの編集";
        dialog.setTitle(title);
        dialog.setHeaderText(null);

        TextArea area = new TextArea(initialValue != null ? initialValue : "");
        area.setWrapText(true);
        int lineCount = Math.max(1, initialValue == null ? 1 : initialValue.split("\n", -1).length);
        int prefRows = Math.max(3, Math.min(18, lineCount + 2));
        area.setPrefRowCount(prefRows);

        double w = Math.min(780, Math.max(300, columnWidthHint <= 0 ? 420 : columnWidthHint * 1.2 + 56));
        area.setPrefWidth(w);

        Label hint =
                new Label(
                        columnTitle != null && !columnTitle.isBlank()
                                ? "列: " + columnTitle.strip()
                                : "セル値を編集してください");
        hint.setStyle("-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 18%);");

        VBox box = new VBox(10, hint, area);
        VBox.setVgrow(area, Priority.ALWAYS);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.getDialogPane().setPrefWidth(w + 40);

        dialog.setOnShown(
                e ->
                        Platform.runLater(
                                () -> {
                                    positionNearAnchor(
                                            dialog.getDialogPane().getScene().getWindow(),
                                            anchorScreenX,
                                            anchorScreenY);
                                    focusForImmediateEdit(area);
                                }));

        Optional<ButtonType> r = dialog.showAndWait();
        if (r.isPresent() && r.get() == ButtonType.OK) {
            return Optional.of(area.getText());
        }
        return Optional.empty();
    }

    /**
     * 表示専用セル向け: 値を折り返し表示するだけのダイアログ（編集不可・閉じるのみ）。
     *
     * @param note 何によって埋まる列かの説明（{@code null} 可）
     */
    public static void viewReadOnly(
            Window owner,
            String columnTitle,
            String value,
            String note,
            double columnWidthHint,
            double anchorScreenX,
            double anchorScreenY) {
        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.WINDOW_MODAL);
        String title =
                columnTitle != null && !columnTitle.isBlank() ? columnTitle.strip() : "セルの内容";
        dialog.setTitle(title);
        dialog.setHeaderText(null);

        String shown = value != null ? value : "";
        TextArea area = new TextArea(shown);
        area.setWrapText(true);
        area.setEditable(false);
        area.setPrefRowCount(Math.max(3, Math.min(18, shown.length() / 60 + 3)));
        double w = Math.min(780, Math.max(360, columnWidthHint <= 0 ? 480 : columnWidthHint * 1.2 + 56));
        area.setPrefWidth(w);

        Label hint = new Label(note != null && !note.isBlank() ? note.strip() : "この列は編集できません。");
        hint.setWrapText(true);
        hint.setStyle("-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 18%);");

        VBox box = new VBox(10, hint, area);
        VBox.setVgrow(area, Priority.ALWAYS);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.CLOSE);
        dialog.getDialogPane().setPrefWidth(w + 40);

        dialog.setOnShown(
                e ->
                        Platform.runLater(
                                () ->
                                        positionNearAnchor(
                                                dialog.getDialogPane().getScene().getWindow(),
                                                anchorScreenX,
                                                anchorScreenY)));

        dialog.showAndWait();
    }

    /**
     * 日付列向け: {@link DatePicker} で暦日を選ぶ（空にする「クリア」付き）。
     */
    public static Optional<String> editDate(
            Window owner,
            String columnTitle,
            String initialValue,
            double anchorScreenX,
            double anchorScreenY) {
        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.WINDOW_MODAL);
        String title =
                columnTitle != null && !columnTitle.isBlank()
                        ? columnTitle.strip()
                        : "日付の選択";
        dialog.setTitle(title);
        dialog.setHeaderText(null);

        DatePicker picker = new DatePicker();
        PlanInputDateColumnSupport.parseCellValue(initialValue).ifPresent(picker::setValue);

        Button clearButton = new Button("クリア");
        clearButton.setOnAction(ev -> picker.setValue(null));

        Label hint =
                new Label(
                        columnTitle != null && !columnTitle.isBlank()
                                ? "列: " + columnTitle.strip()
                                : "日付を選択してください");
        hint.setStyle("-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 18%);");

        HBox pickerRow = new HBox(8, picker, clearButton);
        pickerRow.setAlignment(javafx.geometry.Pos.CENTER_LEFT);

        VBox box = new VBox(10, hint, pickerRow);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.getDialogPane().setPrefWidth(320);

        dialog.setOnShown(
                e ->
                        Platform.runLater(
                                () -> {
                                    positionNearAnchor(
                                            dialog.getDialogPane().getScene().getWindow(),
                                            anchorScreenX,
                                            anchorScreenY);
                                    picker.requestFocus();
                                }));

        Optional<ButtonType> r = dialog.showAndWait();
        if (r.isPresent() && r.get() == ButtonType.OK) {
            return Optional.of(PlanInputDateColumnSupport.formatCellValue(picker.getValue()));
        }
        return Optional.empty();
    }

    /**
     * 日時列向け（配台可能日時 / 配台可能日時_上書き）: 暦日 + 時刻（HH:mm、既定 12:45）。
     */
    public static Optional<String> editDateTime(
            Window owner,
            String columnTitle,
            String initialValue,
            double anchorScreenX,
            double anchorScreenY) {
        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.WINDOW_MODAL);
        String title =
                columnTitle != null && !columnTitle.isBlank()
                        ? columnTitle.strip()
                        : "日時の選択";
        dialog.setTitle(title);
        dialog.setHeaderText(null);

        DatePicker picker = new DatePicker();
        javafx.scene.control.TextField timeField = new javafx.scene.control.TextField();
        timeField.setPromptText("12:45");
        timeField.setPrefColumnCount(6);
        PlanInputDateColumnSupport.parseDateTimeCellValue(initialValue)
                .ifPresentOrElse(
                        dt -> {
                            picker.setValue(dt.toLocalDate());
                            timeField.setText(
                                    String.format("%d:%02d", dt.getHour(), dt.getMinute()));
                        },
                        () -> timeField.setText("12:45"));

        Button clearButton = new Button("クリア");
        clearButton.setOnAction(
                ev -> {
                    picker.setValue(null);
                    timeField.clear();
                });

        Label hint =
                new Label(
                        columnTitle != null && !columnTitle.isBlank()
                                ? "列: " + columnTitle.strip() + "（日付＋時刻 HH:mm）"
                                : "日付と時刻を選択してください");
        hint.setStyle("-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 18%);");

        HBox pickerRow = new HBox(8, picker, timeField, clearButton);
        pickerRow.setAlignment(javafx.geometry.Pos.CENTER_LEFT);

        VBox box = new VBox(10, hint, pickerRow);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.getDialogPane().setPrefWidth(360);

        dialog.setOnShown(
                e ->
                        Platform.runLater(
                                () -> {
                                    positionNearAnchor(
                                            dialog.getDialogPane().getScene().getWindow(),
                                            anchorScreenX,
                                            anchorScreenY);
                                    focusForImmediateEdit(timeField);
                                }));

        Optional<ButtonType> r = dialog.showAndWait();
        if (r.isPresent() && r.get() == ButtonType.OK) {
            java.time.LocalDate d = picker.getValue();
            if (d == null) {
                return Optional.of("");
            }
            java.time.LocalTime t = parseTimeOrDefault(timeField.getText());
            return Optional.of(
                    PlanInputDateColumnSupport.formatDateTimeCellValue(d.atTime(t)));
        }
        return Optional.empty();
    }

    private static java.time.LocalTime parseTimeOrDefault(String raw) {
        if (raw != null) {
            String s = raw.strip();
            if (!s.isEmpty()) {
                try {
                    String[] hm = s.split(":");
                    int hh = Integer.parseInt(hm[0].strip());
                    int mm = hm.length > 1 ? Integer.parseInt(hm[1].strip()) : 0;
                    if (hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59) {
                        return java.time.LocalTime.of(hh, mm);
                    }
                } catch (RuntimeException ignored) {
                    // 解釈不能は既定 12:45。
                }
            }
        }
        return java.time.LocalTime.of(12, 45);
    }

    private static void positionNearAnchor(Window win, double anchorScreenX, double anchorScreenY) {
        if (win == null) {
            return;
        }
        win.sizeToScene();
        double ww = win.getWidth();
        double hh = win.getHeight();
        Rectangle2D bounds = null;
        for (Screen s : Screen.getScreensForRectangle(anchorScreenX, anchorScreenY, 1, 1)) {
            bounds = s.getVisualBounds();
            break;
        }
        if (bounds == null) {
            bounds = Screen.getPrimary().getVisualBounds();
        }
        double x = anchorScreenX - ww * 0.15;
        double y = anchorScreenY - 48;
        x = Math.max(bounds.getMinX(), Math.min(x, bounds.getMaxX() - ww));
        y = Math.max(bounds.getMinY(), Math.min(y, bounds.getMaxY() - hh));
        win.setX(x);
        win.setY(y);
    }

    /** ダイアログ表示直後に入力欄へフォーカスし、既存値を全選択する。 */
    private static void focusForImmediateEdit(TextInputControl field) {
        if (field == null) {
            return;
        }
        field.requestFocus();
        field.selectAll();
    }
}
