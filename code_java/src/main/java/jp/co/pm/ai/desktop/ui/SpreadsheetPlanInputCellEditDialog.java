package jp.co.pm.ai.desktop.ui;

import java.util.Optional;

import javafx.application.Platform;
import javafx.geometry.Rectangle2D;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
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
                                () ->
                                        positionNearAnchor(
                                                dialog.getDialogPane().getScene().getWindow(),
                                                anchorScreenX,
                                                anchorScreenY)));

        Optional<ButtonType> r = dialog.showAndWait();
        if (r.isPresent() && r.get() == ButtonType.OK) {
            return Optional.of(area.getText());
        }
        return Optional.empty();
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
}
