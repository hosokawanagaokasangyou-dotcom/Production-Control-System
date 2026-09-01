package jp.co.pm.ai.desktop.ui;

import java.util.List;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.event.EventHandler;
import javafx.geometry.Bounds;
import javafx.geometry.Insets;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.Control;
import javafx.scene.control.ListView;
import javafx.scene.control.TextField;
import javafx.scene.input.KeyCode;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.VBox;
import javafx.stage.Popup;

import org.controlsfx.control.spreadsheet.SpreadsheetCellEditor;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * セル内は高さ固定の TextField。候補は Popup の ListView（セルサイズを変えない）。
 * マウスを動かして範囲選択しているときは編集をキャンセルする。
 */
public final class MasterDispatchListPopupEditor extends SpreadsheetCellEditor {

    static final double SELECTION_DRAG_THRESHOLD_PX = 8.0;

    private final List<String> items;
    private final SpreadsheetView spreadsheetView;
    private final TextField field = new TextField();
    private final ListView<String> list = new ListView<>();
    private final VBox popupBox = new VBox();
    private final Popup popup = new Popup();
    private boolean editing;
    private Scene hookedScene;
    private EventHandler<MouseEvent> sceneMouseHandler;
    private Double dragOriginX;
    private Double dragOriginY;

    public MasterDispatchListPopupEditor(SpreadsheetView view, List<String> items) {
        super(view);
        this.spreadsheetView = view;
        this.items = List.copyOf(items != null ? items : List.of());
        field.getStyleClass().add("pm-ai-list-editor-field");
        field.setEditable(false);
        field.setFocusTraversable(true);
        field.setMinHeight(18);
        field.setPrefHeight(18);
        field.setMaxHeight(18);
        field.setPadding(new Insets(0, 2, 0, 2));
        field.setOnKeyPressed(
                e -> {
                    if (e.getCode() == KeyCode.ESCAPE) {
                        endEdit(false);
                        e.consume();
                    } else if (e.getCode() == KeyCode.ENTER || e.getCode() == KeyCode.DOWN) {
                        showPicker();
                        e.consume();
                    }
                });
        list.setItems(FXCollections.observableArrayList(this.items));
        list.setFocusTraversable(true);
        list.setOnMouseClicked(
                e -> {
                    if (e.getClickCount() >= 1 && list.getSelectionModel().getSelectedItem() != null) {
                        commitSelected();
                    }
                });
        list.setOnKeyPressed(
                e -> {
                    if (e.getCode() == KeyCode.ENTER) {
                        commitSelected();
                        e.consume();
                    } else if (e.getCode() == KeyCode.ESCAPE) {
                        endEdit(false);
                        e.consume();
                    }
                });
        popupBox.getStyleClass().add("pm-ai-list-picker");
        popupBox.getChildren().add(list);
        var cssUrl =
                SpreadsheetTabularSupport.class.getResource(
                        "/jp/co/pm/ai/desktop/css/delivery-calendar-spreadsheet.css");
        if (cssUrl != null) {
            String css = cssUrl.toExternalForm();
            if (!popupBox.getStylesheets().contains(css)) {
                popupBox.getStylesheets().add(css);
            }
        }
        popup.getContent().add(popupBox);
        popup.setAutoHide(true);
        popup.setHideOnEscape(true);
        popup.setAutoFix(true);
        popup.setOnAutoHide(
                e -> {
                    if (editing) {
                        endEdit(false);
                    }
                });
    }

    static boolean isSelectionDrag(double originX, double originY, double x, double y) {
        return Math.hypot(x - originX, y - originY) > SELECTION_DRAG_THRESHOLD_PX;
    }

    private boolean hasMultiCellSelection() {
        if (spreadsheetView == null || spreadsheetView.getSelectionModel() == null) {
            return false;
        }
        return spreadsheetView.getSelectionModel().getSelectedCells().size() > 1;
    }

    private void commitSelected() {
        String selected = list.getSelectionModel().getSelectedItem();
        if (selected == null) {
            return;
        }
        field.setText(selected);
        endEdit(true);
    }

    private void showPicker() {
        if (!editing || popup.isShowing()) {
            return;
        }
        String current = field.getText() != null ? field.getText() : "";
        list.getSelectionModel().select(current);
        int n = Math.max(1, items.size());
        list.setPrefHeight(Math.min(280, 24.0 * n + 8));
        list.setPrefWidth(Math.max(120, field.getWidth()));
        Bounds b = field.localToScreen(field.getBoundsInLocal());
        if (b == null) {
            return;
        }
        popup.show(field, b.getMinX(), b.getMaxY());
        list.requestFocus();
    }

    private void attachSceneMouseWatch() {
        detachSceneMouseWatch();
        Scene scene = field.getScene();
        if (scene == null) {
            Platform.runLater(
                    () -> {
                        if (editing) {
                            attachSceneMouseWatch();
                        }
                    });
            return;
        }
        hookedScene = scene;
        sceneMouseHandler =
                e -> {
                    if (!editing) {
                        return;
                    }
                    if (e.getEventType() == MouseEvent.MOUSE_DRAGGED && e.isPrimaryButtonDown()) {
                        if (dragOriginX == null) {
                            dragOriginX = e.getScreenX();
                            dragOriginY = e.getScreenY();
                            return;
                        }
                        if (isSelectionDrag(dragOriginX, dragOriginY, e.getScreenX(), e.getScreenY())) {
                            detachSceneMouseWatch();
                            endEdit(false);
                        }
                    } else if (e.getEventType() == MouseEvent.MOUSE_RELEASED) {
                        detachSceneMouseWatch();
                        if (!editing) {
                            return;
                        }
                        if (hasMultiCellSelection()) {
                            endEdit(false);
                            return;
                        }
                        Platform.runLater(this::showPicker);
                    }
                };
        scene.addEventFilter(MouseEvent.MOUSE_DRAGGED, sceneMouseHandler);
        scene.addEventFilter(MouseEvent.MOUSE_RELEASED, sceneMouseHandler);
    }

    private void detachSceneMouseWatch() {
        if (hookedScene != null && sceneMouseHandler != null) {
            hookedScene.removeEventFilter(MouseEvent.MOUSE_DRAGGED, sceneMouseHandler);
            hookedScene.removeEventFilter(MouseEvent.MOUSE_RELEASED, sceneMouseHandler);
        }
        hookedScene = null;
        sceneMouseHandler = null;
        dragOriginX = null;
        dragOriginY = null;
    }

    private static boolean ancestorIsPressed(Node node) {
        Node n = node;
        while (n != null) {
            if (n.isPressed()) {
                return true;
            }
            n = n.getParent();
        }
        return false;
    }

    @Override
    public void startEdit(Object value, String format, Object... options) {
        editing = true;
        String s = value == null ? "" : value.toString();
        field.setText(s);
        field.requestFocus();
        attachSceneMouseWatch();
        Platform.runLater(
                () -> {
                    if (!editing) {
                        return;
                    }
                    if (ancestorIsPressed(field) || hasMultiCellSelection()) {
                        if (hasMultiCellSelection()) {
                            endEdit(false);
                        }
                        return;
                    }
                    showPicker();
                });
    }

    @Override
    public String getControlValue() {
        return field.getText() != null ? field.getText() : "";
    }

    @Override
    public void end() {
        editing = false;
        detachSceneMouseWatch();
        popup.hide();
    }

    @Override
    public Control getEditor() {
        return field;
    }
}
