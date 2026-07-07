package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Optional;
import java.util.function.Function;

import javafx.application.Platform;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TablePosition;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;
import javafx.util.Callback;

import jp.co.pm.ai.desktop.dispatch.DispatchInteractiveRollUnitSupport;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.dispatch.Stage2InProgressNextDayRollInput;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;

/** 段階2直前の翌日ロール入力ダイアログ（①加工途中 / ②アラジン除外）共通 UI。 */
final class Stage2NextDayRollDispatchDialogSupport {

    private Stage2NextDayRollDispatchDialogSupport() {}

    interface RowModel {
        String taskId();

        String process();

        String machineName();

        double referenceM();

        /** ①ダイアログ用: 当日アラジン計画 (m)。無いときは 0。 */
        default double aladdinTodayPlanM() {
            return 0.0;
        }

        /** ①ダイアログ用: 換算数量 (m)。 */
        default double convertedQtyM() {
            return 0.0;
        }

        /** ①ダイアログ用: 配台数量（総量 m）。 */
        default double dispatchQtyM() {
            return 0.0;
        }

        double remainingM();

        Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM unitInfo();

        double unitM();

        int maxRolls();

        StringProperty rollCountProperty();
    }

    record Theme(
            String title,
            String headerText,
            String hintText,
            String referenceColumnLabel,
            String rollsColumnLabel,
            String dialogPaneStyle,
            String hintStyle,
            boolean showAladdinTodayPlanColumn,
            boolean showPlanQtyColumns) {

        Theme(
                String title,
                String headerText,
                String hintText,
                String referenceColumnLabel,
                String rollsColumnLabel,
                String dialogPaneStyle,
                String hintStyle) {
            this(
                    title,
                    headerText,
                    hintText,
                    referenceColumnLabel,
                    rollsColumnLabel,
                    dialogPaneStyle,
                    hintStyle,
                    false,
                    false);
        }

        Theme(
                String title,
                String headerText,
                String hintText,
                String referenceColumnLabel,
                String rollsColumnLabel,
                String dialogPaneStyle,
                String hintStyle,
                boolean showAladdinTodayPlanColumn) {
            this(
                    title,
                    headerText,
                    hintText,
                    referenceColumnLabel,
                    rollsColumnLabel,
                    dialogPaneStyle,
                    hintStyle,
                    showAladdinTodayPlanColumn,
                    false);
        }
    }

    static <T> Optional<List<T>> prompt(
            Window owner,
            List<? extends RowModel> rows,
            Theme theme,
            Function<RowModel, T> toEntry,
            Function<RowModel, Optional<String>> validateRow) {
        if (rows == null || rows.isEmpty()) {
            return Optional.of(List.of());
        }

        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle(theme.title());
        dialog.setHeaderText(theme.headerText());

        Label hint = new Label(theme.hintText());
        hint.setWrapText(true);
        hint.setStyle(theme.hintStyle());

        @SuppressWarnings("unchecked")
        TableView<RowModel> table =
                new TableView<>(FXCollections.observableArrayList((List<RowModel>) rows));
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        table.setEditable(true);
        table.setPrefHeight(Math.min(420, 56 + rows.size() * 28.0));

        TableColumn<RowModel, String> cTask = new TableColumn<>("依頼NO");
        cTask.setCellValueFactory(cd -> new javafx.beans.property.SimpleStringProperty(cd.getValue().taskId()));
        cTask.setEditable(false);
        cTask.setPrefWidth(80);

        TableColumn<RowModel, String> cMach = new TableColumn<>("機械名");
        cMach.setCellValueFactory(
                cd -> new javafx.beans.property.SimpleStringProperty(cd.getValue().machineName()));
        cMach.setEditable(false);
        cMach.setPrefWidth(120);

        TableColumn<RowModel, String> cRef = new TableColumn<>(theme.referenceColumnLabel());
        cRef.setCellValueFactory(
                cd ->
                        new javafx.beans.property.SimpleStringProperty(
                                formatM(cd.getValue().referenceM()) + " m"));
        cRef.setEditable(false);
        cRef.setStyle("-fx-alignment: CENTER-RIGHT;");
        cRef.setPrefWidth(68);

        TableColumn<RowModel, String> cConvQty = null;
        TableColumn<RowModel, String> cDispatchQty = null;
        if (theme.showPlanQtyColumns()) {
            TableColumn<RowModel, String> cConv = new TableColumn<>("換算数量");
            cConv.setCellValueFactory(
                    cd ->
                            new javafx.beans.property.SimpleStringProperty(
                                    formatM(cd.getValue().convertedQtyM()) + " m"));
            cConv.setEditable(false);
            cConv.setStyle("-fx-alignment: CENTER-RIGHT;");
            cConv.setPrefWidth(72);
            cConvQty = cConv;

            TableColumn<RowModel, String> cDispatch = new TableColumn<>("配台数量");
            cDispatch.setCellValueFactory(
                    cd ->
                            new javafx.beans.property.SimpleStringProperty(
                                    formatM(cd.getValue().dispatchQtyM()) + " m"));
            cDispatch.setEditable(false);
            cDispatch.setStyle("-fx-alignment: CENTER-RIGHT;");
            cDispatch.setPrefWidth(72);
            cDispatchQty = cDispatch;
        }

        TableColumn<RowModel, String> cAladdinToday = null;
        if (theme.showAladdinTodayPlanColumn()) {
            TableColumn<RowModel, String> cAladdin = new TableColumn<>("アラジン当日");
            cAladdin.setCellValueFactory(
                    cd -> {
                        double m = cd.getValue().aladdinTodayPlanM();
                        if (m <= 1e-12) {
                            return new javafx.beans.property.SimpleStringProperty("—");
                        }
                        return new javafx.beans.property.SimpleStringProperty(formatM(m) + " m");
                    });
            cAladdin.setEditable(false);
            cAladdin.setStyle("-fx-alignment: CENTER-RIGHT;");
            cAladdin.setPrefWidth(80);
            cAladdinToday = cAladdin;
        }

        TableColumn<RowModel, String> cRem = new TableColumn<>("残量");
        cRem.setCellValueFactory(
                cd ->
                        new javafx.beans.property.SimpleStringProperty(
                                formatM(cd.getValue().remainingM()) + " m"));
        cRem.setEditable(false);
        cRem.setStyle("-fx-alignment: CENTER-RIGHT;");
        cRem.setPrefWidth(68);

        TableColumn<RowModel, String> cUnit = new TableColumn<>("1ロール");
        cUnit.setCellValueFactory(
                cd -> {
                    RowModel r = cd.getValue();
                    if (r.unitM() <= 1e-9) {
                        return new javafx.beans.property.SimpleStringProperty("—");
                    }
                    return new javafx.beans.property.SimpleStringProperty(
                            ResultDispatchNormalizer.formatQty(r.unitM()) + " m");
                });
        cUnit.setEditable(false);
        cUnit.setStyle("-fx-alignment: CENTER-RIGHT;");
        cUnit.setPrefWidth(64);

        TableColumn<RowModel, String> cRolls = new TableColumn<>(theme.rollsColumnLabel());
        cRolls.getStyleClass().add("pm-next-day-roll-input-column");
        cRolls.setCellValueFactory(cd -> cd.getValue().rollCountProperty());
        Callback<TableColumn<RowModel, String>, TableCell<RowModel, String>> rollCellFactory =
                TextFieldTableCell.forTableColumn();
        cRolls.setCellFactory(
                col -> {
                    TableCell<RowModel, String> cell = rollCellFactory.call(col);
                    cell.getStyleClass().add("pm-next-day-roll-input-cell");
                    return cell;
                });
        cRolls.setOnEditCommit(
                ev -> {
                    if (ev.getNewValue() != null) {
                        ev.getRowValue().rollCountProperty().set(ev.getNewValue());
                    }
                });
        cRolls.setEditable(true);
        cRolls.setStyle("-fx-alignment: CENTER-RIGHT;");
        cRolls.setPrefWidth(72);

        TableColumn<RowModel, String> cMeters = new TableColumn<>("換算(m)");
        cMeters.setCellValueFactory(
                cd -> {
                    RowModel r = cd.getValue();
                    Optional<Integer> rolls =
                            Stage2InProgressNextDayRollInput.parseNonNegativeRollCount(
                                    r.rollCountProperty().get());
                    int n = rolls.orElse(0);
                    return new javafx.beans.property.SimpleStringProperty(
                            Stage2InProgressNextDayRollInput.formatConvertedMetersPreview(
                                    n, r.unitM()));
                });
        cMeters.setEditable(false);
        cMeters.setStyle("-fx-alignment: CENTER-RIGHT;");
        cMeters.setPrefWidth(72);

        java.util.List<TableColumn<RowModel, ?>> cols = new java.util.ArrayList<>();
        cols.add(cTask);
        cols.add(cMach);
        cols.add(cRef);
        if (cConvQty != null) {
            cols.add(cConvQty);
        }
        if (cDispatchQty != null) {
            cols.add(cDispatchQty);
        }
        if (cAladdinToday != null) {
            cols.add(cAladdinToday);
        }
        cols.add(cRem);
        cols.add(cUnit);
        cols.add(cRolls);
        cols.add(cMeters);
        table.getColumns().setAll(cols);
        table.getStyleClass().add("pm-next-day-roll-dialog-table");
        int prefW = 640;
        if (theme.showAladdinTodayPlanColumn()) {
            prefW = 720;
        }
        if (theme.showPlanQtyColumns()) {
            prefW += 144;
        }
        dialog.getDialogPane().setPrefWidth(prefW);
        table.getItems().forEach(r -> r.rollCountProperty().addListener((o, a, b) -> table.refresh()));

        StackPane tableHost = new StackPane(table);
        Region rollColumnOverlay = installRollColumnOverlay(table, cRolls, tableHost);

        VBox content = new VBox(10, hint, tableHost);
        VBox.setVgrow(tableHost, Priority.ALWAYS);
        content.setPadding(new Insets(4, 0, 0, 0));
        dialog.getDialogPane().setContent(content);
        if (theme.dialogPaneStyle() != null && !theme.dialogPaneStyle().isBlank()) {
            dialog.getDialogPane().setStyle(theme.dialogPaneStyle());
        }
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.setOnShown(
                ev -> Platform.runLater(() -> repositionRollColumnOverlay(table, cRolls, rollColumnOverlay)));
        dialog.getDialogPane()
                .lookupButton(ButtonType.OK)
                .addEventFilter(
                        javafx.event.ActionEvent.ACTION,
                        ev -> {
                            commitPendingTableCellEdit(table);
                            for (RowModel r : rows) {
                                Optional<String> err = validateRow.apply(r);
                                if (err.isPresent()) {
                                    ev.consume();
                                    showValidationError(dialog, "入力エラー", err.get(), rowDetail(r));
                                    return;
                                }
                            }
                        });

        Optional<ButtonType> result = dialog.showAndWait();
        if (result.isEmpty() || result.get() != ButtonType.OK) {
            return Optional.empty();
        }
        List<T> out = new java.util.ArrayList<>(rows.size());
        for (RowModel r : rows) {
            out.add(toEntry.apply(r));
        }
        return Optional.of(out);
    }

    private static String rowDetail(RowModel row) {
        String unitLine =
                row.unitM() > 1e-9
                        ? DispatchInteractiveRollUnitSupport.rollUnitDialogHeader(
                                row.remainingM(), row.unitInfo(), row.taskId() + " / " + row.machineName())
                        : "依頼NO "
                                + row.taskId()
                                + " / "
                                + row.machineName()
                                + "\n配台ロール単位 (m) を決定できません。";
        return unitLine;
    }

    private static void commitPendingTableCellEdit(TableView<RowModel> table) {
        TablePosition<RowModel, ?> editing = table.getEditingCell();
        if (editing == null) {
            return;
        }
        int rowIdx = editing.getRow();
        if (rowIdx < 0 || rowIdx >= table.getItems().size()) {
            table.edit(-1, null);
            return;
        }
        RowModel row = table.getItems().get(rowIdx);
        String committed = null;
        if (table.getScene() != null) {
            Node focus = table.getScene().getFocusOwner();
            if (focus instanceof TextField tf) {
                committed = tf.getText();
            }
        }
        if (committed == null) {
            for (Node node : table.lookupAll(".table-row-cell")) {
                if (node instanceof TableRow<?> tr && tr.getIndex() == rowIdx) {
                    TextField tf = findTextFieldIn(tr);
                    if (tf != null) {
                        committed = tf.getText();
                        break;
                    }
                }
            }
        }
        if (committed != null) {
            row.rollCountProperty().set(committed);
        }
        table.edit(-1, null);
    }

    private static TextField findTextFieldIn(Parent parent) {
        for (Node child : parent.getChildrenUnmodifiable()) {
            if (child instanceof TextField tf) {
                return tf;
            }
            if (child instanceof Parent p) {
                TextField nested = findTextFieldIn(p);
                if (nested != null) {
                    return nested;
                }
            }
        }
        return null;
    }

    private static void showValidationError(
            Dialog<?> dialog, String title, String header, String content) {
        Dialog<Void> err = new Dialog<>();
        err.initOwner(dialog.getDialogPane().getScene().getWindow());
        err.initModality(Modality.WINDOW_MODAL);
        err.setTitle(title);
        err.setHeaderText(header);
        err.setContentText(content);
        err.getDialogPane().getButtonTypes().setAll(ButtonType.OK);
        err.showAndWait();
    }

    /** 「翌日(ロール)」列に確認用の枠オーバーレイを載せ、列幅・テーブルサイズに追従させる。 */
    private static Region installRollColumnOverlay(
            TableView<RowModel> table,
            TableColumn<RowModel, String> rollsColumn,
            StackPane tableHost) {
        Region overlay = new Region();
        overlay.setMouseTransparent(true);
        overlay.setManaged(false);
        overlay.getStyleClass().add("pm-next-day-roll-column-overlay-frame");
        tableHost.getChildren().add(overlay);

        Runnable reposition = () -> repositionRollColumnOverlay(table, rollsColumn, overlay);
        rollsColumn.widthProperty().addListener((o, a, b) -> reposition.run());
        for (TableColumn<RowModel, ?> c : table.getColumns()) {
            c.widthProperty().addListener((o, a, b) -> reposition.run());
        }
        table.widthProperty().addListener((o, a, b) -> reposition.run());
        table.heightProperty().addListener((o, a, b) -> reposition.run());
        table.layoutXProperty().addListener((o, a, b) -> reposition.run());
        table.layoutYProperty().addListener((o, a, b) -> reposition.run());
        table.skinProperty().addListener((o, a, b) -> Platform.runLater(reposition));
        return overlay;
    }

    private static void repositionRollColumnOverlay(
            TableView<?> table, TableColumn<?, ?> rollsColumn, Region overlay) {
        if (table.getWidth() <= 0 || table.getSkin() == null) {
            overlay.setVisible(false);
            return;
        }
        double x = 0;
        double columnW = -1;
        for (TableColumn<?, ?> c : table.getVisibleLeafColumns()) {
            if (c == rollsColumn) {
                columnW = c.getWidth();
                break;
            }
            x += c.getWidth();
        }
        if (columnW <= 0) {
            overlay.setVisible(false);
            return;
        }
        overlay.setVisible(true);
        overlay.setManaged(false);
        overlay.resizeRelocate(
                table.getLayoutX() + x, table.getLayoutY(), columnW, table.getHeight());
    }

    private static String formatM(double v) {
        if (Math.abs(v - Math.rint(v)) <= 1e-9) {
            return String.valueOf((long) Math.rint(v));
        }
        return String.format(java.util.Locale.ROOT, "%.3f", v)
                .replaceAll("0+$", "")
                .replaceAll("\\.$", "");
    }
}
