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
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.ContentDisplay;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.dispatch.DispatchInteractiveRollUnitSupport;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchNormalizer;
import jp.co.pm.ai.desktop.dispatch.Stage2InProgressNextDayRollInput;
import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;

/** 段階2直前の翌日ロール入力ダイアログ（①加工途中 / ②アラジン除外）共通 UI。 */
final class Stage2NextDayRollDispatchDialogSupport {

    static final ButtonType COPY_HTML =
            new ButtonType("HTMLコピー", ButtonBar.ButtonData.LEFT);

    private Stage2NextDayRollDispatchDialogSupport() {}

    interface RowModel {
        String taskId();

        String process();

        String machineName();

        default String targetReason() {
            return "";
        }

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

        /**
         * 上限本数 {@link #maxRolls()} に対応する m 換算の上限（既定は残量そのもの）。
         *
         * <p>アラジン除外ダイアログのように「残量」と「上限」が異なる実装は、この値と
         * {@link #maxRolls()} を整合させて上書きする（{@code rowDetail} の表示とエラー文言の不一致を防ぐ）。
         */
        default double effectiveCapM() {
            return remainingM();
        }

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

        TableColumn<RowModel, String> cProcess = createProcessColumn();

        TableColumn<RowModel, String> cReason = null;
        if (rows.stream().anyMatch(r -> !r.targetReason().isBlank())) {
            TableColumn<RowModel, String> cTargetReason = new TableColumn<>("対象理由");
            cTargetReason.setCellValueFactory(
                    cd ->
                            new javafx.beans.property.SimpleStringProperty(
                                    cd.getValue().targetReason()));
            cTargetReason.setEditable(false);
            cTargetReason.setPrefWidth(92);
            cReason = cTargetReason;
        }

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

        TableColumn<RowModel, String> cRolls = createRollCountColumn(theme.rollsColumnLabel());

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
        cols.add(cProcess);
        if (cReason != null) {
            cols.add(cReason);
        }
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
        prefW += 90;
        if (cReason != null) {
            prefW += 92;
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
        dialog.getDialogPane().getButtonTypes().setAll(COPY_HTML, ButtonType.OK, ButtonType.CANCEL);
        Node copyHtmlButton = dialog.getDialogPane().lookupButton(COPY_HTML);
        if (copyHtmlButton instanceof javafx.scene.control.Button copyBtn) {
            copyBtn.setDefaultButton(false);
            copyBtn.setCancelButton(false);
            copyBtn.addEventFilter(
                    javafx.event.ActionEvent.ACTION,
                    ev -> {
                        ev.consume();
                        commitPendingTableCellEdit(table);
                        ClipboardTableSupport.copyHtmlTableOnly(toClipboardHtml(theme, rows));
                    });
        }
        Node okButton = dialog.getDialogPane().lookupButton(ButtonType.OK);
        if (okButton instanceof javafx.scene.control.Button okBtn) {
            okBtn.setDefaultButton(true);
        }
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

    static String toClipboardHtml(Theme theme, List<? extends RowModel> rows) {
        Theme t = theme != null ? theme : emptyTheme();
        List<? extends RowModel> safe = rows != null ? rows : List.of();
        boolean includeReason =
                safe.stream().anyMatch(r -> r != null && !r.targetReason().isBlank());
        List<HtmlColumn> columns = clipboardColumns(t, includeReason);
        StringBuilder sb = new StringBuilder();
        appendEscapedBlock(sb, "h2", t.title());
        appendEscapedBlock(sb, "p", t.headerText());
        appendEscapedBlock(sb, "p", t.hintText());
        sb.append(
                "<table border=\"1\" cellspacing=\"0\" cellpadding=\"4\""
                        + " style=\"border-collapse:collapse;font-family:'Meiryo UI',sans-serif;font-size:11pt;\">");
        sb.append("<thead><tr>");
        for (HtmlColumn col : columns) {
            sb.append("<th style=\"background:#D9E1F2;padding:4px 8px;text-align:left;\">")
                    .append(ClipboardTableSupport.escapeHtml(col.header()))
                    .append("</th>");
        }
        sb.append("</tr></thead><tbody>");
        for (RowModel row : safe) {
            if (row == null) {
                continue;
            }
            sb.append("<tr>");
            for (HtmlColumn col : columns) {
                String align = col.alignCss();
                sb.append("<td style=\"padding:4px 8px;")
                        .append(align)
                        .append("\">")
                        .append(ClipboardTableSupport.escapeHtml(col.cell(row)))
                        .append("</td>");
            }
            sb.append("</tr>");
        }
        sb.append("</tbody></table>");
        return sb.toString();
    }

    private static Theme emptyTheme() {
        return new Theme("", "", "", "実加工", "翌日配台(ロール)", "", "");
    }

    private static void appendEscapedBlock(StringBuilder sb, String tag, String text) {
        if (text == null || text.isBlank()) {
            return;
        }
        sb.append('<')
                .append(tag)
                .append('>')
                .append(ClipboardTableSupport.escapeHtml(text))
                .append("</")
                .append(tag)
                .append('>');
    }

    private record HtmlColumn(
            String header, String alignCss, java.util.function.Function<RowModel, String> cellFn) {
        String cell(RowModel row) {
            return cellFn.apply(row);
        }
    }

    private static List<HtmlColumn> clipboardColumns(Theme theme, boolean includeTargetReason) {
        Theme t = theme != null ? theme : emptyTheme();
        List<HtmlColumn> cols = new java.util.ArrayList<>();
        String left = "text-align:left;";
        String right = "text-align:right;";
        String center = "text-align:center;";
        cols.add(new HtmlColumn("依頼NO", left, RowModel::taskId));
        cols.add(new HtmlColumn("機械名", left, RowModel::machineName));
        cols.add(new HtmlColumn("工程名", left, RowModel::process));
        if (includeTargetReason) {
            cols.add(new HtmlColumn("対象理由", left, RowModel::targetReason));
        }
        cols.add(
                new HtmlColumn(
                        t.referenceColumnLabel(),
                        right,
                        r -> formatM(r.referenceM()) + " m"));
        if (t.showPlanQtyColumns()) {
            cols.add(
                    new HtmlColumn(
                            "換算数量", right, r -> formatM(r.convertedQtyM()) + " m"));
            cols.add(
                    new HtmlColumn(
                            "配台数量", right, r -> formatM(r.dispatchQtyM()) + " m"));
        }
        if (t.showAladdinTodayPlanColumn()) {
            cols.add(
                    new HtmlColumn(
                            "アラジン当日",
                            right,
                            r -> {
                                double m = r.aladdinTodayPlanM();
                                if (m <= 1e-12) {
                                    return "—";
                                }
                                return formatM(m) + " m";
                            }));
        }
        cols.add(new HtmlColumn("残量", right, r -> formatM(r.remainingM()) + " m"));
        cols.add(
                new HtmlColumn(
                        "1ロール",
                        right,
                        r -> {
                            if (r.unitM() <= 1e-9) {
                                return "—";
                            }
                            return ResultDispatchNormalizer.formatQty(r.unitM()) + " m";
                        }));
        cols.add(
                new HtmlColumn(
                        t.rollsColumnLabel(),
                        center,
                        r ->
                                clampRollCountChoice(
                                        r.rollCountProperty() != null
                                                ? r.rollCountProperty().get()
                                                : "0",
                                        r.maxRolls())));
        cols.add(
                new HtmlColumn(
                        "換算(m)",
                        right,
                        r -> {
                            String rollsRaw =
                                    r.rollCountProperty() != null
                                            ? r.rollCountProperty().get()
                                            : "0";
                            int n =
                                    Stage2InProgressNextDayRollInput.parseNonNegativeRollCount(
                                                    clampRollCountChoice(rollsRaw, r.maxRolls()))
                                            .orElse(0);
                            return Stage2InProgressNextDayRollInput.formatConvertedMetersPreview(
                                    n, r.unitM());
                        }));
        return List.copyOf(cols);
    }

    static TableColumn<RowModel, String> createProcessColumn() {
        TableColumn<RowModel, String> column = new TableColumn<>("工程名");
        column.setCellValueFactory(
                cd -> new javafx.beans.property.SimpleStringProperty(cd.getValue().process()));
        column.setEditable(false);
        column.setPrefWidth(90);
        return column;
    }

    static TableColumn<RowModel, String> createRollCountColumn(String label) {
        TableColumn<RowModel, String> column = new TableColumn<>(label);
        column.getStyleClass().add("pm-next-day-roll-input-column");
        column.setCellValueFactory(cd -> cd.getValue().rollCountProperty());
        column.setCellFactory(col -> new RollCountComboTableCell());
        column.setEditable(false);
        column.setStyle("-fx-alignment: CENTER;");
        column.setPrefWidth(88);
        return column;
    }

    static List<String> rollCountChoices(int maxRolls) {
        int n = Math.max(0, maxRolls);
        java.util.ArrayList<String> out = new java.util.ArrayList<>(n + 1);
        for (int i = 0; i <= n; i++) {
            out.add(String.valueOf(i));
        }
        return List.copyOf(out);
    }

    static String clampRollCountChoice(String raw, int maxRolls) {
        int max = Math.max(0, maxRolls);
        int n = Stage2InProgressNextDayRollInput.parseNonNegativeRollCount(raw).orElse(0);
        if (n > max) {
            n = max;
        }
        return String.valueOf(n);
    }

    private static final class RollCountComboTableCell extends TableCell<RowModel, String> {
        private final ComboBox<String> combo = new ComboBox<>();
        private boolean syncing;

        RollCountComboTableCell() {
            getStyleClass().add("pm-next-day-roll-input-cell");
            combo.getStyleClass().add("pm-next-day-roll-input-combo");
            combo.setEditable(false);
            combo.setMaxWidth(Double.MAX_VALUE);
            combo.valueProperty()
                    .addListener(
                            (o, a, b) -> {
                                if (syncing || b == null) {
                                    return;
                                }
                                RowModel row = currentRow();
                                if (row != null) {
                                    row.rollCountProperty().set(b);
                                }
                            });
            setGraphic(combo);
            setContentDisplay(ContentDisplay.GRAPHIC_ONLY);
        }

        private RowModel currentRow() {
            if (getTableRow() != null && getTableRow().getItem() != null) {
                return getTableRow().getItem();
            }
            if (getTableView() != null
                    && getIndex() >= 0
                    && getIndex() < getTableView().getItems().size()) {
                return getTableView().getItems().get(getIndex());
            }
            return null;
        }

        @Override
        protected void updateItem(String item, boolean empty) {
            super.updateItem(item, empty);
            RowModel row = currentRow();
            if (empty || row == null) {
                setGraphic(null);
                return;
            }
            List<String> choices = rollCountChoices(row.maxRolls());
            String clamped = clampRollCountChoice(item, row.maxRolls());
            syncing = true;
            try {
                combo.setItems(FXCollections.observableArrayList(choices));
                combo.setVisibleRowCount(Math.min(12, choices.size()));
                combo.setValue(clamped);
            } finally {
                syncing = false;
            }
            if (!clamped.equals(row.rollCountProperty().get())) {
                row.rollCountProperty().set(clamped);
            }
            setGraphic(combo);
        }
    }

    private static String rowDetail(RowModel row) {
        String unitLine =
                row.unitM() > 1e-9
                        ? DispatchInteractiveRollUnitSupport.rollUnitDialogHeader(
                                row.effectiveCapM(), row.unitInfo(), row.taskId() + " / " + row.machineName())
                        : "依頼NO "
                                + row.taskId()
                                + " / "
                                + row.machineName()
                                + "\n配台ロール単位 (m) を決定できません。";
        return unitLine;
    }

    private static void commitPendingTableCellEdit(TableView<RowModel> table) {
        for (Node node : table.lookupAll(".combo-box.pm-next-day-roll-input-combo")) {
            if (!(node instanceof ComboBox<?> combo) || combo.getValue() == null) {
                continue;
            }
            TableRow<?> tr = findTableRow(combo);
            if (tr != null && tr.getItem() instanceof RowModel row) {
                row.rollCountProperty().set(String.valueOf(combo.getValue()));
            }
        }
        table.edit(-1, null);
    }

    private static TableRow<?> findTableRow(Node node) {
        Parent p = node.getParent();
        while (p != null) {
            if (p instanceof TableRow<?> tr) {
                return tr;
            }
            p = p.getParent();
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
