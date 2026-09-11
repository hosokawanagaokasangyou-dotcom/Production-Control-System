package jp.co.pm.ai.desktop.reconciliation;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.Supplier;

import javafx.application.Platform;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.geometry.Orientation;
import javafx.scene.Parent;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ChoiceDialog;
import javafx.scene.control.ComboBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.SplitPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.input.KeyCode;
import javafx.scene.input.MouseButton;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup.PlanEntry;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup.PipelineScanIndex;

/**
 * 受注検索。左条件・右結果。選択した依頼NOの後加工検査表を開ける。
 */
public final class JuchuOrderSearchPane {

    private JuchuOrderSearchPane() {}

    public static Parent build(Supplier<List<OrderRecord>> recordsSupplier) {
        return build(recordsSupplier, Map::of, () -> null);
    }

    public static Parent build(
            Supplier<List<OrderRecord>> recordsSupplier,
            Supplier<Map<String, String>> uiEnv,
            Supplier<Window> ownerSupplier) {
        Objects.requireNonNull(recordsSupplier, "recordsSupplier");
        Supplier<Map<String, String>> env = uiEnv != null ? uiEnv : Map::of;
        Supplier<Window> owner = ownerSupplier != null ? ownerSupplier : () -> null;

        LocalDate today = LocalDate.now();
        DatePicker from = new DatePicker(JuchuOrderSearch.defaultDeliveryFrom(today));
        DatePicker to = new DatePicker(JuchuOrderSearch.defaultDeliveryTo(today));
        ComboBox<String> product = keywordCombo("製品名（部分一致・候補から選択可）");
        ComboBox<String> raw = keywordCombo("投入原反（部分一致・候補から選択可）");
        ComboBox<String> machine = keywordCombo("機械名（部分一致・候補から選択可）");
        ComboBox<String> process = keywordCombo("工程名（部分一致・候補から選択可）");
        Button search = new Button("検索");
        Button openKensa = new Button("検査表を開く");
        Button rebuildIndex = new Button("検査表索引を更新");
        Label statusMessage = new Label("");
        statusMessage.setWrapText(true);

        VBox conditions = new VBox(8);
        conditions.setPadding(new Insets(12));
        conditions
                .getChildren()
                .addAll(
                        labeled("納期 From", from),
                        labeled("納期 To", to),
                        labeled("製品", product),
                        labeled("投入原反", raw),
                        labeled("機械名", machine),
                        labeled("工程名", process),
                        search,
                        openKensa,
                        rebuildIndex,
                        statusMessage);

        ScrollPane leftScroll = new ScrollPane(conditions);
        leftScroll.setFitToWidth(true);
        leftScroll.getStyleClass().add("form-scroll-pane");

        Label countLabel = new Label("0 件");
        ObservableList<OrderRecord> items = FXCollections.observableArrayList();
        TableView<OrderRecord> table = new TableView<>(items);
        BooleanProperty indexBusy = new SimpleBooleanProperty(false);
        PipelineScanIndex[] planIndex = {PipelineScanIndex.empty()};
        Runnable refreshKeywordCandidates =
                () -> {
                    List<OrderRecord> recs = recordsSupplier.get();
                    if (planIndex[0].planEntriesByTaskId().isEmpty()) {
                        planIndex[0] = loadPlanIndex(env.get());
                    }
                    setComboCandidates(
                            product, JuchuOrderSearch.productCandidates(recs));
                    setComboCandidates(
                            raw, JuchuOrderSearch.rawMaterialCandidates(recs));
                    setComboCandidates(
                            machine,
                            JuchuOrderSearch.machineCandidates(
                                    recs, collectPlanNames(planIndex[0], true)));
                    setComboCandidates(
                            process,
                            JuchuOrderSearch.processCandidates(
                                    recs, collectPlanNames(planIndex[0], false)));
                };
        product.setOnShowing(e -> refreshKeywordCandidates.run());
        raw.setOnShowing(e -> refreshKeywordCandidates.run());
        machine.setOnShowing(e -> refreshKeywordCandidates.run());
        process.setOnShowing(e -> refreshKeywordCandidates.run());
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        table.setPlaceholder(new Label("条件を指定して検索してください"));
        table.getColumns()
                .addAll(
                        col("依頼No", r -> nullToEmpty(r.getReqNo())),
                        col("希望納期", r -> dbValue(r, "希望納期")),
                        col("調整納期", r -> dbValue(r, "調整納期")),
                        col("製品", r -> dbValue(r, "製品")),
                        col("原反", r -> JuchuOrderSearch.displayRawMaterial(r.getDbValues())),
                        col(
                                "機械名",
                                r ->
                                        JuchuOrderSearch.displayMachine(
                                                r.getDbValues(),
                                                planHaystack(planIndex[0], r, true))),
                        col(
                                "工程名",
                                r ->
                                        JuchuOrderSearch.displayProcess(
                                                r.getDbValues(),
                                                planHaystack(planIndex[0], r, false))),
                        col("ユーザー", r -> nullToEmpty(r.getUser())),
                        col("入力日", r -> dbValue(r, "入力日")));

        openKensa
                .disableProperty()
                .bind(
                        table.getSelectionModel()
                                .selectedItemProperty()
                                .isNull()
                                .or(indexBusy));
        rebuildIndex.disableProperty().bind(indexBusy);

        table.setRowFactory(
                tv -> {
                    TableRow<OrderRecord> row = new TableRow<>();
                    row.setOnMouseClicked(
                            e -> {
                                if (e.getClickCount() != 2
                                        || row.isEmpty()
                                        || e.getButton() != MouseButton.PRIMARY) {
                                    return;
                                }
                                openInspectionSheet(
                                        row.getItem(), env, owner, statusMessage, indexBusy);
                            });
                    return row;
                });
        table.setOnKeyPressed(
                e -> {
                    if (e.getCode() != KeyCode.ENTER) {
                        return;
                    }
                    openInspectionSheet(
                            table.getSelectionModel().getSelectedItem(),
                            env,
                            owner,
                            statusMessage,
                            indexBusy);
                    e.consume();
                });

        VBox right = new VBox(8, countLabel, table);
        right.setPadding(new Insets(12));
        VBox.setVgrow(table, Priority.ALWAYS);

        search.setOnAction(
                e -> {
                    var c =
                            new JuchuOrderSearchCriteria(
                                    from.getValue(),
                                    to.getValue(),
                                    comboText(product),
                                    comboText(raw),
                                    comboText(machine),
                                    comboText(process));
                    Optional<String> err = c.validationError();
                    if (err.isPresent()) {
                        statusMessage.setText(err.get());
                        items.clear();
                        countLabel.setText("0 件");
                        return;
                    }
                    planIndex[0] = loadPlanIndex(env.get());
                    PipelineScanIndex index = planIndex[0];
                    refreshKeywordCandidates.run();
                    List<OrderRecord> hits =
                            JuchuOrderSearch.filter(
                                    recordsSupplier.get(),
                                    c,
                                    r -> planHaystack(index, r, true),
                                    r -> planHaystack(index, r, false));
                    items.setAll(hits);
                    String countText = hits.size() + " 件";
                    statusMessage.setText(countText);
                    countLabel.setText(countText);
                });

        openKensa.setOnAction(
                e ->
                        openInspectionSheet(
                                table.getSelectionModel().getSelectedItem(),
                                env,
                                owner,
                                statusMessage,
                                indexBusy));
        rebuildIndex.setOnAction(
                e -> rebuildInspectionIndex(env, owner, statusMessage, indexBusy));

        SplitPane split = new SplitPane(leftScroll, right);
        split.setOrientation(Orientation.HORIZONTAL);
        split.setDividerPositions(0.32);
        SplitPane.setResizableWithParent(leftScroll, Boolean.TRUE);
        SplitPane.setResizableWithParent(right, Boolean.TRUE);

        VBox root = new VBox(split);
        VBox.setVgrow(split, Priority.ALWAYS);
        root.setMaxWidth(Double.MAX_VALUE);
        root.setMaxHeight(Double.MAX_VALUE);
        return root;
    }

    private static void rebuildInspectionIndex(
            Supplier<Map<String, String>> uiEnv,
            Supplier<Window> owner,
            Label status,
            BooleanProperty busy) {
        Map<String, String> ui = uiEnv.get();
        Path dir = InspectionSheetOpenService.resolveDir(ui);
        if (!InspectionSheetOpenService.dirReachable(ui)) {
            status.setText("検査表フォルダにアクセスできません: " + dir);
            return;
        }
        busy.set(true);
        status.setText("検査表索引を更新中…");
        InspectionSheetOpenService.startBackgroundRebuild(
                        ui,
                        (done, total) ->
                                Platform.runLater(
                                        () -> status.setText("検査表索引を更新中… " + done + " / " + total)))
                .whenComplete(
                        (r, ex) ->
                                Platform.runLater(
                                        () -> {
                                            busy.set(false);
                                            if (ex != null) {
                                                status.setText(
                                                        "検査表索引の更新に失敗: " + ex.getMessage());
                                                return;
                                            }
                                            status.setText(
                                                    "検査表索引 "
                                                            + r.rows().size()
                                                            + " 件（Excel読込 "
                                                            + r.readExcelCount()
                                                            + "）");
                                        }));
    }

    private static void openInspectionSheet(
            OrderRecord record,
            Supplier<Map<String, String>> uiEnv,
            Supplier<Window> owner,
            Label status,
            BooleanProperty busy) {
        if (record == null) {
            return;
        }
        String irai = record.getReqNo();
        if (irai == null || irai.isBlank()) {
            status.setText("依頼NOが空です");
            return;
        }
        Map<String, String> ui = uiEnv.get();
        Path dir = InspectionSheetOpenService.resolveDir(ui);
        if (!InspectionSheetOpenService.dirReachable(ui)) {
            status.setText(
                    "検査表フォルダ未設定または未到達です。環境変数 "
                            + AppPaths.KEY_PM_AI_INSPECTION_SHEET_DIR
                            + " を設定してください: "
                            + dir);
            return;
        }
        busy.set(true);
        status.setText("検査表を検索中…");
        Task<List<InspectionSheetIndexStore.Row>> task =
                new Task<>() {
                    @Override
                    protected List<InspectionSheetIndexStore.Row> call() throws Exception {
                        return InspectionSheetOpenService.find(ui, irai);
                    }
                };
        task.setOnSucceeded(
                e -> {
                    busy.set(false);
                    List<InspectionSheetIndexStore.Row> hits = task.getValue();
                    if (hits == null || hits.isEmpty()) {
                        status.setText("検査表が見つかりません: " + irai);
                        return;
                    }
                    InspectionSheetIndexStore.Row chosen = chooseHit(hits, owner.get());
                    if (chosen == null) {
                        status.setText("検査表を開く操作をキャンセルしました");
                        return;
                    }
                    try {
                        InspectionSheetOpenService.open(chosen);
                        status.setText("検査表を開きました: " + chosen.fileName());
                    } catch (Exception ex) {
                        status.setText("検査表を開けませんでした: " + ex.getMessage());
                        alert(owner.get(), "検査表を開けませんでした: " + ex.getMessage());
                    }
                });
        task.setOnFailed(
                e -> {
                    busy.set(false);
                    Throwable ex = task.getException();
                    status.setText("検査表検索に失敗: " + (ex != null ? ex.getMessage() : ""));
                });
        Thread t = new Thread(task, "inspection-sheet-open");
        t.setDaemon(true);
        t.start();
    }

    private static InspectionSheetIndexStore.Row chooseHit(
            List<InspectionSheetIndexStore.Row> hits, Window owner) {
        if (hits.size() == 1) {
            return hits.get(0);
        }
        List<String> labels = new java.util.ArrayList<>();
        for (InspectionSheetIndexStore.Row row : hits) {
            String date = row.processingDate() != null ? row.processingDate().toString() : "";
            labels.add(row.fileName() + (date.isEmpty() ? "" : "（" + date + "）"));
        }
        ChoiceDialog<String> dialog = new ChoiceDialog<>(labels.get(0), labels);
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle("検査表が複数あります");
        dialog.setHeaderText(null);
        dialog.setContentText("開く検査表を選んでください");
        Optional<String> selected = dialog.showAndWait();
        if (selected.isEmpty()) {
            return null;
        }
        int idx = labels.indexOf(selected.get());
        if (idx < 0) {
            return hits.get(0);
        }
        return hits.get(idx);
    }

    private static void alert(Window owner, String message) {
        Alert a = new Alert(AlertType.ERROR);
        if (owner != null) {
            a.initOwner(owner);
        }
        a.setTitle("検査表");
        a.setHeaderText(null);
        a.setContentText(message);
        a.showAndWait();
    }

    private static PipelineScanIndex loadPlanIndex(Map<String, String> ui) {
        try {
            Path path = AppPaths.resolveShapedAladdinPlanJsonPath(ui);
            if (path == null || !Files.isRegularFile(path)) {
                return PipelineScanIndex.empty();
            }
            AladdinShapedPlanQtyLookup.ShapedTable table =
                    AladdinShapedPlanQtyLookup.loadShapedTable(path);
            return AladdinShapedPlanQtyLookup.buildPipelineScanIndex(table.headers(), table.rows());
        } catch (RuntimeException ex) {
            return PipelineScanIndex.empty();
        }
    }

    private static String planHaystack(
            PipelineScanIndex index, OrderRecord record, boolean machine) {
        if (index == null || record == null) {
            return "";
        }
        List<PlanEntry> entries = index.planEntriesFor(record.getReqNo());
        if (entries == null || entries.isEmpty()) {
            return "";
        }
        Set<String> names = new LinkedHashSet<>();
        for (PlanEntry entry : entries) {
            String value = machine ? entry.machineName() : entry.processName();
            if (value != null && !value.isBlank()) {
                names.add(value.strip());
            }
        }
        return String.join(" ", names);
    }

    private static ComboBox<String> keywordCombo(String prompt) {
        ComboBox<String> combo = new ComboBox<>();
        combo.setEditable(true);
        combo.setMaxWidth(Double.MAX_VALUE);
        combo.setVisibleRowCount(12);
        combo.setPromptText(prompt);
        HBox.setHgrow(combo, Priority.ALWAYS);
        return combo;
    }

    private static String comboText(ComboBox<String> combo) {
        if (combo == null) {
            return "";
        }
        if (combo.isEditable() && combo.getEditor() != null) {
            String typed = combo.getEditor().getText();
            if (typed != null) {
                return typed;
            }
        }
        return combo.getValue() != null ? combo.getValue() : "";
    }

    private static void setComboCandidates(ComboBox<String> combo, List<String> names) {
        if (combo == null) {
            return;
        }
        String typed = comboText(combo);
        combo.getItems().setAll(names != null ? names : List.of());
        if (combo.getEditor() != null) {
            combo.getEditor().setText(typed);
        }
    }

    private static List<String> collectPlanNames(PipelineScanIndex index, boolean machine) {
        if (index == null) {
            return List.of();
        }
        Set<String> names = new LinkedHashSet<>();
        for (List<PlanEntry> entries : index.planEntriesByTaskId().values()) {
            if (entries == null) {
                continue;
            }
            for (PlanEntry entry : entries) {
                String value = machine ? entry.machineName() : entry.processName();
                if (value != null && !value.isBlank()) {
                    names.add(value.strip());
                }
            }
        }
        return List.copyOf(names);
    }

    private static VBox labeled(String caption, javafx.scene.Node field) {
        Label label = new Label(caption);
        VBox box = new VBox(4, label, field);
        HBox.setHgrow(field, Priority.ALWAYS);
        if (field instanceof TextField tf) {
            tf.setMaxWidth(Double.MAX_VALUE);
        } else if (field instanceof ComboBox<?> cb) {
            cb.setMaxWidth(Double.MAX_VALUE);
        } else if (field instanceof DatePicker dp) {
            dp.setMaxWidth(Double.MAX_VALUE);
        }
        return box;
    }

    private static TableColumn<OrderRecord, String> col(
            String title, java.util.function.Function<OrderRecord, String> value) {
        TableColumn<OrderRecord, String> column = new TableColumn<>(title);
        column.setCellValueFactory(
                cd -> {
                    OrderRecord row = cd.getValue();
                    return new SimpleStringProperty(row != null ? value.apply(row) : "");
                });
        return column;
    }

    private static String dbValue(OrderRecord record, String key) {
        Map<String, String> db = record.getDbValues();
        if (db == null) {
            return "";
        }
        return nullToEmpty(db.get(key));
    }

    private static String nullToEmpty(String value) {
        return value != null ? value : "";
    }
}
