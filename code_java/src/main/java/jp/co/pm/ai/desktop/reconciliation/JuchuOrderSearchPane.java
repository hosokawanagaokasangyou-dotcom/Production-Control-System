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
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Supplier;

import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.geometry.Orientation;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ChoiceDialog;
import javafx.scene.control.ComboBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Separator;
import javafx.scene.control.SplitPane;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.input.KeyCode;
import javafx.scene.input.MouseButton;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.scene.text.TextAlignment;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup.PlanEntry;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup.PipelineScanIndex;

/**
 * 受注検索。左条件・右結果。選択した依頼NOの後加工検査表を開ける。
 */
public final class JuchuOrderSearchPane {

    private static final String PROP_ALL_CANDIDATES = "juchuOrderSearch.allCandidates";
    private static final String PROP_UPDATING_FILTER = "juchuOrderSearch.updatingFilter";

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
        ComboBox<String> irai = keywordCombo("部分一致");
        ComboBox<String> product = keywordCombo("部分一致");
        ComboBox<String> raw = keywordCombo("部分一致");
        ComboBox<String> machine = keywordCombo("部分一致");
        ComboBox<String> process = keywordCombo("部分一致");
        Button search = new Button("検索");
        search.getStyleClass().add("btn-reload");
        search.setMaxWidth(Double.MAX_VALUE);
        Button openKensa = new Button("検査表を開く");
        openKensa.getStyleClass().add("btn-reload");
        openKensa.setMaxWidth(Double.MAX_VALUE);
        Tooltip.install(openKensa, new Tooltip("選択した行の検査表を開きます"));
        Button rebuildIndex = new Button("検査表索引を更新");
        rebuildIndex.getStyleClass().add("btn-save-local");
        rebuildIndex.setMaxWidth(Double.MAX_VALUE);
        Label statusMessage = new Label("");
        statusMessage.setWrapText(true);
        statusMessage.getStyleClass().add("top-status");

        Label leftTitle = new Label("検索条件");
        leftTitle.getStyleClass().add("pane-title-left");
        Label hint =
                new Label("キーワードは部分一致。空欄なら納期期間のみ（最新200件）。");
        hint.setWrapText(true);
        hint.getStyleClass().add("paper-main-subtitle");
        HBox dateRow =
                new HBox(8, labeledGrow("納期 From", from), labeledGrow("納期 To", to));
        HBox actionRow = new HBox(8, search, openKensa);
        HBox.setHgrow(search, Priority.ALWAYS);
        HBox.setHgrow(openKensa, Priority.ALWAYS);
        VBox conditions = new VBox(8);
        conditions.setPadding(new Insets(12));
        conditions
                .getChildren()
                .addAll(
                        leftTitle,
                        hint,
                        dateRow,
                        labeled("依頼NO", irai),
                        labeled("製品", product),
                        labeled("投入原反", raw),
                        labeled("機械名", machine),
                        labeled("工程名", process),
                        actionRow,
                        new Separator(),
                        rebuildIndex,
                        statusMessage);

        ScrollPane leftScroll = new ScrollPane(conditions);
        leftScroll.setFitToWidth(true);
        leftScroll.getStyleClass().add("form-scroll-pane");

        Label countLabel = new Label("未検索");
        countLabel.getStyleClass().add("top-status");
        countLabel.setWrapText(true);
        ObservableList<OrderRecord> items = FXCollections.observableArrayList();
        TableView<OrderRecord> table = new TableView<>(items);
        table.getStyleClass().add("juchu-order-search-table");
        BooleanProperty rebuildBusy = new SimpleBooleanProperty(false);
        BooleanProperty searchBusy = new SimpleBooleanProperty(false);
        BooleanProperty openBusy = new SimpleBooleanProperty(false);
        PipelineScanIndex[] planIndex = {PipelineScanIndex.empty()};
        KonanDailyReportLookup[] dailyIndex = {KonanDailyReportLookup.empty()};
        AtomicReference<List<InspectionSheetIndexStore.Row>> kensaIndex =
                new AtomicReference<>(List.of());
        Runnable refreshKeywordCandidates =
                () -> {
                    List<OrderRecord> recs = recordsSupplier.get();
                    if (planIndex[0].planEntriesByTaskId().isEmpty()
                            && planIndex[0].machineNamesByTaskId().isEmpty()) {
                        planIndex[0] = loadPlanIndex(env.get());
                    }
                    setComboCandidates(
                            irai, JuchuOrderSearch.iraiNoCandidates(recs));
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
        irai.setOnShowing(e -> refreshKeywordCandidates.run());
        product.setOnShowing(e -> refreshKeywordCandidates.run());
        raw.setOnShowing(e -> refreshKeywordCandidates.run());
        machine.setOnShowing(e -> refreshKeywordCandidates.run());
        process.setOnShowing(e -> refreshKeywordCandidates.run());
        installKeywordFilter(irai, refreshKeywordCandidates);
        installKeywordFilter(product, refreshKeywordCandidates);
        installKeywordFilter(raw, refreshKeywordCandidates);
        installKeywordFilter(machine, refreshKeywordCandidates);
        installKeywordFilter(process, refreshKeywordCandidates);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        Label placeholder = new Label("条件を指定して検索してください");
        placeholder.getStyleClass().add("juchu-order-search-placeholder");
        placeholder.setWrapText(true);
        placeholder.setMaxWidth(420);
        placeholder.setAlignment(Pos.CENTER);
        placeholder.setTextAlignment(TextAlignment.CENTER);
        table.setPlaceholder(placeholder);
        table.getColumns()
                .addAll(
                        kensaPresenceCol(kensaIndex),
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
                                                machineHaystack(planIndex[0], dailyIndex[0], r))),
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
                        Bindings.createBooleanBinding(
                                () ->
                                        shouldDisableOpenInspectionSheet(
                                                table.getSelectionModel().getSelectedIndex(),
                                                openBusy.get()),
                                table.getSelectionModel().selectedIndexProperty(),
                                table.getSelectionModel().selectedItemProperty(),
                                openBusy));
        rebuildIndex.disableProperty().bind(rebuildBusy);
        search.disableProperty().bind(searchBusy);

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
                                        row.getItem(), env, owner, statusMessage, openBusy);
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
                            openBusy);
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
                                    comboText(irai),
                                    comboText(product),
                                    comboText(raw),
                                    comboText(machine),
                                    comboText(process));
                    Optional<String> err = c.validationError();
                    if (err.isPresent()) {
                        statusMessage.setText(err.get());
                        items.clear();
                        countLabel.getStyleClass().remove("juchu-order-search-count-warn");
                        countLabel.setText("未検索");
                        placeholder.setText("条件を指定して検索してください");
                        return;
                    }
                    List<OrderRecord> src = recordsSupplier.get();
                    List<OrderRecord> snapshot =
                            src == null ? List.of() : new java.util.ArrayList<>(src);
                    Map<String, String> uiSnap = env.get();
                    searchBusy.set(true);
                    statusMessage.setText("検索中…");
                    Task<SearchOutcome> task =
                            new Task<>() {
                                @Override
                                protected SearchOutcome call() {
                                    PipelineScanIndex index = loadPlanIndex(uiSnap);
                                    KonanDailyReportLookup daily =
                                            KonanDailyReportLookup.load(uiSnap, null);
                                    List<InspectionSheetIndexStore.Row> kensa = List.of();
                                    try {
                                        kensa = InspectionSheetOpenService.loadIndex(uiSnap);
                                    } catch (Exception ignored) {
                                        kensa = List.of();
                                    }
                                    JuchuOrderSearch.FilterResult result =
                                            JuchuOrderSearch.filterDetailed(
                                                    snapshot,
                                                    c,
                                                    r -> machineHaystack(index, daily, r),
                                                    r -> planHaystack(index, r, false));
                                    return new SearchOutcome(result, index, daily, kensa);
                                }
                            };
                    task.setOnSucceeded(
                            ev -> {
                                searchBusy.set(false);
                                SearchOutcome outcome = task.getValue();
                                planIndex[0] =
                                        outcome.index() != null
                                                ? outcome.index()
                                                : PipelineScanIndex.empty();
                                dailyIndex[0] =
                                        outcome.daily() != null
                                                ? outcome.daily()
                                                : KonanDailyReportLookup.empty();
                                kensaIndex.set(
                                        outcome.kensa() != null ? outcome.kensa() : List.of());
                                refreshKeywordCandidates.run();
                                JuchuOrderSearch.FilterResult result = outcome.result();
                                items.setAll(result.records());
                                countLabel.getStyleClass().remove("juchu-order-search-count-warn");
                                if (result.records().isEmpty()) {
                                    placeholder.setText("該当する受注はありません");
                                    countLabel.setText("0 件");
                                } else if (result.truncated()) {
                                    countLabel.getStyleClass().add("juchu-order-search-count-warn");
                                    countLabel.setText(
                                            result.records().size()
                                                    + " 件（期間のみ・全 "
                                                    + result.matchCount()
                                                    + " 件中の最新）");
                                } else {
                                    countLabel.setText(result.records().size() + " 件");
                                }
                                statusMessage.setText("");
                                if (!result.records().isEmpty()) {
                                    table.getSelectionModel().selectFirst();
                                    table.requestFocus();
                                }
                            });
                    task.setOnFailed(
                            ev -> {
                                searchBusy.set(false);
                                Throwable ex = task.getException();
                                statusMessage.setText(
                                        "検索に失敗: " + (ex != null ? ex.getMessage() : ""));
                            });
                    Thread worker = new Thread(task, "juchu-order-search");
                    worker.setDaemon(true);
                    worker.setPriority(Thread.MIN_PRIORITY);
                    worker.start();
                });

        openKensa.setOnAction(
                e ->
                        openInspectionSheet(
                                table.getSelectionModel().getSelectedItem(),
                                env,
                                owner,
                                statusMessage,
                                openBusy));
        rebuildIndex.setOnAction(
                e ->
                        rebuildInspectionIndex(
                                env, owner, statusMessage, rebuildBusy, kensaIndex, table));
        bindEnterToSearch(from, search);
        bindEnterToSearch(to, search);
        bindEnterToSearch(irai, search);
        bindEnterToSearch(product, search);
        bindEnterToSearch(raw, search);
        bindEnterToSearch(machine, search);
        bindEnterToSearch(process, search);
        Tooltip.install(rebuildIndex, new Tooltip("検査表フォルダを再走査して索引を作り直します"));

        SplitPane split = new SplitPane(leftScroll, right);
        split.setOrientation(Orientation.HORIZONTAL);
        split.setDividerPositions(0.32);
        SplitPane.setResizableWithParent(leftScroll, Boolean.TRUE);
        SplitPane.setResizableWithParent(right, Boolean.TRUE);

        VBox root = new VBox(split);
        VBox.setVgrow(split, Priority.ALWAYS);
        root.getStyleClass().add("form-tab-container");
        root.setMaxWidth(Double.MAX_VALUE);
        root.setMaxHeight(Double.MAX_VALUE);
        return root;
    }

    private record SearchOutcome(
            JuchuOrderSearch.FilterResult result,
            PipelineScanIndex index,
            KonanDailyReportLookup daily,
            List<InspectionSheetIndexStore.Row> kensa) {}

    static boolean shouldDisableOpenInspectionSheet(int selectedIndex, boolean openBusy) {
        return selectedIndex < 0 || openBusy;
    }

    private static void rebuildInspectionIndex(
            Supplier<Map<String, String>> uiEnv,
            Supplier<Window> owner,
            Label status,
            BooleanProperty busy,
            AtomicReference<List<InspectionSheetIndexStore.Row>> kensaIndex,
            TableView<OrderRecord> table) {
        Map<String, String> ui = uiEnv.get();
        busy.set(true);
        status.setText(
                InspectionSheetIndexProgress.format(
                        InspectionSheetIndexProgress.PHASE_WALK, 0, 0));
        AtomicLong lastUiNs = new AtomicLong(0);
        AtomicBoolean uiPending = new AtomicBoolean(false);
        AtomicReference<String> latestPhase = new AtomicReference<>(InspectionSheetIndexProgress.PHASE_WALK);
        AtomicInteger latestDone = new AtomicInteger(0);
        AtomicInteger latestTotal = new AtomicInteger(0);
        InspectionSheetOpenService.startBackgroundRebuild(
                        ui,
                        (phase, done, total) -> {
                            latestPhase.set(phase);
                            latestDone.set(done);
                            latestTotal.set(total);
                            boolean force = total > 0 && done >= total;
                            long now = System.nanoTime();
                            if (!InspectionSheetIndexProgress.shouldPublishUi(
                                    lastUiNs.get(), now, force)) {
                                return;
                            }
                            lastUiNs.set(now);
                            if (!uiPending.compareAndSet(false, true)) {
                                return;
                            }
                            Platform.runLater(
                                    () -> {
                                        uiPending.set(false);
                                        status.setText(
                                                InspectionSheetIndexProgress.format(
                                                        latestPhase.get(),
                                                        latestDone.get(),
                                                        latestTotal.get()));
                                    });
                        })
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
                                            if (kensaIndex != null && r != null) {
                                                kensaIndex.set(r.rows());
                                            }
                                            if (table != null) {
                                                table.refresh();
                                            }
                                            if (r.rows().isEmpty() && !r.warnings().isEmpty()) {
                                                status.setText(r.warnings().get(0));
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
        busy.set(true);
        status.setText("検査表を検索中…");
        Task<List<InspectionSheetIndexStore.Row>> task =
                new Task<>() {
                    @Override
                    protected List<InspectionSheetIndexStore.Row> call() throws Exception {
                        if (!InspectionSheetOpenService.dirReachable(ui)) {
                            throw new java.io.IOException(
                                    "検査表フォルダ未設定または未到達です。環境変数 "
                                            + AppPaths.KEY_PM_AI_INSPECTION_SHEET_DIR
                                            + " を設定してください: "
                                            + InspectionSheetOpenService.resolveDir(ui));
                        }
                        return InspectionSheetOpenService.find(ui, irai);
                    }
                };
        task.setOnSucceeded(
                e -> {
                    busy.set(false);
                    List<InspectionSheetIndexStore.Row> hits = task.getValue();
                    if (hits == null || hits.isEmpty()) {
                        if (InspectionSheetOpenService.isRebuildInFlight()) {
                            status.setText("検査表索引を更新中です。完了後に再検索してください: " + irai);
                            return;
                        }
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
        t.setPriority(Thread.MIN_PRIORITY);
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
            PipelineScanIndex plan =
                    indexFromShapedPath(AppPaths.resolveShapedAladdinPlanJsonPath(ui));
            PipelineScanIndex actuals =
                    indexFromShapedPath(AppPaths.resolveShapedProcessingActualsJsonPath(ui));
            return AladdinShapedPlanQtyLookup.merge(plan, actuals);
        } catch (RuntimeException ex) {
            return PipelineScanIndex.empty();
        }
    }

    private static PipelineScanIndex indexFromShapedPath(Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            return PipelineScanIndex.empty();
        }
        AladdinShapedPlanQtyLookup.ShapedTable table =
                AladdinShapedPlanQtyLookup.loadShapedTable(path);
        return AladdinShapedPlanQtyLookup.buildPipelineScanIndex(table.headers(), table.rows());
    }

    private static String machineHaystack(
            PipelineScanIndex index, KonanDailyReportLookup daily, OrderRecord record) {
        if (record == null) {
            return "";
        }
        Set<String> names = new LinkedHashSet<>();
        if (index != null) {
            names.addAll(index.machineNamesFor(record.getReqNo()));
        }
        if (daily != null) {
            for (KonanDailyReportLookup.OrderDailyReportEntry entry :
                    daily.entriesForOrder(record.getReqNo())) {
                if (entry != null
                        && entry.machineName() != null
                        && !entry.machineName().isBlank()) {
                    names.add(entry.machineName().strip());
                }
            }
        }
        return String.join(" ", names);
    }

    private static String planHaystack(
            PipelineScanIndex index, OrderRecord record, boolean machine) {
        if (index == null || record == null) {
            return "";
        }
        if (machine) {
            List<String> named = index.machineNamesFor(record.getReqNo());
            if (named != null && !named.isEmpty()) {
                return String.join(" ", named);
            }
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
        combo.getProperties().put(PROP_ALL_CANDIDATES, FXCollections.observableArrayList());
        return combo;
    }

    private static void installKeywordFilter(
            ComboBox<String> combo, Runnable ensureCandidates) {
        combo.getEditor()
                .textProperty()
                .addListener(
                        (obs, oldText, newText) -> {
                            if (Boolean.TRUE.equals(
                                    combo.getProperties().get(PROP_UPDATING_FILTER))) {
                                return;
                            }
                            if (ensureCandidates != null && allCandidates(combo).isEmpty()) {
                                ensureCandidates.run();
                            }
                            applyCandidateFilter(combo);
                            if (!combo.isFocused()) {
                                return;
                            }
                            String typed = newText != null ? newText : "";
                            if (typed.isBlank()) {
                                return;
                            }
                            Platform.runLater(
                                    () -> {
                                        if (combo.isFocused() && !combo.isShowing()) {
                                            combo.show();
                                        }
                                    });
                        });
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
        allCandidates(combo).setAll(names != null ? names : List.of());
        applyCandidateFilter(combo);
    }

    private static ObservableList<String> allCandidates(ComboBox<String> combo) {
        Object existing = combo.getProperties().get(PROP_ALL_CANDIDATES);
        if (existing instanceof ObservableList<?> list) {
            @SuppressWarnings("unchecked")
            ObservableList<String> typed = (ObservableList<String>) list;
            return typed;
        }
        ObservableList<String> created = FXCollections.observableArrayList();
        combo.getProperties().put(PROP_ALL_CANDIDATES, created);
        return created;
    }

    private static void applyCandidateFilter(ComboBox<String> combo) {
        if (combo == null) {
            return;
        }
        String typed = comboText(combo);
        List<String> shown = JuchuOrderSearch.filterCandidates(allCandidates(combo), typed);
        TextField editor = combo.getEditor();
        int caret = editor != null ? editor.getCaretPosition() : 0;
        boolean showing = combo.isShowing();
        combo.getProperties().put(PROP_UPDATING_FILTER, Boolean.TRUE);
        try {
            combo.getItems().setAll(shown);
            if (editor != null) {
                editor.setText(typed);
                int pos = typed == null ? 0 : Math.min(Math.max(caret, 0), typed.length());
                editor.positionCaret(pos);
            }
        } finally {
            combo.getProperties().put(PROP_UPDATING_FILTER, Boolean.FALSE);
        }
        if (showing && !combo.isShowing()) {
            combo.show();
        }
    }

    private static List<String> collectPlanNames(PipelineScanIndex index, boolean machine) {
        if (index == null) {
            return List.of();
        }
        Set<String> names = new LinkedHashSet<>();
        if (machine) {
            for (List<String> rowNames : index.machineNamesByTaskId().values()) {
                if (rowNames == null) {
                    continue;
                }
                for (String value : rowNames) {
                    if (value != null && !value.isBlank()) {
                        names.add(value.strip());
                    }
                }
            }
        }
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

    private static VBox labeledGrow(String caption, javafx.scene.Node field) {
        VBox box = labeled(caption, field);
        HBox.setHgrow(box, Priority.ALWAYS);
        return box;
    }

    private static void bindEnterToSearch(DatePicker picker, Button search) {
        if (picker == null || search == null) {
            return;
        }
        javafx.event.EventHandler<javafx.scene.input.KeyEvent> fire =
                e -> {
                    if (e.getCode() == KeyCode.ENTER) {
                        search.fire();
                        e.consume();
                    }
                };
        picker.setOnKeyPressed(fire);
        if (picker.getEditor() != null) {
            picker.getEditor().setOnKeyPressed(fire);
        }
    }

    private static void bindEnterToSearch(ComboBox<String> combo, Button search) {
        if (combo == null || search == null || combo.getEditor() == null) {
            return;
        }
        combo.getEditor()
                .setOnKeyPressed(
                        e -> {
                            if (e.getCode() != KeyCode.ENTER || combo.isShowing()) {
                                return;
                            }
                            search.fire();
                            e.consume();
                        });
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

    private static TableColumn<OrderRecord, String> kensaPresenceCol(
            AtomicReference<List<InspectionSheetIndexStore.Row>> kensaIndex) {
        TableColumn<OrderRecord, String> column =
                col(
                        "検査表",
                        r ->
                                InspectionSheetLookup.presenceLabel(
                                        InspectionSheetLookup.hasSheet(
                                                kensaIndex.get(), r.getReqNo())));
        column.setMinWidth(72);
        column.setPrefWidth(80);
        column.setMaxWidth(96);
        column.setResizable(false);
        column.setStyle("-fx-alignment: CENTER;");
        column.setCellFactory(
                col ->
                        new TableCell<>() {
                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                getStyleClass().remove("juchu-kensa-present");
                                if (empty || item == null || item.isBlank()) {
                                    setText(null);
                                    return;
                                }
                                setText(item);
                                if (!getStyleClass().contains("juchu-kensa-present")) {
                                    getStyleClass().add("juchu-kensa-present");
                                }
                            }
                        });
        return column;
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
