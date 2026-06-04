package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Orientation;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.SplitPane;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo;

/**
 * 依頼書入力タブ内「後加工商品マスタ編集」カードの UI。
 */
public final class PostProcessingProductMasterEditorPane {

    public record Context(
            Supplier<Map<String, String>> uiEnv,
            Supplier<java.util.List<ProductInfo>> integratedProductCatalog,
            Supplier<PostProcessingProductMasterSearch.MasterReferencePrefixFilters>
                    masterCandidatePrefixFilters,
            Runnable invalidateReferenceCache,
            Consumer<String> log) {}

    private static final List<String> UPLOAD_TABLE_COLUMNS =
            List.of("商品コード", "商品名1", "発泡体品番", "発泡体幅", "発泡体長さ");

    /** 参照マスタ検索結果テーブルの列（依頼書候補表示に近い並び）。 */
    private static final List<SearchResultColumn> SEARCH_RESULT_COLUMNS =
            List.of(
                    new SearchResultColumn("商品コード", 140, PostProcessingProductMasterEditorPane::shohinCode),
                    new SearchResultColumn("商品名1", 220, PostProcessingProductMasterEditorPane::shohinName1),
                    new SearchResultColumn("品番", 80, PostProcessingProductMasterEditorPane::foamPartNo),
                    new SearchResultColumn("品名", 64, PostProcessingProductMasterEditorPane::foamName),
                    new SearchResultColumn("タイプ", 72, h -> rowColumn(h, "発泡体タイプ")),
                    new SearchResultColumn("幅", 72, h -> rowColumn(h, "発泡体幅")),
                    new SearchResultColumn("長さ", 72, h -> rowColumn(h, "発泡体長さ")),
                    new SearchResultColumn("色", 56, h -> rowColumn(h, "発泡体色")));

    private PostProcessingProductMasterEditorPane() {}

    /** 依頼書入力の専用タブ用（横幅いっぱい）。ディスク読込はバックグラウンド。 */
    public static VBox buildTabContent(Window owner, Context ctx) {
        VBox content = buildContent(owner, ctx, false, true);
        content.getStyleClass().add("form-tab-container");
        content.setFillWidth(true);
        content.setMaxWidth(Double.MAX_VALUE);
        VBox.setVgrow(content, Priority.ALWAYS);
        return content;
    }

    /** 【設定】タブ内カード用（幅上限あり）。 */
    public static VBox buildCard(Window owner, Context ctx, double maxCardWidth) {
        VBox card = buildContent(owner, ctx, true, true);
        card.getStyleClass().add("settings-card");
        card.setMaxWidth(maxCardWidth);
        card.setPrefWidth(maxCardWidth);
        return card;
    }

    private static VBox buildContent(
            Window owner, Context ctx, boolean compactCardTitle, boolean deferInitialLoad) {
        Supplier<Map<String, String>> uiEnv = ctx.uiEnv();
        Consumer<String> log = ctx.log() != null ? ctx.log() : s -> {};

        Label statusLabel = new Label("参照マスタを読み込んでください。");
        statusLabel.setStyle("-fx-font-size: 11px; -fx-font-weight: bold;");
        statusLabel.setWrapText(true);

        TextField refPathField = new TextField();
        refPathField.setEditable(false);
        TextField uploadPathField = new TextField();
        HBox.setHgrow(uploadPathField, Priority.ALWAYS);

        TextField fCode = new TextField();
        TextField fPart = new TextField();
        TextField fType = new TextField();
        TextField fLength = new TextField();
        TextField fName = new TextField();
        for (TextField tf : List.of(fCode, fPart, fType, fLength, fName)) {
            tf.setPromptText("部分一致");
            tf.setStyle("-fx-font-size: 11px;");
        }

        TableView<PostProcessingProductMasterIo.SearchHit> searchResults =
                buildSearchResultTable(compactCardTitle);

        Map<String, TextField> fieldByColumn = new LinkedHashMap<>();
        Map<String, Label> codeNameLabelByColumn = new LinkedHashMap<>();
        Map<String, ComboBox<String>> codeComboByColumn = new LinkedHashMap<>();
        AtomicReference<PostProcessingKouteiNaiyoMasterLookup.Snapshot> kouteiNaiyoLookupRef =
                new AtomicReference<>(PostProcessingKouteiNaiyoMasterLookup.Snapshot.empty());
        AtomicReference<PostProcessingShuruiMasterLookup.Snapshot> shuruiLookupRef =
                new AtomicReference<>(PostProcessingShuruiMasterLookup.Snapshot.empty());
        AtomicReference<PostProcessingKeiriBunruiMasterLookup.Snapshot> keiriBunruiLookupRef =
                new AtomicReference<>(PostProcessingKeiriBunruiMasterLookup.Snapshot.empty());
        AtomicReference<PostProcessingYotoMasterLookup.Snapshot> yotoLookupRef =
                new AtomicReference<>(PostProcessingYotoMasterLookup.Snapshot.empty());
        AtomicReference<PostProcessingShohinBunrui4MasterLookup.Snapshot> bunrui4LookupRef =
                new AtomicReference<>(PostProcessingShohinBunrui4MasterLookup.Snapshot.empty());
        AtomicReference<PostProcessingZaikoBunruiMasterLookup.Snapshot> zaikoBunruiLookupRef =
                new AtomicReference<>(PostProcessingZaikoBunruiMasterLookup.Snapshot.empty());
        AtomicReference<PostProcessingPlanMachineLookup.Snapshot> planMachineLookupRef =
                new AtomicReference<>(PostProcessingPlanMachineLookup.Snapshot.empty());
        TabPane formTabs = new TabPane();
        formTabs.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);
        formTabs.setMinHeight(compactCardTitle ? 280 : 360);
        VBox.setVgrow(formTabs, Priority.ALWAYS);

        List<String> referenceHeaders = new ArrayList<>();
        AtomicReference<PostProcessingProductMasterEditorModel> editorModelRef =
                new AtomicReference<>(new PostProcessingProductMasterEditorModel(List.of()));
        Map<String, String> templateRow = new LinkedHashMap<>();

        AtomicBoolean suppressFormDirtyTracking = new AtomicBoolean(false);
        AtomicBoolean formPreparing = new AtomicBoolean(deferInitialLoad);
        AtomicBoolean formDirty = new AtomicBoolean(false);
        LinkedHashMap<String, String> formBaseline = new LinkedHashMap<>();
        AtomicReference<Runnable> refreshInteractionStateRef = new AtomicReference<>(() -> {});
        AtomicReference<Runnable> captureFormBaselineRef = new AtomicReference<>(() -> {});
        AtomicReference<Runnable> updateDirtyFromFormRef = new AtomicReference<>(() -> {});
        AtomicReference<Runnable> wireFormDirtyTrackingRef = new AtomicReference<>(() -> {});
        AtomicReference<Button> btnDuplicateCheckRef = new AtomicReference<>();

        ObservableList<Map<String, String>> uploadRows = FXCollections.observableArrayList();
        TableView<Map<String, String>> uploadTable = new TableView<>(uploadRows);
        double uploadTableHeight = compactCardTitle ? 120 : 160;
        uploadTable.setPrefHeight(uploadTableHeight);
        uploadTable.setMinHeight(uploadTableHeight);
        uploadTable.setMaxHeight(uploadTableHeight);
        uploadTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        for (String col : UPLOAD_TABLE_COLUMNS) {
            TableColumn<Map<String, String>, String> tc = new TableColumn<>(col);
            tc.setCellValueFactory(cd -> new javafx.beans.property.SimpleStringProperty(
                    cd.getValue() != null ? cd.getValue().getOrDefault(col, "") : ""));
            uploadTable.getColumns().add(tc);
        }

        Runnable refreshPathsFromEnv =
                () -> {
                    Map<String, String> ui = uiEnv.get();
                    Path ref = PostProcessingProductMasterIo.resolveReferencePath(ui);
                    Path up = PostProcessingProductMasterIo.resolveUploadPath(ui);
                    refPathField.setText(ref.toString());
                    uploadPathField.setText(up.toString());
                };

        Runnable runShohinCodeDuplicateCheck =
                () -> {
                    if (formPreparing.get()) {
                        return;
                    }
                    TextField codeField = fieldByColumn.get("商品コード");
                    if (codeField == null) {
                        return;
                    }
                    String code = codeField.getText().trim();
                    editorModelRef.get().set("商品コード", code);
                    Path ref = Path.of(refPathField.getText().trim());
                    Map<String, String> excludeUploadRow =
                            uploadTable.getSelectionModel().getSelectedItem();
                    statusLabel.setText("商品コードを確認中…");
                    Thread worker =
                            new Thread(
                                    () -> {
                                        try {
                                            PostProcessingProductMasterDuplicateCheck.Result result =
                                                    PostProcessingProductMasterDuplicateCheck.check(
                                                            code,
                                                            ref,
                                                            uploadRows,
                                                            excludeUploadRow);
                                            Platform.runLater(
                                                    () -> {
                                                        if (result.usable()) {
                                                            statusLabel.setText(
                                                                    String.join(
                                                                            " ",
                                                                            result.messages()));
                                                        } else {
                                                            statusLabel.setText("重複あり");
                                                            showError(
                                                                    "重複チェック",
                                                                    String.join(
                                                                            "\n",
                                                                            result.messages()));
                                                        }
                                                    });
                                        } catch (Exception ex) {
                                            Platform.runLater(
                                                    () -> {
                                                        statusLabel.setText(
                                                                "重複チェック失敗: "
                                                                        + ex.getMessage());
                                                        showError(
                                                                "重複チェック", ex.getMessage());
                                                    });
                                        }
                                    },
                                    "postproc-master-dup-check");
                    worker.setDaemon(true);
                    worker.start();
                };

        Runnable rebuildForm =
                () -> {
                    formPreparing.set(true);
                    refreshInteractionStateRef.get().run();
                    fieldByColumn.clear();
                    codeNameLabelByColumn.clear();
                    codeComboByColumn.clear();
                    formTabs.getTabs().clear();
                    PostProcessingKouteiNaiyoMasterLookup.Snapshot kouteiNaiyo =
                            kouteiNaiyoLookupRef.get();
                    PostProcessingShuruiMasterLookup.Snapshot shurui = shuruiLookupRef.get();
                    PostProcessingKeiriBunruiMasterLookup.Snapshot keiriBunrui =
                            keiriBunruiLookupRef.get();
                    PostProcessingYotoMasterLookup.Snapshot yoto = yotoLookupRef.get();
                    PostProcessingShohinBunrui4MasterLookup.Snapshot bunrui4 =
                            bunrui4LookupRef.get();
                    PostProcessingZaikoBunruiMasterLookup.Snapshot zaikoBunrui =
                            zaikoBunruiLookupRef.get();
                    PostProcessingPlanMachineLookup.Snapshot planMachine =
                            planMachineLookupRef.get();
                    for (PostProcessingProductMasterColumnGroups.TabGroup group :
                            PostProcessingProductMasterColumnGroups.tabGroups()) {
                        GridPane grid = new GridPane();
                        grid.setHgap(8);
                        grid.setVgap(6);
                        grid.setPadding(new Insets(8));
                        ColumnConstraints labelCol = new ColumnConstraints();
                        labelCol.setMinWidth(
                                PostProcessingProductMasterCodeFieldRows.labelColumnPrefWidth());
                        labelCol.setPrefWidth(
                                PostProcessingProductMasterCodeFieldRows.labelColumnPrefWidth());
                        labelCol.setMaxWidth(
                                PostProcessingProductMasterCodeFieldRows.labelColumnPrefWidth()
                                        + 16);
                        ColumnConstraints fieldCol = new ColumnConstraints();
                        fieldCol.setPrefWidth(
                                PostProcessingProductMasterCodeFieldRows.fieldColumnPrefWidth());
                        fieldCol.setMaxWidth(
                                PostProcessingProductMasterCodeFieldRows.fieldColumnMaxWidth());
                        fieldCol.setHgrow(Priority.NEVER);
                        grid.getColumnConstraints().addAll(labelCol, fieldCol);
                        grid.setMaxWidth(
                                PostProcessingProductMasterCodeFieldRows.fieldColumnMaxWidth()
                                        + PostProcessingProductMasterCodeFieldRows
                                                .labelColumnPrefWidth()
                                        + 32);
                        int row = 0;
                        for (String colName : group.columnNames()) {
                            if (!editorModelRef.get().headers().contains(colName)) {
                                continue;
                            }
                            PostProcessingProductMasterCodeFieldRows.addColumnRow(
                                    grid,
                                    row++,
                                    colName,
                                    editorModelRef.get(),
                                    kouteiNaiyo,
                                    shurui,
                                    keiriBunrui,
                                    yoto,
                                    bunrui4,
                                    zaikoBunrui,
                                    planMachine,
                                    fieldByColumn,
                                    codeNameLabelByColumn,
                                    codeComboByColumn,
                                    col -> updateDirtyFromFormRef.get().run(),
                                    runShohinCodeDuplicateCheck,
                                    btnDuplicateCheckRef);
                        }
                        ScrollPane sp = new ScrollPane(grid);
                        sp.setFitToWidth(true);
                        Tab tab = new Tab(group.tabTitle(), sp);
                        formTabs.getTabs().add(tab);
                    }
                    wireFormDirtyTrackingRef.get().run();
                    formPreparing.set(false);
                    captureFormBaselineRef.get().run();
                };

        Runnable reloadKouteiNaiyoLookup =
                () -> {
                    try {
                        kouteiNaiyoLookupRef.set(
                                PostProcessingKouteiNaiyoMasterLookup.snapshot(uiEnv.get()));
                    } catch (IOException ex) {
                        kouteiNaiyoLookupRef.set(
                                PostProcessingKouteiNaiyoMasterLookup.Snapshot.empty());
                        log.accept("[postproc-master] koutei/naiyo: " + ex.getMessage());
                    }
                };

        Runnable reloadShuruiLookup =
                () -> {
                    try {
                        shuruiLookupRef.set(
                                PostProcessingShuruiMasterLookup.snapshot(uiEnv.get()));
                    } catch (IOException ex) {
                        shuruiLookupRef.set(PostProcessingShuruiMasterLookup.Snapshot.empty());
                        log.accept("[postproc-master] shurui: " + ex.getMessage());
                    }
                };

        Runnable reloadKeiriBunruiLookup =
                () -> {
                    try {
                        keiriBunruiLookupRef.set(
                                PostProcessingKeiriBunruiMasterLookup.snapshot(uiEnv.get()));
                    } catch (IOException ex) {
                        keiriBunruiLookupRef.set(
                                PostProcessingKeiriBunruiMasterLookup.Snapshot.empty());
                        log.accept("[postproc-master] keiri: " + ex.getMessage());
                    }
                };

        Runnable reloadYotoLookup =
                () -> {
                    try {
                        yotoLookupRef.set(
                                PostProcessingYotoMasterLookup.snapshot(uiEnv.get()));
                    } catch (IOException ex) {
                        yotoLookupRef.set(PostProcessingYotoMasterLookup.Snapshot.empty());
                        log.accept("[postproc-master] yoto: " + ex.getMessage());
                    }
                };

        Runnable reloadBunrui4Lookup =
                () -> {
                    try {
                        bunrui4LookupRef.set(
                                PostProcessingShohinBunrui4MasterLookup.snapshot(uiEnv.get()));
                    } catch (IOException ex) {
                        bunrui4LookupRef.set(
                                PostProcessingShohinBunrui4MasterLookup.Snapshot.empty());
                        log.accept("[postproc-master] bunrui4: " + ex.getMessage());
                    }
                };

        Runnable reloadZaikoBunruiLookup =
                () -> {
                    try {
                        zaikoBunruiLookupRef.set(
                                PostProcessingZaikoBunruiMasterLookup.snapshot(uiEnv.get()));
                    } catch (IOException ex) {
                        zaikoBunruiLookupRef.set(
                                PostProcessingZaikoBunruiMasterLookup.Snapshot.empty());
                        log.accept("[postproc-master] zaiko: " + ex.getMessage());
                    }
                };

        Runnable reloadPlanMachineLookup =
                () -> {
                    try {
                        planMachineLookupRef.set(
                                PostProcessingPlanMachineLookup.snapshot(uiEnv.get()));
                    } catch (IOException ex) {
                        planMachineLookupRef.set(
                                PostProcessingPlanMachineLookup.Snapshot.empty());
                        log.accept("[postproc-master] plan machine: " + ex.getMessage());
                    }
                };

        Runnable reloadAllMasterLookups =
                () -> {
                    reloadKouteiNaiyoLookup.run();
                    reloadShuruiLookup.run();
                    reloadKeiriBunruiLookup.run();
                    reloadYotoLookup.run();
                    reloadBunrui4Lookup.run();
                    reloadZaikoBunruiLookup.run();
                    reloadPlanMachineLookup.run();
                };

        Runnable syncFormFieldsFromModel =
                () -> {
                    suppressFormDirtyTracking.set(true);
                    try {
                        for (Map.Entry<String, TextField> e : fieldByColumn.entrySet()) {
                            e.getValue().setText(editorModelRef.get().get(e.getKey()));
                        }
                    } finally {
                        suppressFormDirtyTracking.set(false);
                    }
                };

        Runnable resetFormFromModel =
                () -> {
                    syncFormFieldsFromModel.run();
                    captureFormBaselineRef.get().run();
                };

        Runnable loadReferenceHeaders =
                () -> {
                    formPreparing.set(true);
                    refreshInteractionStateRef.get().run();
                    try {
                        Path ref = Path.of(refPathField.getText().trim());
                        if (!Files.isRegularFile(ref)) {
                            statusLabel.setText("参照マスタが見つかりません: " + ref);
                            return;
                        }
                        referenceHeaders.clear();
                        referenceHeaders.addAll(PostProcessingProductMasterIo.readHeaders(ref));
                        editorModelRef.set(
                                new PostProcessingProductMasterEditorModel(referenceHeaders));
                        reloadAllMasterLookups.run();
                        rebuildForm.run();
                        statusLabel.setText(
                                masterLookupStatusMessage(
                                        referenceHeaders.size(),
                                        kouteiNaiyoLookupRef.get(),
                                        shuruiLookupRef.get(),
                                        keiriBunruiLookupRef.get(),
                                        yotoLookupRef.get(),
                                        bunrui4LookupRef.get(),
                                        zaikoBunruiLookupRef.get(),
                                        planMachineLookupRef.get()));
                    } catch (Exception ex) {
                        formPreparing.set(false);
                        refreshInteractionStateRef.get().run();
                        statusLabel.setText("参照マスタ読込失敗: " + ex.getMessage());
                        log.accept("[postproc-master] ref headers: " + ex.getMessage());
                    }
                };

        Runnable loadUploadFile =
                () -> {
                    try {
                        Path up = Path.of(uploadPathField.getText().trim());
                        if (!Files.isRegularFile(up)) {
                            uploadRows.clear();
                            statusLabel.setText("アップロード用ファイルがありません（新規作成可）");
                            return;
                        }
                        if (referenceHeaders.isEmpty()) {
                            loadReferenceHeaders.run();
                        }
                        applyUploadSheet(
                                PostProcessingProductMasterIo.readUploadWorkbook(up),
                                referenceHeaders,
                                uploadRows,
                                statusLabel);
                    } catch (Exception ex) {
                        showError("読込エラー", ex.getMessage());
                        log.accept("[postproc-master] upload read: " + ex.getMessage());
                    }
                };

        Runnable saveUpload =
                () -> {
                    try {
                        if (referenceHeaders.isEmpty()) {
                            loadReferenceHeaders.run();
                        }
                        for (Map.Entry<String, TextField> e : fieldByColumn.entrySet()) {
                            editorModelRef.get().set(e.getKey(), e.getValue().getText());
                        }
                        Path up = Path.of(uploadPathField.getText().trim());
                        List<Map<String, String>> rows = new ArrayList<>(uploadRows);
                        PostProcessingProductMasterIo.writeUploadWorkbook(
                                up, referenceHeaders, rows);
                        statusLabel.setText("保存しました: " + up);
                        log.accept("[postproc-master] saved " + up + " rows=" + rows.size());
                    } catch (Exception ex) {
                        showError("保存エラー", ex.getMessage());
                        log.accept("[postproc-master] save: " + ex.getMessage());
                    }
                };

        Button btnSearch = new Button("検索");
        btnSearch.getStyleClass().add("btn-reload");
        btnSearch.setOnAction(
                e -> {
                    if (formPreparing.get() || formDirty.get()) {
                        return;
                    }
                    statusLabel.setText("検索中...");
                    Path ref = Path.of(refPathField.getText().trim());
                    PostProcessingProductMasterIo.SearchFilter filter =
                            new PostProcessingProductMasterIo.SearchFilter(
                                    fCode.getText(),
                                    fPart.getText(),
                                    fType.getText(),
                                    fLength.getText(),
                                    fName.getText());
                    Thread t =
                            new Thread(
                                    () -> {
                                        try {
                                            List<ProductInfo> catalog =
                                                    ctx.integratedProductCatalog() != null
                                                            ? ctx.integratedProductCatalog().get()
                                                            : List.of();
                                            PostProcessingProductMasterSearch.MasterReferencePrefixFilters
                                                    prefixes =
                                                            ctx.masterCandidatePrefixFilters() != null
                                                                    ? ctx.masterCandidatePrefixFilters()
                                                                            .get()
                                                                    : PostProcessingProductMasterSearch
                                                                            .MasterReferencePrefixFilters
                                                                            .none();
                                            List<PostProcessingProductMasterIo.SearchHit> hits =
                                                    PostProcessingProductMasterIo.searchReference(
                                                            ref, filter, 200, catalog, prefixes);
                                            Platform.runLater(
                                                    () -> {
                                                        if (formPreparing.get() || formDirty.get()) {
                                                            return;
                                                        }
                                                        searchResults
                                                                .setItems(
                                                                        FXCollections
                                                                                .observableArrayList(
                                                                                        hits));
                                                        statusLabel.setText(
                                                                "検索結果 "
                                                                        + hits.size()
                                                                        + " 件");
                                                    });
                                        } catch (Exception ex) {
                                            Platform.runLater(
                                                    () ->
                                                            statusLabel.setText(
                                                                    "検索失敗: "
                                                                            + ex.getMessage()));
                                        }
                                    },
                                    "postproc-master-search");
                    t.setDaemon(true);
                    t.start();
                });

        Runnable applySearchHitToForm =
                () -> {
                    if (formPreparing.get() || formDirty.get()) {
                        return;
                    }
                    PostProcessingProductMasterIo.SearchHit hit =
                            searchResults.getSelectionModel().getSelectedItem();
                    if (hit == null) {
                        return;
                    }
                    templateRow.clear();
                    templateRow.putAll(hit.rowByColumn());
                    suppressFormDirtyTracking.set(true);
                    try {
                        editorModelRef.get().applyTemplateRow(templateRow);
                        syncFormFieldsFromModel.run();
                    } finally {
                        suppressFormDirtyTracking.set(false);
                    }
                    captureFormBaselineRef.get().run();
                    statusLabel.setText("雛形をフォームに反映しました（商品コードは手修正）。");
                };
        searchResults
                .getSelectionModel()
                .selectedItemProperty()
                .addListener((obs, old, hit) -> applySearchHitToForm.run());

        Button btnCancelEdit = new Button("編集キャンセル");
        btnCancelEdit.getStyleClass().add("btn-settings-del");
        btnCancelEdit.setDisable(true);
        btnCancelEdit.setTooltip(
                new Tooltip("右側フォームの未保存の変更を破棄し、直前の確定状態に戻します。"));
        btnCancelEdit.setOnAction(
                e -> {
                    if (formPreparing.get() || !formDirty.get()) {
                        return;
                    }
                    suppressFormDirtyTracking.set(true);
                    try {
                        for (Map.Entry<String, String> entry : formBaseline.entrySet()) {
                            editorModelRef.get().set(entry.getKey(), entry.getValue());
                        }
                        syncFormFieldsFromModel.run();
                    } finally {
                        suppressFormDirtyTracking.set(false);
                    }
                    captureFormBaselineRef.get().run();
                    statusLabel.setText("編集をキャンセルしました。");
                });

        Button btnAddRow = new Button("③ 一覧へ追加");
        btnAddRow.getStyleClass().add("btn-settings-add");
        btnAddRow.setTooltip(
                new Tooltip(
                        "右フォームの内容を下の一覧に1行追加します（未保存）。"
                                + " Excel へ書き出すには「④ Excel保存」を押してください。"));
        btnAddRow.setOnAction(
                e -> {
                    for (Map.Entry<String, TextField> ent : fieldByColumn.entrySet()) {
                        editorModelRef.get().set(ent.getKey(), ent.getValue().getText());
                    }
                    List<String> codes = new ArrayList<>();
                    for (Map<String, String> r : uploadRows) {
                        codes.add(r.getOrDefault("商品コード", ""));
                    }
                    codes.add(editorModelRef.get().get("商品コード"));
                    var v = editorModelRef.get().validateForUpload(codes);
                    if (!v.ok()) {
                        showError("検証", String.join("\n", v.messages()));
                        return;
                    }
                    uploadRows.add(new LinkedHashMap<>(editorModelRef.get().snapshot()));
                    statusLabel.setText("行を追加しました（未保存）。");
                });

        Button btnRemoveRow = new Button("削除");
        btnRemoveRow.getStyleClass().add("btn-settings-del");
        btnRemoveRow.setTooltip(new Tooltip("一覧で選択した行を削除します（未保存）。"));
        btnRemoveRow.setOnAction(
                e -> {
                    Map<String, String> sel = uploadTable.getSelectionModel().getSelectedItem();
                    if (sel != null) {
                        uploadRows.remove(sel);
                    }
                });

        Button btnDupRow = new Button("複製");
        btnDupRow.getStyleClass().add("btn-reload");
        btnDupRow.setTooltip(new Tooltip("一覧で選択した行を複製して末尾に追加します（未保存）。"));
        btnDupRow.setOnAction(
                e -> {
                    Map<String, String> sel = uploadTable.getSelectionModel().getSelectedItem();
                    if (sel != null) {
                        uploadRows.add(new LinkedHashMap<>(sel));
                    }
                });

        uploadTable
                .getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (obs, o, row) -> {
                            if (row == null || formPreparing.get() || formDirty.get()) {
                                return;
                            }
                            editorModelRef.get().loadFromRow(row);
                            resetFormFromModel.run();
                        });

        Button btnNewUpload = new Button("① 空ファイル作成");
        btnNewUpload.getStyleClass().add("btn-reload");
        btnNewUpload.setTooltip(
                new Tooltip(
                        "参照マスタと同じ見出しの空 xlsx を作成します（初回のみ。既存ファイルがある場合は不要）。"));
        btnNewUpload.setOnAction(
                e -> {
                    try {
                        if (referenceHeaders.isEmpty()) {
                            loadReferenceHeaders.run();
                        }
                        Path ref = Path.of(refPathField.getText().trim());
                        Path up = Path.of(uploadPathField.getText().trim());
                        PostProcessingProductMasterIo.createEmptyUploadFromReference(ref, up);
                        uploadRows.clear();
                        statusLabel.setText("空のアップロード用ファイルを作成しました。");
                    } catch (Exception ex) {
                        showError("新規作成エラー", ex.getMessage());
                    }
                });

        Button btnBrowseUpload = new Button("ファイルを選ぶ…");
        btnBrowseUpload.getStyleClass().add("btn-reload");
        btnBrowseUpload.setTooltip(new Tooltip("アップロード用 Excel の保存先パスを選びます。"));
        btnBrowseUpload.setOnAction(
                e -> {
                    FileChooser ch = new FileChooser();
                    ch.setTitle("アップロード用 後加工商品マスタ");
                    ch.getExtensionFilters()
                            .add(new FileChooser.ExtensionFilter("Excel", "*.xlsx"));
                    var f = ch.showOpenDialog(owner);
                    if (f != null) {
                        uploadPathField.setText(f.getAbsolutePath());
                        loadUploadFile.run();
                    }
                });

        Button btnSave = new Button("④ Excel保存");
        btnSave.getStyleClass().add("btn-transfer");
        btnSave.setTooltip(
                new Tooltip(
                        "一覧の全行をアップロード用 Excel（"
                                + PostProcessingProductMasterIo.DEFAULT_UPLOAD_FILE_NAME
                                + " 等）に書き込みます。"));
        btnSave.setOnAction(e -> saveUpload.run());

        Button btnReservedAladdinMasterUpsert =
                RequestFormReservedButton.ALADDIN_MASTER_UPSERT.toButton(statusLabel);

        Button btnReloadUpload = new Button("再読込");
        btnReloadUpload.getStyleClass().add("btn-reload");
        btnReloadUpload.setTooltip(new Tooltip("アップロード用 Excel から一覧を読み直します（未保存の変更は失われます）。"));
        btnReloadUpload.setOnAction(e -> loadUploadFile.run());

        refreshPathsFromEnv.run();

        Label note =
                new Label(
                        "※ Aladdin の単価タブ（上代・販売・購買単価等）は本 xlsx に含まれません。"
                                + " 本番マスタ（後加工商品マスタ.xlsx）への直接上書きはしません。");
        note.setWrapText(true);
        note.setStyle("-fx-font-size: 10px;");

        Label uploadWorkflow =
                new Label(
                        "【手順】① 空ファイル作成（初回）"
                                + " → ② 左で雛形検索・右フォーム入力"
                                + " → ③ 一覧へ追加"
                                + " → ④ Excel保存");
        uploadWorkflow.setWrapText(true);
        uploadWorkflow.getStyleClass().add("postproc-master-workflow-hint");

        Region filterPaneSpacer = new Region();
        VBox.setVgrow(filterPaneSpacer, Priority.ALWAYS);

        VBox filterBox = new VBox(6);
        filterBox.setFillWidth(true);
        filterBox.getChildren()
                .addAll(
                        new Label("参照マスタ検索（雛形）"),
                        gridFilterRow("商品コード:", fCode),
                        gridFilterRow("品番:", fPart),
                        gridFilterRow("タイプ:", fType),
                        gridFilterRow("長さ:", fLength),
                        gridFilterRow("品名:", fName),
                        btnSearch,
                        searchResults,
                        new HBox(6, btnCancelEdit),
                        filterPaneSpacer);
        filterBox.setMinWidth(260);
        filterBox.setPrefWidth(Region.USE_COMPUTED_SIZE);
        filterBox.setMaxWidth(Double.MAX_VALUE);

        HBox pathRow1 = new HBox(6, new Label("参照:"), refPathField);
        HBox.setHgrow(refPathField, Priority.ALWAYS);
        HBox pathRow2 = new HBox(6, new Label("アップロード用:"), uploadPathField, btnBrowseUpload);
        HBox.setHgrow(uploadPathField, Priority.ALWAYS);

        Label fileActionsCaption = new Label("ファイル");
        fileActionsCaption.getStyleClass().add("postproc-master-toolbar-caption");
        HBox fileActionRow =
                new HBox(8, fileActionsCaption, btnNewUpload, btnReloadUpload);
        fileActionRow.setAlignment(Pos.CENTER_LEFT);
        fileActionRow.getStyleClass().add("postproc-master-toolbar");

        Region tableActionSpacer = new Region();
        HBox.setHgrow(tableActionSpacer, Priority.ALWAYS);
        HBox tableActionRow =
                new HBox(
                        10,
                        toolbarGroup("行", btnAddRow, btnRemoveRow, btnDupRow),
                        tableActionSpacer,
                        toolbarGroup("確定", btnSave, btnReservedAladdinMasterUpsert));
        tableActionRow.setAlignment(Pos.CENTER_LEFT);
        tableActionRow.getStyleClass().add("postproc-master-toolbar");

        VBox topPaths = new VBox(4, pathRow1, pathRow2, fileActionRow);

        VBox editorPane = new VBox(formTabs);
        editorPane.setFillWidth(true);
        editorPane.setMinWidth(280);
        VBox.setVgrow(formTabs, Priority.ALWAYS);

        Runnable refreshInteractionState =
                () -> {
                    boolean preparing = formPreparing.get();
                    boolean dirty = formDirty.get();
                    boolean searchLocked = preparing || dirty;

                    for (TextField tf : List.of(fCode, fPart, fType, fLength, fName)) {
                        tf.setDisable(searchLocked);
                    }
                    btnSearch.setDisable(searchLocked);
                    searchResults.setDisable(searchLocked);
                    btnCancelEdit.setDisable(preparing || !dirty);

                    uploadTable.setDisable(preparing || dirty);
                    btnAddRow.setDisable(preparing);
                    btnRemoveRow.setDisable(preparing);
                    btnDupRow.setDisable(preparing);
                    btnNewUpload.setDisable(preparing);
                    btnBrowseUpload.setDisable(preparing);
                    btnReloadUpload.setDisable(preparing);
                    btnSave.setDisable(preparing);
                    btnReservedAladdinMasterUpsert.setDisable(preparing);
                    formTabs.setDisable(preparing);
                    editorPane.setDisable(preparing);
                    filterBox.setDisable(preparing);
                    topPaths.setDisable(preparing);
                    tableActionRow.setDisable(preparing);
                    Button dupBtn = btnDuplicateCheckRef.get();
                    if (dupBtn != null) {
                        dupBtn.setDisable(preparing);
                    }
                };

        Runnable updateDirtyFromForm =
                () -> {
                    if (suppressFormDirtyTracking.get() || formPreparing.get()) {
                        return;
                    }
                    Map<String, String> current = editorModelRef.get().snapshot();
                    formDirty.set(!Objects.equals(current, formBaseline));
                    refreshInteractionState.run();
                };

        Runnable captureFormBaseline =
                () -> {
                    for (Map.Entry<String, TextField> e : fieldByColumn.entrySet()) {
                        editorModelRef.get().set(e.getKey(), e.getValue().getText());
                    }
                    formBaseline.clear();
                    formBaseline.putAll(editorModelRef.get().snapshot());
                    formDirty.set(false);
                    refreshInteractionState.run();
                };

        refreshInteractionStateRef.set(refreshInteractionState);
        captureFormBaselineRef.set(captureFormBaseline);
        updateDirtyFromFormRef.set(updateDirtyFromForm);

        SplitPane mainSplit = new SplitPane(filterBox, editorPane);
        mainSplit.setOrientation(Orientation.HORIZONTAL);
        SplitPane.setResizableWithParent(filterBox, Boolean.TRUE);
        SplitPane.setResizableWithParent(editorPane, Boolean.TRUE);
        mainSplit.setDividerPositions(0.4);
        VBox.setVgrow(filterBox, Priority.ALWAYS);

        VBox root = new VBox(10);
        root.setMaxWidth(Double.MAX_VALUE);
        Label title = new Label("後加工商品マスタ編集");
        title.getStyleClass().add(compactCardTitle ? "settings-card-title" : "paper-main-title");
        root.getChildren()
                .addAll(
                        title,
                        note,
                        uploadWorkflow,
                        topPaths,
                        new Label("アップロード用ファイルの行"),
                        uploadTable,
                        tableActionRow,
                        statusLabel,
                        mainSplit);
        VBox.setVgrow(mainSplit, Priority.ALWAYS);

        refreshInteractionStateRef.get().run();
        if (deferInitialLoad) {
            statusLabel.setText("マスタデータを読み込んでいます…");
            startDeferredInitialization(
                    uiEnv,
                    statusLabel,
                    referenceHeaders,
                    editorModelRef,
                    uploadRows,
                    kouteiNaiyoLookupRef,
                    shuruiLookupRef,
                    keiriBunruiLookupRef,
                    yotoLookupRef,
                    bunrui4LookupRef,
                    zaikoBunruiLookupRef,
                    planMachineLookupRef,
                    reloadAllMasterLookups,
                    rebuildForm,
                    loadUploadFile,
                    log);
        } else {
            loadReferenceHeaders.run();
            loadUploadFile.run();
        }
        return root;
    }

    /**
     * 参照マスタ見出し・キャッシュ・アップロード用ファイルをバックグラウンドで読み、
     * フォーム構築（{@code rebuildForm}）だけ UI スレッドで行う。
     */
    private static void startDeferredInitialization(
            Supplier<Map<String, String>> uiEnv,
            Label statusLabel,
            List<String> referenceHeaders,
            AtomicReference<PostProcessingProductMasterEditorModel> editorModelRef,
            ObservableList<Map<String, String>> uploadRows,
            AtomicReference<PostProcessingKouteiNaiyoMasterLookup.Snapshot> kouteiNaiyoLookupRef,
            AtomicReference<PostProcessingShuruiMasterLookup.Snapshot> shuruiLookupRef,
            AtomicReference<PostProcessingKeiriBunruiMasterLookup.Snapshot> keiriBunruiLookupRef,
            AtomicReference<PostProcessingYotoMasterLookup.Snapshot> yotoLookupRef,
            AtomicReference<PostProcessingShohinBunrui4MasterLookup.Snapshot> bunrui4LookupRef,
            AtomicReference<PostProcessingZaikoBunruiMasterLookup.Snapshot> zaikoBunruiLookupRef,
            AtomicReference<PostProcessingPlanMachineLookup.Snapshot> planMachineLookupRef,
            Runnable reloadAllMasterLookups,
            Runnable rebuildForm,
            Runnable loadUploadFile,
            Consumer<String> log) {
        Thread worker =
                new Thread(
                        () -> {
                            List<String> headers = new ArrayList<>();
                            PlanInputTabularIo.TabularSheet uploadSheet = null;
                            PostProcessingKouteiNaiyoMasterLookup.Snapshot kouteiNaiyo =
                                    PostProcessingKouteiNaiyoMasterLookup.Snapshot.empty();
                            PostProcessingShuruiMasterLookup.Snapshot shurui =
                                    PostProcessingShuruiMasterLookup.Snapshot.empty();
                            PostProcessingKeiriBunruiMasterLookup.Snapshot keiriBunrui =
                                    PostProcessingKeiriBunruiMasterLookup.Snapshot.empty();
                            PostProcessingYotoMasterLookup.Snapshot yoto =
                                    PostProcessingYotoMasterLookup.Snapshot.empty();
                            PostProcessingShohinBunrui4MasterLookup.Snapshot bunrui4 =
                                    PostProcessingShohinBunrui4MasterLookup.Snapshot.empty();
                            PostProcessingZaikoBunruiMasterLookup.Snapshot zaikoBunrui =
                                    PostProcessingZaikoBunruiMasterLookup.Snapshot.empty();
                            PostProcessingPlanMachineLookup.Snapshot planMachine =
                                    PostProcessingPlanMachineLookup.Snapshot.empty();
                            boolean refOk = false;
                            boolean uploadOk = false;
                            try {
                                kouteiNaiyo =
                                        PostProcessingKouteiNaiyoMasterLookup.snapshot(
                                                uiEnv.get());
                                shurui =
                                        PostProcessingShuruiMasterLookup.snapshot(uiEnv.get());
                                keiriBunrui =
                                        PostProcessingKeiriBunruiMasterLookup.snapshot(
                                                uiEnv.get());
                                yoto = PostProcessingYotoMasterLookup.snapshot(uiEnv.get());
                                bunrui4 =
                                        PostProcessingShohinBunrui4MasterLookup.snapshot(
                                                uiEnv.get());
                                zaikoBunrui =
                                        PostProcessingZaikoBunruiMasterLookup.snapshot(
                                                uiEnv.get());
                                planMachine =
                                        PostProcessingPlanMachineLookup.snapshot(uiEnv.get());
                                Path ref =
                                        PostProcessingProductMasterIo.resolveReferencePath(
                                                uiEnv.get());
                                if (Files.isRegularFile(ref)) {
                                    headers =
                                            PostProcessingProductMasterIo.readHeaders(ref);
                                    PostProcessingProductMasterReferenceCache.snapshot(ref);
                                    refOk = true;
                                }
                                Path up =
                                        PostProcessingProductMasterIo.resolveUploadPath(
                                                uiEnv.get());
                                if (Files.isRegularFile(up)) {
                                    uploadSheet =
                                            PostProcessingProductMasterIo.readUploadWorkbook(
                                                    up);
                                    uploadOk = true;
                                }
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () ->
                                                statusLabel.setText(
                                                        "マスタ読込失敗: " + ex.getMessage()));
                                log.accept("[postproc-master] deferred init: " + ex.getMessage());
                            }
                            boolean finalRefOk = refOk;
                            boolean finalUploadOk = uploadOk;
                            List<String> loadedHeaders = List.copyOf(headers);
                            PlanInputTabularIo.TabularSheet loadedUpload = uploadSheet;
                            PostProcessingKouteiNaiyoMasterLookup.Snapshot loadedKouteiNaiyo =
                                    kouteiNaiyo;
                            PostProcessingShuruiMasterLookup.Snapshot loadedShurui = shurui;
                            PostProcessingKeiriBunruiMasterLookup.Snapshot loadedKeiriBunrui =
                                    keiriBunrui;
                            PostProcessingYotoMasterLookup.Snapshot loadedYoto = yoto;
                            PostProcessingShohinBunrui4MasterLookup.Snapshot loadedBunrui4 =
                                    bunrui4;
                            PostProcessingZaikoBunruiMasterLookup.Snapshot loadedZaikoBunrui =
                                    zaikoBunrui;
                            PostProcessingPlanMachineLookup.Snapshot loadedPlanMachine =
                                    planMachine;
                            Platform.runLater(
                                    () -> {
                                        kouteiNaiyoLookupRef.set(loadedKouteiNaiyo);
                                        shuruiLookupRef.set(loadedShurui);
                                        keiriBunruiLookupRef.set(loadedKeiriBunrui);
                                        yotoLookupRef.set(loadedYoto);
                                        bunrui4LookupRef.set(loadedBunrui4);
                                        zaikoBunruiLookupRef.set(loadedZaikoBunrui);
                                        planMachineLookupRef.set(loadedPlanMachine);
                                        if (finalRefOk && !loadedHeaders.isEmpty()) {
                                            referenceHeaders.clear();
                                            referenceHeaders.addAll(loadedHeaders);
                                            editorModelRef.set(
                                                    new PostProcessingProductMasterEditorModel(
                                                            referenceHeaders));
                                            rebuildForm.run();
                                        }
                                        if (finalUploadOk && loadedUpload != null) {
                                            try {
                                                applyUploadSheet(
                                                        loadedUpload,
                                                        referenceHeaders,
                                                        uploadRows,
                                                        statusLabel);
                                            } catch (IllegalArgumentException ex) {
                                                statusLabel.setText(ex.getMessage());
                                            }
                                        } else if (finalRefOk) {
                                            statusLabel.setText(
                                                    "参照マスタ準備完了。"
                                                            + masterLookupSuffix(
                                                                    loadedKouteiNaiyo,
                                                                    loadedShurui,
                                                                    loadedKeiriBunrui,
                                                                    loadedYoto,
                                                                    loadedBunrui4,
                                                                    loadedZaikoBunrui,
                                                                    loadedPlanMachine)
                                                            + "アップロード用ファイルがありません（新規作成可）");
                                        } else {
                                            statusLabel.setText(
                                                    "参照マスタを確認してください。");
                                        }
                                    });
                        },
                        "postproc-master-deferred-init");
        worker.setDaemon(true);
        worker.start();
    }

    private static void applyUploadSheet(
            PlanInputTabularIo.TabularSheet sheet,
            List<String> referenceHeaders,
            ObservableList<Map<String, String>> uploadRows,
            Label statusLabel)
            throws IllegalArgumentException {
        if (sheet == null || sheet.headers() == null) {
            uploadRows.clear();
            statusLabel.setText("アップロード用ファイルがありません（新規作成可）");
            return;
        }
        PostProcessingProductMasterColumnGroups.validateHeadersMatch(
                referenceHeaders, sheet.headers());
        uploadRows.clear();
        for (List<String> row : sheet.rows()) {
            uploadRows.add(
                    new LinkedHashMap<>(
                            PostProcessingProductMasterIo.rowToMap(sheet.headers(), row)));
        }
        statusLabel.setText("アップロード用 " + uploadRows.size() + " 行を読み込みました。");
    }

    private static String masterLookupStatusMessage(
            int headerCount,
            PostProcessingKouteiNaiyoMasterLookup.Snapshot kouteiNaiyo,
            PostProcessingShuruiMasterLookup.Snapshot shurui,
            PostProcessingKeiriBunruiMasterLookup.Snapshot keiriBunrui,
            PostProcessingYotoMasterLookup.Snapshot yoto,
            PostProcessingShohinBunrui4MasterLookup.Snapshot bunrui4,
            PostProcessingZaikoBunruiMasterLookup.Snapshot zaikoBunrui,
            PostProcessingPlanMachineLookup.Snapshot planMachine) {
        return "参照マスタ見出し "
                + headerCount
                + " 列。"
                + masterLookupSuffix(
                        kouteiNaiyo,
                        shurui,
                        keiriBunrui,
                        yoto,
                        bunrui4,
                        zaikoBunrui,
                        planMachine);
    }

    private static String masterLookupSuffix(
            PostProcessingKouteiNaiyoMasterLookup.Snapshot kouteiNaiyo,
            PostProcessingShuruiMasterLookup.Snapshot shurui,
            PostProcessingKeiriBunruiMasterLookup.Snapshot keiriBunrui,
            PostProcessingYotoMasterLookup.Snapshot yoto,
            PostProcessingShohinBunrui4MasterLookup.Snapshot bunrui4,
            PostProcessingZaikoBunruiMasterLookup.Snapshot zaikoBunrui,
            PostProcessingPlanMachineLookup.Snapshot planMachine) {
        StringBuilder sb = new StringBuilder();
        if (kouteiNaiyo != null && kouteiNaiyo.loaded()) {
            sb.append(" 工程")
                    .append(kouteiNaiyo.kouteiCodeToName().size())
                    .append("件・加工内容")
                    .append(kouteiNaiyo.naiyoCodeToEntry().size())
                    .append("件。");
        }
        if (shurui != null && shurui.loaded()) {
            sb.append(" 種類").append(shurui.codeToName().size()).append("件。");
        }
        if (keiriBunrui != null && keiriBunrui.loaded()) {
            sb.append(" 経理分類").append(keiriBunrui.codeToName().size()).append("件。");
        }
        if (yoto != null && yoto.loaded()) {
            sb.append(" 用途").append(yoto.codeToName().size()).append("件。");
        }
        if (bunrui4 != null && bunrui4.loaded()) {
            sb.append(" 商品分類4").append(bunrui4.codeToName().size()).append("件。");
        }
        if (zaikoBunrui != null && zaikoBunrui.loaded()) {
            sb.append(" 在庫分類").append(zaikoBunrui.codeToName().size()).append("件。");
        }
        if (planMachine != null && planMachine.loaded()) {
            sb.append(" 機械")
                    .append(planMachine.machineCodeToName().size())
                    .append("件（加工計画）。");
        } else if (planMachine != null
                && !planMachine.hasCodeColumn()
                && !planMachine.hasNameColumn()) {
            sb.append(" 加工計画に機械/機械名列なし。");
        } else if (planMachine != null && !planMachine.loaded()) {
            sb.append(" 加工計画の機械一覧未構築。");
        }
        if ((kouteiNaiyo == null || !kouteiNaiyo.loaded())
                && (shurui == null || !shurui.loaded())
                && (keiriBunrui == null || !keiriBunrui.loaded())
                && (yoto == null || !yoto.loaded())
                && (bunrui4 == null || !bunrui4.loaded())
                && (zaikoBunrui == null || !zaikoBunrui.loaded())
                && (planMachine == null || !planMachine.loaded())) {
            if (sb.isEmpty()) {
                sb.append(" 連携マスタ未読込。");
            }
        }
        return sb.toString();
    }

    private record SearchResultColumn(
            String title, double prefWidth, java.util.function.Function<
                            PostProcessingProductMasterIo.SearchHit, String>
                    extractor) {}

    private static TableView<PostProcessingProductMasterIo.SearchHit> buildSearchResultTable(
            boolean compactCardTitle) {
        TableView<PostProcessingProductMasterIo.SearchHit> table = new TableView<>();
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        double searchTableHeight = compactCardTitle ? 180 : 280;
        table.setPrefHeight(searchTableHeight);
        table.setMinHeight(120);
        table.setMaxHeight(searchTableHeight);
        table.setPlaceholder(new Label("「検索」で雛形一覧を表示（最大200件）"));
        table.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
        for (SearchResultColumn spec : SEARCH_RESULT_COLUMNS) {
            TableColumn<PostProcessingProductMasterIo.SearchHit, String> col =
                    new TableColumn<>(spec.title());
            col.setPrefWidth(spec.prefWidth());
            col.setMinWidth(Math.min(spec.prefWidth(), 48));
            col.setCellValueFactory(
                    cd ->
                            new javafx.beans.property.SimpleStringProperty(
                                    spec.extractor().apply(cd.getValue())));
            table.getColumns().add(col);
        }
        return table;
    }

    private static String shohinCode(PostProcessingProductMasterIo.SearchHit h) {
        return h != null ? nullToEmpty(h.shohinCode()) : "";
    }

    private static String shohinName1(PostProcessingProductMasterIo.SearchHit h) {
        return h != null ? nullToEmpty(h.shohinName1()) : "";
    }

    private static String foamPartNo(PostProcessingProductMasterIo.SearchHit h) {
        return h != null ? nullToEmpty(h.foamPartNo()) : "";
    }

    private static String foamName(PostProcessingProductMasterIo.SearchHit h) {
        return h != null ? nullToEmpty(h.foamName()) : "";
    }

    private static String rowColumn(PostProcessingProductMasterIo.SearchHit h, String key) {
        if (h == null || h.rowByColumn() == null) {
            return "";
        }
        return nullToEmpty(h.rowByColumn().get(key));
    }

    private static String nullToEmpty(String s) {
        return s != null ? s.trim() : "";
    }

    private static HBox gridFilterRow(String label, TextField field) {
        Label lbl = new Label(label);
        lbl.setMinWidth(72);
        HBox.setHgrow(field, Priority.ALWAYS);
        return new HBox(6, lbl, field);
    }

    private static HBox toolbarGroup(String caption, Button... buttons) {
        Label lbl = new Label(caption);
        lbl.getStyleClass().add("postproc-master-toolbar-caption");
        HBox group = new HBox(6, lbl);
        group.setAlignment(Pos.CENTER_LEFT);
        group.getChildren().addAll(buttons);
        return group;
    }

    private static void showError(String title, String message) {
        Alert a = new Alert(Alert.AlertType.ERROR);
        a.setTitle(title);
        a.setHeaderText(null);
        a.setContentText(message);
        a.showAndWait();
    }
}
