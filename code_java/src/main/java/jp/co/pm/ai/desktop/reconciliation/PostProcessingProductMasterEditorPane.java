package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.SelectionMode;
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

import jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo;

/**
 * 依頼書入力タブ内「後加工商品マスタ編集」カードの UI。
 */
public final class PostProcessingProductMasterEditorPane {

    public record Context(
            Supplier<Map<String, String>> uiEnv,
            Supplier<ReconciliationApp.ProductRow> firstProductRow,
            Supplier<String> formKakoKbnLabel,
            Runnable runIntegratedMaster,
            Consumer<String> log) {}

    private static final List<String> UPLOAD_TABLE_COLUMNS =
            List.of("商品コード", "商品名1", "発泡体品番", "発泡体幅", "発泡体長さ");

    /** 参照マスタ検索結果テーブルの列（依頼書候補表示に近い並び）。 */
    private static final List<SearchResultColumn> SEARCH_RESULT_COLUMNS =
            List.of(
                    new SearchResultColumn("商品コード", 128, PostProcessingProductMasterEditorPane::shohinCode),
                    new SearchResultColumn("商品名1", 200, PostProcessingProductMasterEditorPane::shohinName1),
                    new SearchResultColumn("品番", 72, PostProcessingProductMasterEditorPane::foamPartNo),
                    new SearchResultColumn("品名", 56, PostProcessingProductMasterEditorPane::foamName),
                    new SearchResultColumn("タイプ", 64, h -> rowColumn(h, "発泡体タイプ")),
                    new SearchResultColumn("幅", 64, h -> rowColumn(h, "発泡体幅")),
                    new SearchResultColumn("長さ", 64, h -> rowColumn(h, "発泡体長さ")),
                    new SearchResultColumn("色", 48, h -> rowColumn(h, "発泡体色")));

    private PostProcessingProductMasterEditorPane() {}

    /** 依頼書入力の専用タブ用（横幅いっぱい）。 */
    public static VBox buildTabContent(Window owner, Context ctx) {
        VBox content = buildContent(owner, ctx, false);
        content.getStyleClass().add("form-tab-container");
        content.setFillWidth(true);
        content.setMaxWidth(Double.MAX_VALUE);
        VBox.setVgrow(content, Priority.ALWAYS);
        return content;
    }

    /** 【設定】タブ内カード用（幅上限あり）。 */
    public static VBox buildCard(Window owner, Context ctx, double maxCardWidth) {
        VBox card = buildContent(owner, ctx, true);
        card.getStyleClass().add("settings-card");
        card.setMaxWidth(maxCardWidth);
        card.setPrefWidth(maxCardWidth);
        return card;
    }

    private static VBox buildContent(Window owner, Context ctx, boolean compactCardTitle) {
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
        TabPane formTabs = new TabPane();
        formTabs.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);
        formTabs.setMinHeight(compactCardTitle ? 280 : 360);
        VBox.setVgrow(formTabs, Priority.ALWAYS);

        List<String> referenceHeaders = new ArrayList<>();
        AtomicReference<PostProcessingProductMasterEditorModel> editorModelRef =
                new AtomicReference<>(new PostProcessingProductMasterEditorModel(List.of()));
        Map<String, String> templateRow = new LinkedHashMap<>();

        ObservableList<Map<String, String>> uploadRows = FXCollections.observableArrayList();
        TableView<Map<String, String>> uploadTable = new TableView<>(uploadRows);
        uploadTable.setPrefHeight(compactCardTitle ? 120 : 160);
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

        Runnable rebuildForm =
                () -> {
                    fieldByColumn.clear();
                    formTabs.getTabs().clear();
                    for (PostProcessingProductMasterColumnGroups.TabGroup group :
                            PostProcessingProductMasterColumnGroups.tabGroups()) {
                        GridPane grid = new GridPane();
                        grid.setHgap(8);
                        grid.setVgap(6);
                        grid.setPadding(new Insets(8));
                        ColumnConstraints labelCol = new ColumnConstraints();
                        labelCol.setMinWidth(140);
                        labelCol.setPrefWidth(160);
                        ColumnConstraints fieldCol = new ColumnConstraints();
                        fieldCol.setHgrow(Priority.ALWAYS);
                        grid.getColumnConstraints().addAll(labelCol, fieldCol);
                        int row = 0;
                        for (String colName : group.columnNames()) {
                            if (!editorModelRef.get().headers().contains(colName)) {
                                continue;
                            }
                            Label lbl = new Label(colName + ":");
                            lbl.setStyle("-fx-font-size: 11px;");
                            TextField tf = new TextField(editorModelRef.get().get(colName));
                            tf.setStyle("-fx-font-size: 11px;");
                            tf.textProperty()
                                    .addListener(
                                            (obs, o, n) ->
                                                    editorModelRef
                                                            .get()
                                                            .set(colName, n != null ? n : ""));
                            fieldByColumn.put(colName, tf);
                            grid.add(lbl, 0, row);
                            grid.add(tf, 1, row++);
                        }
                        ScrollPane sp = new ScrollPane(grid);
                        sp.setFitToWidth(true);
                        Tab tab = new Tab(group.tabTitle(), sp);
                        formTabs.getTabs().add(tab);
                    }
                };

        Runnable syncFormFromModel =
                () -> {
                    for (Map.Entry<String, TextField> e : fieldByColumn.entrySet()) {
                        e.getValue().setText(editorModelRef.get().get(e.getKey()));
                    }
                };

        Runnable loadReferenceHeaders =
                () -> {
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
                        rebuildForm.run();
                        statusLabel.setText(
                                "参照マスタ見出し "
                                        + referenceHeaders.size()
                                        + " 列を読み込みました。");
                    } catch (Exception ex) {
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
                        var sheet = PostProcessingProductMasterIo.readUploadWorkbook(up);
                        PostProcessingProductMasterColumnGroups.validateHeadersMatch(
                                referenceHeaders, sheet.headers());
                        uploadRows.clear();
                        for (List<String> row : sheet.rows()) {
                            uploadRows.add(
                                    new LinkedHashMap<>(
                                            PostProcessingProductMasterIo.rowToMap(
                                                    sheet.headers(), row)));
                        }
                        statusLabel.setText(
                                "アップロード用 "
                                        + uploadRows.size()
                                        + " 行を読み込みました。");
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
                                            List<PostProcessingProductMasterIo.SearchHit> hits =
                                                    PostProcessingProductMasterIo.searchReference(
                                                            ref, filter, 200);
                                            Platform.runLater(
                                                    () -> {
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

        Runnable applySelectedSearchHit =
                () -> {
                    PostProcessingProductMasterIo.SearchHit hit =
                            searchResults.getSelectionModel().getSelectedItem();
                    if (hit != null) {
                        templateRow.clear();
                        templateRow.putAll(hit.rowByColumn());
                    }
                };
        searchResults
                .getSelectionModel()
                .selectedItemProperty()
                .addListener((obs, old, hit) -> applySelectedSearchHit.run());
        searchResults.setOnMouseClicked(
                e -> {
                    if (e.getClickCount() == 2) {
                        applySelectedSearchHit.run();
                        editorModelRef.get().applyTemplateRow(templateRow);
                        syncFormFromModel.run();
                        statusLabel.setText("雛形をフォームに反映しました（商品コードは手修正）。");
                    }
                });

        Button btnTemplateToForm = new Button("雛形をフォームへ");
        btnTemplateToForm.getStyleClass().add("btn-reload");
        btnTemplateToForm.setOnAction(
                e -> {
                    if (templateRow.isEmpty()) {
                        showError("雛形未選択", "検索結果から雛形行を選択してください。");
                        return;
                    }
                    editorModelRef.get().applyTemplateRow(templateRow);
                    syncFormFromModel.run();
                    statusLabel.setText("雛形をフォームに反映しました（商品コードは手修正）。");
                });

        Button btnNewFromTemplate = new Button("雛形から新規行");
        btnNewFromTemplate.getStyleClass().add("btn-transfer");
        btnNewFromTemplate.setOnAction(
                e -> {
                    if (templateRow.isEmpty()) {
                        showError("雛形未選択", "検索結果から雛形行を選択してください。");
                        return;
                    }
                    editorModelRef.get().applyTemplateRow(templateRow);
                    syncFormFromModel.run();
                    statusLabel.setText("雛形を適用しました。編集後「行追加」で一覧へ。");
                });

        Button btnTransfer = new Button("依頼書製品行から転記");
        btnTransfer.getStyleClass().add("btn-transfer");
        btnTransfer.setOnAction(
                e -> {
                    if (templateRow.isEmpty()) {
                        showError(
                                "雛形未選択",
                                "転記前に検索結果から雛形行を選び「雛形をフォームへ」を実行するか、"
                                        + "雛形から新規行で雛形を適用してください。");
                        return;
                    }
                    ReconciliationApp.ProductRow row = ctx.firstProductRow().get();
                    if (row == null) {
                        showError("製品行なし", "依頼書フォームに製品行がありません。");
                        return;
                    }
                    editorModelRef.get().applyTemplateRow(templateRow);
                    editorModelRef
                            .get()
                            .applyRequestFormProductRow(row, ctx.formKakoKbnLabel().get());
                    syncFormFromModel.run();
                    statusLabel.setText("依頼書製品行を転記しました（商品コードは雛形ベース）。");
                });

        Button btnAddRow = new Button("行追加");
        btnAddRow.getStyleClass().add("btn-settings-add");
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

        Button btnRemoveRow = new Button("行削除");
        btnRemoveRow.getStyleClass().add("btn-settings-del");
        btnRemoveRow.setOnAction(
                e -> {
                    Map<String, String> sel = uploadTable.getSelectionModel().getSelectedItem();
                    if (sel != null) {
                        uploadRows.remove(sel);
                    }
                });

        Button btnDupRow = new Button("行複製");
        btnDupRow.getStyleClass().add("btn-reload");
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
                            if (row != null) {
                                editorModelRef.get().loadFromRow(row);
                                syncFormFromModel.run();
                            }
                        });

        Button btnNewUpload = new Button("アップロード用を新規作成");
        btnNewUpload.getStyleClass().add("btn-reload");
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

        Button btnBrowseUpload = new Button("参照…");
        btnBrowseUpload.getStyleClass().add("btn-reload");
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

        Button btnSave = new Button("保存");
        btnSave.getStyleClass().add("btn-transfer");
        btnSave.setOnAction(e -> saveUpload.run());

        Button btnReloadUpload = new Button("再読込");
        btnReloadUpload.getStyleClass().add("btn-reload");
        btnReloadUpload.setOnAction(e -> loadUploadFile.run());

        Button btnIntegrated = new Button("統合マスタ再生成");
        btnIntegrated.getStyleClass().add("btn-transfer");
        btnIntegrated.setTooltip(
                new Tooltip("create_integrated_master.py を実行し、依頼書候補を更新します。"));
        btnIntegrated.setOnAction(e -> ctx.runIntegratedMaster().run());

        refreshPathsFromEnv.run();
        loadReferenceHeaders.run();
        loadUploadFile.run();

        Label note =
                new Label(
                        "※ Aladdin の単価タブ（上代・販売・購買単価等）は本 xlsx に含まれません。"
                                + " 本番マスタへの上書きは行わず、アップロード用ファイルを編集してください。");
        note.setWrapText(true);
        note.setStyle("-fx-font-size: 10px;");

        VBox filterBox = new VBox(6);
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
                        new HBox(6, btnTemplateToForm, btnNewFromTemplate));

        HBox pathRow1 = new HBox(6, new Label("参照:"), refPathField);
        HBox.setHgrow(refPathField, Priority.ALWAYS);
        HBox pathRow2 =
                new HBox(
                        6,
                        new Label("アップロード用:"),
                        uploadPathField,
                        btnBrowseUpload,
                        btnNewUpload,
                        btnReloadUpload,
                        btnSave);
        HBox actionRow =
                new HBox(
                        8,
                        btnTransfer,
                        btnAddRow,
                        btnRemoveRow,
                        btnDupRow,
                        btnIntegrated);
        actionRow.setAlignment(Pos.CENTER_LEFT);

        VBox topPaths = new VBox(4, pathRow1, pathRow2);
        HBox mainRow = new HBox(12, filterBox, formTabs);
        HBox.setHgrow(formTabs, Priority.ALWAYS);
        HBox.setHgrow(mainRow, Priority.ALWAYS);
        filterBox.setMinWidth(compactCardTitle ? 300 : 440);
        filterBox.setPrefWidth(compactCardTitle ? 300 : 480);
        filterBox.setMaxWidth(compactCardTitle ? 340 : 560);
        filterBox.setMaxHeight(Double.MAX_VALUE);
        VBox.setVgrow(searchResults, Priority.ALWAYS);

        VBox root = new VBox(10);
        root.setMaxWidth(Double.MAX_VALUE);
        Label title = new Label("後加工商品マスタ編集");
        title.getStyleClass().add(compactCardTitle ? "settings-card-title" : "paper-main-title");
        root.getChildren()
                .addAll(
                        title,
                        note,
                        topPaths,
                        statusLabel,
                        mainRow,
                        new Label("アップロード用ファイルの行"),
                        uploadTable,
                        actionRow);
        VBox.setVgrow(mainRow, Priority.ALWAYS);
        return root;
    }

    private record SearchResultColumn(
            String title, double prefWidth, java.util.function.Function<
                            PostProcessingProductMasterIo.SearchHit, String>
                    extractor) {}

    private static TableView<PostProcessingProductMasterIo.SearchHit> buildSearchResultTable(
            boolean compactCardTitle) {
        TableView<PostProcessingProductMasterIo.SearchHit> table = new TableView<>();
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        table.setPrefHeight(compactCardTitle ? 180 : 340);
        table.setMinHeight(140);
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

    private static void showError(String title, String message) {
        Alert a = new Alert(Alert.AlertType.ERROR);
        a.setTitle(title);
        a.setHeaderText(null);
        a.setContentText(message);
        a.showAndWait();
    }
}
