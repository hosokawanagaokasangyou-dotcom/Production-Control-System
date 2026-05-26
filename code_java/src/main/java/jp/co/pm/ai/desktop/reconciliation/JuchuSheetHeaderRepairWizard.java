package jp.co.pm.ai.desktop.reconciliation;

import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.beans.property.ObjectProperty;
import javafx.beans.property.SimpleObjectProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextArea;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.Window;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.function.Supplier;

/**
 * 受注ﾌｧｲﾙの列定義ウィザード（既知列・未知列）。
 */
public final class JuchuSheetHeaderRepairWizard {

    public enum Result {
        CANCEL,
        CONTINUE,
        FIXED
    }

    public enum FixAction {
        REDEFINE("期待定義をExcel見出しで再定義"),
        ALIAS("実際の見出しを別名として許容"),
        EXCLUDE("転記/吸出しから除外"),
        SKIP("対応しない");

        private final String label;

        FixAction(String label) {
            this.label = label;
        }

        @Override
        public String toString() {
            return label;
        }
    }

    public enum UnknownAction {
        SKIP("対応しない"),
        IGNORE("転記・検証対象外（無視）"),
        ALIAS_TO_KNOWN("既知列の別名として登録");

        private final String label;

        UnknownAction(String label) {
            this.label = label;
        }

        @Override
        public String toString() {
            return label;
        }
    }

    private enum DialogMode {
        MANAGE,
        TRANSFER_PROMPT
    }

    static final class KnownRow {
        final JuchuHeaderMismatch mismatch;
        final ObjectProperty<FixAction> action = new SimpleObjectProperty<>();
        /** ComboBox 表示用（例: {@code BV列: 商品(原反)}）。 */
        final StringProperty selectedPickLabel = new SimpleStringProperty();
        /** レジストリ保存用の見出し文字列。 */
        final StringProperty selectedExcelHeader = new SimpleStringProperty();

        KnownRow(
                JuchuHeaderMismatch mismatch,
                FixAction defaultAction,
                String defaultPickLabel,
                String defaultHeaderText) {
            this.mismatch = mismatch;
            this.action.set(defaultAction);
            this.selectedPickLabel.set(defaultPickLabel != null ? defaultPickLabel : "");
            this.selectedExcelHeader.set(defaultHeaderText != null ? defaultHeaderText : "");
        }

        void applyPickSelection(
                String comboValue, List<JuchuSheetColumnLayout.ExcelHeaderPick> picks) {
            String label = comboValue != null ? comboValue.strip() : "";
            selectedPickLabel.set(label);
            JuchuSheetColumnLayout.ExcelHeaderPick pick = resolvePick(label, picks);
            if (pick != null) {
                selectedExcelHeader.set(pick.headerText());
            } else if (!label.isEmpty()) {
                selectedExcelHeader.set(label);
            } else {
                selectedExcelHeader.set("");
            }
        }

        boolean matching(JuchuHeaderAliasRegistry registry, String path) {
            return JuchuSheetColumnLayout.headerMatches(
                    mismatch.column(), mismatch.actualHeader(), registry, path);
        }

        String formItem() {
            return mismatch.formItemDescription();
        }

        String columnLetter() {
            return mismatch.columnLetter();
        }

        String expected() {
            return mismatch.expectedHeader();
        }

        String actual() {
            return mismatch.actualEmpty() ? "（空）" : mismatch.actualHeader();
        }

        String status(JuchuHeaderAliasRegistry registry, String path) {
            if (registry != null && registry.isExcludedFromTransfer(path, mismatch.column())) {
                return "転記除外";
            }
            return matching(registry, path) ? "一致" : "不一致";
        }

        FixAction getAction() {
            return action.get();
        }

        void setAction(FixAction value) {
            action.set(value);
        }

        String getSelectedExcelHeader() {
            return selectedExcelHeader.get();
        }

        void setSelectedExcelHeader(String value) {
            selectedExcelHeader.set(value);
        }

        String getSelectedPickLabel() {
            return selectedPickLabel.get();
        }
    }

    static final class UnknownRow {
        final JuchuUnknownExcelColumn column;
        final ObjectProperty<UnknownAction> action = new SimpleObjectProperty<>(UnknownAction.SKIP);
        final ObjectProperty<JuchuSheetColumnLayout.Col> aliasTarget =
                new SimpleObjectProperty<>();

        UnknownRow(JuchuUnknownExcelColumn column) {
            this.column = column;
            if (column.ignored()) {
                this.action.set(UnknownAction.IGNORE);
            }
        }

        String columnLetter() {
            return column.columnLetter();
        }

        String headerText() {
            return column.headerText();
        }

        String status() {
            return column.ignored() ? "無視済み" : "未設定";
        }

        UnknownAction getAction() {
            return action.get();
        }

        void setAction(UnknownAction value) {
            action.set(value);
        }

        JuchuSheetColumnLayout.Col getAliasTarget() {
            return aliasTarget.get();
        }

        void setAliasTarget(JuchuSheetColumnLayout.Col value) {
            aliasTarget.set(value);
        }
    }

    private JuchuSheetHeaderRepairWizard() {}

    /** 設定タブ／手動起動。 */
    public static void showManage(
            Window owner, File juchuFile, JuchuHeaderAliasRegistry registry) {
        showDialog(owner, juchuFile, registry, DialogMode.MANAGE, null);
    }

    /** 転記前の警告フロー。 */
    public static Result showTransferPrompt(
            Window owner,
            File juchuFile,
            List<JuchuHeaderMismatch> mismatches,
            JuchuHeaderAliasRegistry registry) {
        if (mismatches == null || mismatches.isEmpty()) {
            return Result.CONTINUE;
        }
        return showDialog(owner, juchuFile, registry, DialogMode.TRANSFER_PROMPT, mismatches);
    }

    /** @deprecated {@link #showTransferPrompt} を使用 */
    @Deprecated
    public static Result showAndWait(
            Window owner,
            File juchuFile,
            List<JuchuHeaderMismatch> mismatches,
            JuchuHeaderAliasRegistry registry) {
        return showTransferPrompt(owner, juchuFile, mismatches, registry);
    }

    private static Result showDialog(
            Window owner,
            File juchuFile,
            JuchuHeaderAliasRegistry registry,
            DialogMode mode,
            List<JuchuHeaderMismatch> transferMismatches) {
        Objects.requireNonNull(juchuFile, "juchuFile");
        Objects.requireNonNull(registry, "registry");

        SheetContext sheetContext;
        try {
            sheetContext = loadSheetContext(juchuFile, registry);
        } catch (Exception ex) {
            showError(
                    owner,
                    "受注ファイルの見出し一覧を読み込めませんでした: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex));
            return Result.CANCEL;
        }

        String pathKey = juchuFile.getAbsolutePath();
        final List<JuchuSheetColumnLayout.ExcelHeaderPick> excelHeaderPicks =
                new ArrayList<>(sheetContext.headerPicks());
        final Supplier<List<JuchuSheetColumnLayout.ExcelHeaderPick>> headerPicksSupplier =
                () -> excelHeaderPicks;
        final javafx.collections.ObservableList<String> pickLabelItems =
                FXCollections.observableArrayList(headerPickLabels(excelHeaderPicks));
        List<KnownRow> allKnownRows = buildKnownRows(sheetContext, registry, pathKey);
        List<UnknownRow> unknownRows = buildUnknownRows(sheetContext, registry, pathKey);

        Spinner<Integer> headerRowSpinner =
                new Spinner<>(
                        new SpinnerValueFactory.IntegerSpinnerValueFactory(
                                1,
                                200,
                                registry.headerRowOneBasedFor(pathKey)));
        headerRowSpinner.setEditable(true);
        headerRowSpinner.setPrefWidth(90);
        Label headerRowLabel = new Label("見出し行:");
        Label headerRowHint = new Label("行目（1始まり・既定3）");
        Button btnApplyHeaderRow = new Button("見出し行を反映");
        HBox headerRowRow =
                new HBox(
                        8,
                        headerRowLabel,
                        headerRowSpinner,
                        headerRowHint,
                        btnApplyHeaderRow);
        headerRowRow.setAlignment(Pos.CENTER_LEFT);

        boolean mismatchesOnlyDefault = mode == DialogMode.TRANSFER_PROMPT;
        CheckBox chkMismatchesOnly = new CheckBox("不一致の既知列のみ表示");
        chkMismatchesOnly.setSelected(mismatchesOnlyDefault);
        chkMismatchesOnly.setVisible(mode == DialogMode.MANAGE);

        TableView<KnownRow> knownTable =
                createKnownTable(
                        allKnownRows,
                        headerPicksSupplier,
                        pickLabelItems,
                        registry,
                        pathKey);
        refreshKnownTableItems(knownTable, allKnownRows, registry, pathKey, mismatchesOnlyDefault);
        chkMismatchesOnly.setOnAction(
                e ->
                        refreshKnownTableItems(
                                knownTable,
                                allKnownRows,
                                registry,
                                pathKey,
                                chkMismatchesOnly.isSelected()));

        TableView<UnknownRow> unknownTable = createUnknownTable(unknownRows);

        Tab tabKnown = new Tab("既知の列（フォーム転記定義）");
        tabKnown.setClosable(false);
        VBox knownBox = new VBox(8, chkMismatchesOnly, knownTable);
        VBox.setVgrow(knownTable, Priority.ALWAYS);
        knownBox.setPadding(new Insets(8, 0, 0, 0));
        tabKnown.setContent(knownBox);

        Tab tabUnknown = new Tab("未知の列（定義外のExcel見出し）");
        tabUnknown.setClosable(false);
        Label unknownIntro =
                new Label(
                        "転記定義に無い列位置の見出しです。「無視」で一覧から除外、「既知列の別名」で"
                                + " 既知列の検証用別名に登録できます。");
        unknownIntro.setWrapText(true);
        unknownIntro.setStyle("-fx-text-fill: #555;");
        VBox unknownBox = new VBox(8, unknownIntro, unknownTable);
        VBox.setVgrow(unknownTable, Priority.ALWAYS);
        unknownBox.setPadding(new Insets(8, 0, 0, 0));
        tabUnknown.setContent(unknownBox);

        TabPane tabPane = new TabPane(tabKnown, tabUnknown);
        tabPane.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);
        if (mode == DialogMode.TRANSFER_PROMPT) {
            tabKnown.setText("既知の列 — 不一致 " + (transferMismatches != null ? transferMismatches.size() : 0) + " 件");
        }

        Runnable reloadSheetRows =
                () -> {
                    try {
                        SheetContext refreshed = loadSheetContext(juchuFile, registry);
                        excelHeaderPicks.clear();
                        excelHeaderPicks.addAll(refreshed.headerPicks());
                        pickLabelItems.setAll(headerPickLabels(excelHeaderPicks));
                        allKnownRows.clear();
                        allKnownRows.addAll(buildKnownRows(refreshed, registry, pathKey));
                        unknownRows.clear();
                        unknownRows.addAll(buildUnknownRows(refreshed, registry, pathKey));
                        refreshKnownTableItems(
                                knownTable,
                                allKnownRows,
                                registry,
                                pathKey,
                                chkMismatchesOnly.isSelected());
                        unknownTable.setItems(FXCollections.observableArrayList(unknownRows));
                        knownTable.refresh();
                        unknownTable.refresh();
                    } catch (Exception ex) {
                        showError(
                                owner,
                                "見出し行の反映に失敗しました: "
                                        + (ex.getMessage() != null ? ex.getMessage() : ex));
                    }
                };
        btnApplyHeaderRow.setOnAction(
                e -> {
                    Integer row = headerRowSpinner.getValue();
                    if (row != null) {
                        registry.setHeaderRowOneBasedFor(pathKey, row);
                    }
                    try {
                        registry.saveToDisk();
                    } catch (Exception ex) {
                        showError(
                                owner,
                                "見出し行の保存に失敗しました: "
                                        + (ex.getMessage() != null ? ex.getMessage() : ex));
                        return;
                    }
                    reloadSheetRows.run();
                });

        Stage stage = new Stage();
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle("受注シート列定義 — 修正ウィザード");

        String fileName = juchuFile.getName();
        Label intro =
                new Label(
                        mode == DialogMode.TRANSFER_PROMPT
                                ? "受注ファイル「"
                                        + fileName
                                        + "」の見出し定義に不一致があります。"
                                        + " 「既知の列」タブでフォーム項目ごとに修正し、「未知の列」タブで定義外見出しを設定できます。"
                                : "受注ファイル「"
                                        + fileName
                                        + "」の見出し行・列定義を確認・編集します。"
                                        + " タブで「既知の列（転記定義）」と「未知の列（定義外）」を切り替えてください。");
        intro.setWrapText(true);

        Label statusLabel = new Label("");
        statusLabel.setWrapText(true);

        final Result[] outcome = {mode == DialogMode.MANAGE ? Result.FIXED : Result.CANCEL};

        Button btnApply = new Button("適用して再検証");
        btnApply.setDefaultButton(true);
        btnApply.setOnAction(
                e -> {
                    try {
                        knownTable.edit(-1, null);
                        Integer headerRow = headerRowSpinner.getValue();
                        if (headerRow != null) {
                            registry.setHeaderRowOneBasedFor(pathKey, headerRow);
                        }
                        commitKnownRowPickSelections(allKnownRows, excelHeaderPicks);
                        applyAll(allKnownRows, unknownRows, registry, pathKey, excelHeaderPicks);
                        registry.saveToDisk();
                        List<JuchuHeaderMismatch> remaining = readMismatches(juchuFile, registry);
                        reloadSheetRows.run();
                        if (remaining.isEmpty()) {
                            statusLabel.setText("列定義を適用しました。不一致は解消されています。");
                            if (mode == DialogMode.TRANSFER_PROMPT) {
                                outcome[0] = Result.FIXED;
                                stage.close();
                            }
                        } else {
                            statusLabel.setText(
                                    "不一致が "
                                            + remaining.size()
                                            + " 件残っています。設定を見直して再度「適用」してください。");
                            tabKnown.setText("既知の列 — 不一致 " + remaining.size() + " 件");
                        }
                    } catch (Exception ex) {
                        showError(
                                stage,
                                ex.getMessage() != null ? ex.getMessage() : String.valueOf(ex));
                    }
                });

        Button btnContinue = new Button("未修正のまま続行");
        btnContinue.setManaged(mode == DialogMode.TRANSFER_PROMPT);
        btnContinue.setVisible(mode == DialogMode.TRANSFER_PROMPT);
        btnContinue.setOnAction(
                e -> {
                    outcome[0] = Result.CONTINUE;
                    stage.close();
                });

        Button btnClose = new Button(mode == DialogMode.MANAGE ? "閉じる" : "中止");
        btnClose.setCancelButton(true);
        btnClose.setOnAction(
                e -> {
                    if (mode == DialogMode.TRANSFER_PROMPT) {
                        outcome[0] = Result.CANCEL;
                    }
                    stage.close();
                });

        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox buttons =
                mode == DialogMode.MANAGE
                        ? new HBox(10, btnClose, spacer, btnApply)
                        : new HBox(10, btnClose, btnContinue, spacer, btnApply);
        buttons.setAlignment(Pos.CENTER_RIGHT);
        buttons.setPadding(new Insets(8, 0, 0, 0));

        VBox center = new VBox(10, intro, headerRowRow, tabPane, statusLabel);
        center.setPadding(new Insets(12));
        VBox.setVgrow(tabPane, Priority.ALWAYS);

        BorderPane root = new BorderPane();
        root.setCenter(center);
        root.setBottom(buttons);
        Scene scene = new Scene(root, 1120, 560);
        stage.setScene(scene);
        stage.showAndWait();
        return outcome[0];
    }

    private static void refreshKnownTableItems(
            TableView<KnownRow> table,
            List<KnownRow> allKnownRows,
            JuchuHeaderAliasRegistry registry,
            String pathKey,
            boolean mismatchesOnly) {
        List<KnownRow> visible = new ArrayList<>();
        for (KnownRow row : allKnownRows) {
            if (!mismatchesOnly || !row.matching(registry, pathKey)) {
                visible.add(row);
            }
        }
        table.setItems(FXCollections.observableArrayList(visible));
        table.refresh();
    }

    private static List<KnownRow> buildKnownRows(
            SheetContext ctx, JuchuHeaderAliasRegistry registry, String pathKey) {
        List<KnownRow> rows = new ArrayList<>();
        for (JuchuHeaderMismatch m :
                JuchuSheetColumnLayout.collectAllKnownColumns(ctx.headerRow(), registry, pathKey)) {
            FixAction defaultAction;
            if (registry.isExcludedFromTransfer(pathKey, m.column())) {
                defaultAction = FixAction.EXCLUDE;
            } else if (JuchuSheetColumnLayout.headerMatches(
                    m.column(), m.actualHeader(), registry, pathKey)) {
                defaultAction = FixAction.SKIP;
            } else {
                defaultAction = FixAction.REDEFINE;
            }
            String defaultPick =
                    defaultSelectedPickLabel(m, ctx.headerPicks(), registry, pathKey);
            rows.add(
                    new KnownRow(
                            m,
                            defaultAction,
                            defaultPick,
                            headerTextForPickLabel(defaultPick, ctx.headerPicks())));
        }
        return rows;
    }

    private static List<UnknownRow> buildUnknownRows(
            SheetContext ctx, JuchuHeaderAliasRegistry registry, String pathKey) {
        List<UnknownRow> rows = new ArrayList<>();
        for (JuchuUnknownExcelColumn col :
                JuchuSheetColumnLayout.collectUnknownExcelColumns(
                        ctx.headerRow(), registry, pathKey)) {
            rows.add(new UnknownRow(col));
        }
        return rows;
    }

    private static TableView<KnownRow> createKnownTable(
            List<KnownRow> rows,
            Supplier<List<JuchuSheetColumnLayout.ExcelHeaderPick>> headerPicksSupplier,
            javafx.collections.ObservableList<String> pickLabelItems,
            JuchuHeaderAliasRegistry registry,
            String pathKey) {
        TableView<KnownRow> table = new TableView<>();
        table.setEditable(true);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);
        table.setPrefHeight(280);

        TableColumn<KnownRow, String> colStatus = new TableColumn<>("状態");
        colStatus.setCellValueFactory(
                c ->
                        new SimpleStringProperty(
                                c.getValue().status(registry, pathKey)));

        TableColumn<KnownRow, String> colFormItem = new TableColumn<>("フォーム項目");
        colFormItem.setCellValueFactory(
                c -> new SimpleStringProperty(c.getValue().formItem()));

        TableColumn<KnownRow, String> colLetter = new TableColumn<>("列");
        colLetter.setCellValueFactory(
                c -> new SimpleStringProperty(c.getValue().columnLetter()));

        TableColumn<KnownRow, String> colExpected = new TableColumn<>("期待見出し");
        colExpected.setCellValueFactory(
                c -> new SimpleStringProperty(c.getValue().expected()));

        TableColumn<KnownRow, String> colActual = new TableColumn<>("Excel見出し");
        colActual.setCellValueFactory(
                c -> new SimpleStringProperty(c.getValue().actual()));

        TableColumn<KnownRow, String> colPick = new TableColumn<>("採用Excel見出し");
        colPick.setCellValueFactory(c -> c.getValue().selectedPickLabel);
        colPick.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final ComboBox<String> combo = new ComboBox<>();
                            private KnownRow boundRow;
                            private boolean syncingCombo;

                            {
                                combo.setEditable(true);
                                combo.setMaxWidth(Double.MAX_VALUE);
                                enableComboPopupInTableCell(combo, this);
                                combo.valueProperty()
                                        .addListener(
                                                (obs, oldV, newV) -> {
                                                    if (syncingCombo || boundRow == null) {
                                                        return;
                                                    }
                                                    boundRow.applyPickSelection(
                                                            newV, headerPicksSupplier.get());
                                                });
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    boundRow = null;
                                    combo.disableProperty().unbind();
                                    setGraphic(null);
                                } else {
                                    boundRow = getTableView().getItems().get(getIndex());
                                    combo.setItems(pickLabelItems);
                                    combo.disableProperty().unbind();
                                    if (boundRow != null) {
                                        combo.disableProperty()
                                                .bind(
                                                        Bindings.notEqual(
                                                                boundRow.action,
                                                                FixAction.REDEFINE));
                                        syncingCombo = true;
                                        try {
                                            List<JuchuSheetColumnLayout.ExcelHeaderPick> picks =
                                                    headerPicksSupplier.get();
                                            String label = boundRow.getSelectedPickLabel();
                                            if (label == null || label.isBlank()) {
                                                label =
                                                        displayLabelForHeaderText(
                                                                boundRow.getSelectedExcelHeader(),
                                                                picks,
                                                                boundRow.columnLetter());
                                            }
                                            combo.setValue(label);
                                        } finally {
                                            syncingCombo = false;
                                        }
                                    }
                                    setGraphic(combo);
                                }
                            }
                        });

        TableColumn<KnownRow, FixAction> colAction = new TableColumn<>("対応");
        colAction.setCellValueFactory(c -> c.getValue().action);
        colAction.setCellFactory(
                col ->
                        new TableCell<KnownRow, FixAction>() {
                            private final ComboBox<FixAction> combo = new ComboBox<>();
                            private KnownRow boundRow;

                            {
                                combo.setMaxWidth(Double.MAX_VALUE);
                                enableComboPopupInTableCell(combo, this);
                                combo.valueProperty()
                                        .addListener(
                                                (obs, oldV, newV) -> {
                                                    if (boundRow == null || newV == null) {
                                                        return;
                                                    }
                                                    boundRow.setAction(newV);
                                                });
                            }

                            @Override
                            protected void updateItem(FixAction item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    boundRow = null;
                                    setGraphic(null);
                                } else {
                                    boundRow = getTableView().getItems().get(getIndex());
                                    combo.setItems(
                                            FXCollections.observableArrayList(FixAction.values()));
                                    combo.setValue(
                                            boundRow != null ? boundRow.getAction() : FixAction.SKIP);
                                    setGraphic(combo);
                                }
                            }
                        });

        table.getColumns()
                .addAll(colStatus, colFormItem, colLetter, colExpected, colActual, colPick, colAction);
        return table;
    }

    private static TableView<UnknownRow> createUnknownTable(List<UnknownRow> rows) {
        TableView<UnknownRow> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setEditable(true);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);
        table.setPrefHeight(280);

        TableColumn<UnknownRow, String> colStatus = new TableColumn<>("状態");
        colStatus.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().status()));

        TableColumn<UnknownRow, String> colLetter = new TableColumn<>("列");
        colLetter.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().columnLetter()));

        TableColumn<UnknownRow, String> colHeader = new TableColumn<>("Excel見出し");
        colHeader.setCellValueFactory(c -> new SimpleStringProperty(c.getValue().headerText()));

        TableColumn<UnknownRow, JuchuSheetColumnLayout.Col> colTarget = new TableColumn<>("既知列（別名先）");
        colTarget.setCellValueFactory(c -> c.getValue().aliasTarget);
        colTarget.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final ComboBox<JuchuSheetColumnLayout.Col> combo =
                                    new ComboBox<>();
                            private UnknownRow boundRow;

                            {
                                combo.setItems(
                                        FXCollections.observableArrayList(
                                                JuchuSheetColumnLayout.Col.values()));
                                combo.setConverter(
                                        new javafx.util.StringConverter<>() {
                                            @Override
                                            public String toString(
                                                    JuchuSheetColumnLayout.Col object) {
                                                return object == null
                                                        ? ""
                                                        : object.formItemDescription();
                                            }

                                            @Override
                                            public JuchuSheetColumnLayout.Col fromString(
                                                    String string) {
                                                return null;
                                            }
                                        });
                                combo.setMaxWidth(Double.MAX_VALUE);
                                enableComboPopupInTableCell(combo, this);
                                combo.valueProperty()
                                        .addListener(
                                                (obs, oldV, newV) -> {
                                                    if (boundRow != null) {
                                                        boundRow.setAliasTarget(newV);
                                                    }
                                                });
                            }

                            @Override
                            protected void updateItem(
                                    JuchuSheetColumnLayout.Col item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    boundRow = null;
                                    combo.disableProperty().unbind();
                                    setGraphic(null);
                                } else {
                                    boundRow = getTableView().getItems().get(getIndex());
                                    combo.disableProperty().unbind();
                                    if (boundRow != null) {
                                        combo.disableProperty()
                                                .bind(
                                                        Bindings.notEqual(
                                                                boundRow.action,
                                                                UnknownAction.ALIAS_TO_KNOWN));
                                        combo.setValue(boundRow.getAliasTarget());
                                    }
                                    setGraphic(combo);
                                }
                            }
                        });

        TableColumn<UnknownRow, UnknownAction> colAction = new TableColumn<>("対応");
        colAction.setCellValueFactory(c -> c.getValue().action);
        colAction.setCellFactory(
                col ->
                        new TableCell<UnknownRow, UnknownAction>() {
                            private final ComboBox<UnknownAction> combo = new ComboBox<>();

                            {
                                combo.setMaxWidth(Double.MAX_VALUE);
                                enableComboPopupInTableCell(combo, this);
                                combo.valueProperty()
                                        .addListener(
                                                (obs, oldV, newV) -> {
                                                    UnknownRow row =
                                                            getTableView().getItems().get(getIndex());
                                                    if (row != null && newV != null) {
                                                        row.setAction(newV);
                                                    }
                                                });
                            }

                            @Override
                            protected void updateItem(UnknownAction item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    setGraphic(null);
                                } else {
                                    UnknownRow row = getTableView().getItems().get(getIndex());
                                    combo.setItems(
                                            FXCollections.observableArrayList(UnknownAction.values()));
                                    combo.setValue(
                                            row != null ? row.getAction() : UnknownAction.SKIP);
                                    setGraphic(combo);
                                }
                            }
                        });

        table.getColumns().addAll(colStatus, colLetter, colHeader, colTarget, colAction);
        return table;
    }

    private static void applyAll(
            List<KnownRow> knownRows,
            List<UnknownRow> unknownRows,
            JuchuHeaderAliasRegistry registry,
            String pathKey,
            List<JuchuSheetColumnLayout.ExcelHeaderPick> headerPicks)
            throws Exception {
        validateBeforeApply(knownRows, headerPicks);
        validateBeforeApplyUnknown(unknownRows);
        boolean changed = applyKnownFixes(knownRows, registry, pathKey, headerPicks);
        changed |= applyUnknownFixes(unknownRows, registry, pathKey);
        if (changed) {
            registry.saveToDisk();
        }
    }

    private static void commitKnownRowPickSelections(
            List<KnownRow> rows, List<JuchuSheetColumnLayout.ExcelHeaderPick> headerPicks) {
        for (KnownRow row : rows) {
            String label = row.getSelectedPickLabel();
            if (label != null && !label.isBlank()) {
                row.applyPickSelection(label, headerPicks);
            }
        }
    }

    private static boolean applyKnownFixes(
            List<KnownRow> rows,
            JuchuHeaderAliasRegistry registry,
            String pathKey,
            List<JuchuSheetColumnLayout.ExcelHeaderPick> headerPicks) {
        boolean changed = false;
        for (KnownRow row : rows) {
            JuchuHeaderMismatch m = row.mismatch;
            switch (row.getAction()) {
                case REDEFINE -> {
                    String header = resolveSelectedHeaderText(row, headerPicks);
                    if (header.isBlank()) {
                        throw new IllegalStateException(
                                m.columnLetter() + "列: 採用する Excel 見出しを選んでください。");
                    }
                    registry.setExpectedOverride(pathKey, m.column(), header);
                    String pickLabel = row.getSelectedPickLabel();
                    if (pickLabel != null && !pickLabel.isBlank()) {
                        registry.setExpectedPickLabel(pathKey, m.column(), pickLabel);
                    } else {
                        registry.setExpectedPickLabel(pathKey, m.column(), null);
                    }
                    registry.clearExcludedFromTransfer(pathKey, m.column());
                    if (!m.actualEmpty()) {
                        String actual = m.actualHeader().strip();
                        if (!actual.isEmpty()
                                && !JuchuSheetColumnLayout.normalizeHeader(actual)
                                        .equals(JuchuSheetColumnLayout.normalizeHeader(header))) {
                            registry.addAlias(pathKey, m.column(), actual);
                        }
                    }
                    changed = true;
                }
                case ALIAS -> {
                    if (!m.actualEmpty()) {
                        registry.addAlias(pathKey, m.column(), m.actualHeader());
                        registry.clearExcludedFromTransfer(pathKey, m.column());
                        changed = true;
                    }
                }
                case EXCLUDE -> {
                    registry.setExcludedFromTransfer(pathKey, m.column());
                    changed = true;
                }
                case SKIP -> {
                    if (registry.isExcludedFromTransfer(pathKey, m.column())) {
                        registry.clearExcludedFromTransfer(pathKey, m.column());
                        changed = true;
                    }
                }
            }
        }
        return changed;
    }

    private static boolean applyUnknownFixes(
            List<UnknownRow> rows, JuchuHeaderAliasRegistry registry, String pathKey) {
        boolean changed = false;
        for (UnknownRow row : rows) {
            String letter = row.columnLetter();
            switch (row.getAction()) {
                case IGNORE -> {
                    registry.setUnknownColumnIgnored(pathKey, letter);
                    changed = true;
                }
                case ALIAS_TO_KNOWN -> {
                    JuchuSheetColumnLayout.Col target = row.getAliasTarget();
                    if (target == null) {
                        throw new IllegalStateException(
                                letter + "列: 別名登録先の既知列を選んでください。");
                    }
                    registry.addAlias(pathKey, target, row.headerText());
                    registry.clearUnknownColumnIgnored(pathKey, letter);
                    changed = true;
                }
                case SKIP -> {
                    if (row.column.ignored()) {
                        registry.clearUnknownColumnIgnored(pathKey, letter);
                        changed = true;
                    }
                }
            }
        }
        return changed;
    }

    private record SheetContext(Row headerRow, List<JuchuSheetColumnLayout.ExcelHeaderPick> headerPicks) {}

    private static SheetContext loadSheetContext(File juchuFile, JuchuHeaderAliasRegistry registry)
            throws Exception {
        try (FileInputStream fis = new FileInputStream(juchuFile);
                Workbook wb = PoiWorkbookOpener.open(fis)) {
            Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");
            if (sheet == null) {
                throw new IllegalStateException("受注ﾌｧｲﾙ シートが見つかりません。");
            }
            int headerRowIndex =
                    JuchuSheetColumnLayout.resolveHeaderRowIndex(
                            registry, juchuFile.getAbsolutePath());
            Row hRow = sheet.getRow(headerRowIndex);
            if (hRow == null) {
                int rowOneBased = headerRowIndex + 1;
                throw new IllegalStateException(
                        "受注ﾌｧｲﾙ: 見出し行（行" + rowOneBased + "）が存在しません。");
            }
            return new SheetContext(hRow, JuchuSheetColumnLayout.readExcelHeaderPicks(hRow));
        }
    }

    static String defaultSelectedPickLabel(
            JuchuHeaderMismatch mismatch,
            List<JuchuSheetColumnLayout.ExcelHeaderPick> picks,
            JuchuHeaderAliasRegistry registry,
            String pathKey) {
        if (registry != null && pathKey != null) {
            var savedPick = registry.expectedPickLabelFor(pathKey, mismatch.column());
            if (savedPick.isPresent()) {
                String label = savedPick.get().strip();
                if (!label.isEmpty()
                        && (resolvePick(label, picks) != null
                                || headerTextForPickLabel(label, picks).equals(
                                        registry
                                                .expectedOverrideFor(pathKey, mismatch.column())
                                                .orElse("")))) {
                    return label;
                }
            }
            var overrideOpt = registry.expectedOverrideFor(pathKey, mismatch.column());
            if (overrideOpt.isPresent()) {
                String label = displayLabelForHeaderText(overrideOpt.get(), picks, null);
                if (!label.isBlank()) {
                    return label;
                }
            }
        }
        if (!mismatch.actualEmpty()) {
            for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
                if (pick.columnLetter().equals(mismatch.columnLetter())) {
                    return pick.displayLabel();
                }
            }
            return mismatch.actualHeader();
        }
        String letter = mismatch.columnLetter();
        for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
            if (pick.columnLetter().equals(letter)) {
                return pick.displayLabel();
            }
        }
        String primary = mismatch.column().primaryHeader();
        for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
            if (JuchuSheetColumnLayout.normalizeHeader(pick.headerText())
                    .equals(JuchuSheetColumnLayout.normalizeHeader(primary))) {
                return pick.displayLabel();
            }
        }
        return primary;
    }

    static JuchuSheetColumnLayout.ExcelHeaderPick resolvePick(
            String comboValue, List<JuchuSheetColumnLayout.ExcelHeaderPick> picks) {
        if (comboValue == null || picks == null) {
            return null;
        }
        String trimmed = comboValue.strip();
        if (trimmed.isEmpty()) {
            return null;
        }
        for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
            if (pick.displayLabel().equals(trimmed)) {
                return pick;
            }
        }
        int colon = trimmed.indexOf("列:");
        if (colon > 0) {
            String letter = trimmed.substring(0, colon).strip().toUpperCase();
            String headerPart = trimmed.substring(colon + 2).strip();
            for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
                if (pick.columnLetter().equalsIgnoreCase(letter)
                        && pick.headerText().equals(headerPart)) {
                    return pick;
                }
            }
            for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
                if (pick.columnLetter().equalsIgnoreCase(letter)) {
                    return pick;
                }
            }
        }
        for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
            if (pick.headerText().equals(trimmed)) {
                return pick;
            }
        }
        return null;
    }

    private static String headerTextForPickLabel(
            String pickLabel, List<JuchuSheetColumnLayout.ExcelHeaderPick> picks) {
        JuchuSheetColumnLayout.ExcelHeaderPick pick = resolvePick(pickLabel, picks);
        if (pick != null) {
            return pick.headerText();
        }
        return pickLabel != null ? pickLabel.strip() : "";
    }

    private static List<String> headerPickLabels(
            List<JuchuSheetColumnLayout.ExcelHeaderPick> picks) {
        Set<String> labels = new LinkedHashSet<>();
        for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
            labels.add(pick.displayLabel());
        }
        return List.copyOf(labels);
    }

    private static String displayLabelForHeaderText(
            String headerText,
            List<JuchuSheetColumnLayout.ExcelHeaderPick> picks,
            String preferColumnLetter) {
        if (headerText == null || headerText.isBlank()) {
            return "";
        }
        if (preferColumnLetter != null && !preferColumnLetter.isBlank()) {
            for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
                if (pick.columnLetter().equalsIgnoreCase(preferColumnLetter.strip())
                        && JuchuSheetColumnLayout.normalizeHeader(pick.headerText())
                                .equals(JuchuSheetColumnLayout.normalizeHeader(headerText))) {
                    return pick.displayLabel();
                }
            }
        }
        for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
            if (JuchuSheetColumnLayout.normalizeHeader(pick.headerText())
                    .equals(JuchuSheetColumnLayout.normalizeHeader(headerText))) {
                return pick.displayLabel();
            }
        }
        return headerText;
    }

    private static void enableComboPopupInTableCell(ComboBox<?> combo, TableCell<?, ?> cell) {
        combo.addEventFilter(
                MouseEvent.MOUSE_RELEASED,
                event -> {
                    if (!combo.isDisabled()) {
                        Platform.runLater(combo::show);
                    }
                    event.consume();
                });
        cell.addEventFilter(
                MouseEvent.MOUSE_RELEASED,
                event -> {
                    if (!cell.isEmpty() && !combo.isDisabled()) {
                        Platform.runLater(combo::show);
                        event.consume();
                    }
                });
    }

    private static String resolveSelectedHeaderText(
            KnownRow row, List<JuchuSheetColumnLayout.ExcelHeaderPick> headerPicks) {
        String pickLabel = row.getSelectedPickLabel();
        if (pickLabel != null && !pickLabel.isBlank()) {
            String fromPick = headerTextForPickLabel(pickLabel, headerPicks).strip();
            if (!fromPick.isBlank()) {
                return fromPick;
            }
        }
        String header = row.getSelectedExcelHeader();
        if (header != null && !header.isBlank()) {
            return header.strip();
        }
        return headerTextForPickLabel(row.getSelectedPickLabel(), headerPicks).strip();
    }

    private static void validateBeforeApply(
            List<KnownRow> rows, List<JuchuSheetColumnLayout.ExcelHeaderPick> headerPicks) {
        for (KnownRow row : rows) {
            if (row.getAction() != FixAction.REDEFINE) {
                continue;
            }
            if (resolveSelectedHeaderText(row, headerPicks).isBlank()) {
                throw new IllegalStateException(
                        row.columnLetter() + "列: 採用する Excel 見出しを選んでください。");
            }
        }
    }

    private static void validateBeforeApplyUnknown(List<UnknownRow> rows) {
        for (UnknownRow row : rows) {
            if (row.getAction() == UnknownAction.ALIAS_TO_KNOWN && row.getAliasTarget() == null) {
                throw new IllegalStateException(
                        row.columnLetter() + "列: 別名登録先の既知列を選んでください。");
            }
        }
    }

    static List<JuchuHeaderMismatch> readMismatches(
            File juchuFile, JuchuHeaderAliasRegistry registry) throws Exception {
        try (FileInputStream fis = new FileInputStream(juchuFile);
                Workbook wb = PoiWorkbookOpener.open(fis)) {
            Sheet sheet = wb.getSheet("受注ﾌｧｲﾙ");
            if (sheet == null) {
                return List.of();
            }
            int headerRowIndex =
                    JuchuSheetColumnLayout.resolveHeaderRowIndex(
                            registry, juchuFile.getAbsolutePath());
            Row hRow = sheet.getRow(headerRowIndex);
            return JuchuSheetColumnLayout.collectHeaderMismatches(
                    hRow, registry, juchuFile.getAbsolutePath());
        }
    }

    private static void showError(Window owner, String message) {
        Alert alert = new Alert(Alert.AlertType.ERROR);
        if (owner != null) {
            alert.initOwner(owner);
        }
        alert.setTitle("受注シート列定義");
        alert.setHeaderText(null);
        TextArea area = new TextArea(message);
        area.setEditable(false);
        area.setWrapText(true);
        area.setPrefRowCount(3);
        alert.getDialogPane().setContent(area);
        alert.getButtonTypes().setAll(ButtonType.OK);
        alert.showAndWait();
    }
}
