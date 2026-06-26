package jp.co.pm.ai.desktop.reconciliation;

import javafx.application.Platform;
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
 * 受注ﾌｧｲﾙの列定義ウィザード。
 * 依頼書フォームの各項目について、採用する受注列（転記・読込・検証の物理列）を選ぶ。
 */
public final class JuchuSheetHeaderRepairWizard {

    public enum Result {
        CANCEL,
        CONTINUE,
        FIXED
    }

    public enum FixAction {
        REDEFINE("この列を採用"),
        ALIAS("実際の見出しを別名として許容"),
        EXCLUDE("転記しない"),
        SKIP("変更なし");

        private final String label;

        FixAction(String label) {
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
            return rowStatus(this, registry, path, List.of());
        }

        /** {@link #status} の拡張（採用列候補一覧付き）。 */
        static String rowStatus(
                KnownRow row,
                JuchuHeaderAliasRegistry registry,
                String pathKey,
                List<JuchuSheetColumnLayout.ExcelHeaderPick> picks) {
            if (registry != null && registry.isExcludedFromTransfer(pathKey, row.mismatch.column())) {
                return "転記除外";
            }
            if (isAdoptionSaved(registry, pathKey, row.mismatch.column())) {
                String current = row.getSelectedPickLabel();
                var saved = registry.expectedPickLabelFor(pathKey, row.mismatch.column());
                if (saved.isPresent()
                        && current != null
                        && !current.isBlank()
                        && !current.strip().equals(saved.get().strip())) {
                    return "未保存";
                }
                return "採用済";
            }
            if (needsAdoptionPersist(row, registry, pathKey, picks)) {
                return "要適用";
            }
            return row.matching(registry, pathKey) ? "一致" : "不一致";
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
        CheckBox chkMismatchesOnly = new CheckBox("不一致のフォーム項目のみ表示");
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

        VBox formMappingBox = new VBox(8, chkMismatchesOnly, knownTable);
        VBox.setVgrow(knownTable, Priority.ALWAYS);
        formMappingBox.setPadding(new Insets(8, 0, 0, 0));

        Label mismatchBanner = new Label("");
        mismatchBanner.setWrapText(true);
        mismatchBanner.setManaged(false);
        mismatchBanner.setVisible(false);
        if (mode == DialogMode.TRANSFER_PROMPT) {
            int mismatchCount = transferMismatches != null ? transferMismatches.size() : 0;
            mismatchBanner.setText("転記前に確認: 見出し不一致 " + mismatchCount + " 件");
            mismatchBanner.setStyle("-fx-font-weight: bold;");
            mismatchBanner.setManaged(true);
            mismatchBanner.setVisible(true);
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
                        refreshKnownTableItems(
                                knownTable,
                                allKnownRows,
                                registry,
                                pathKey,
                                chkMismatchesOnly.isSelected());
                        knownTable.refresh();
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
        stage.setTitle("受注列の採用 — フォーム転記");

        String fileName = juchuFile.getName();
        Label intro =
                new Label(
                        mode == DialogMode.TRANSFER_PROMPT
                                ? "受注ファイル「"
                                        + fileName
                                        + "」を転記する前に、フォーム項目ごとに採用する受注列を確認してください。"
                                        + " 推定候補はあらかじめ選ばれています。必要なら変更して「適用して再検証」を押してください。"
                                : "依頼書フォームの転記項目について、受注ファイル「"
                                        + fileName
                                        + "」のどの列見出しを採用するかを設定します。"
                                        + " 見出し行の変更や、項目ごとの採用列の修正ができます。");
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
                        int promoted =
                                promoteRowsNeedingAdoptionPersist(
                                        allKnownRows, registry, pathKey, excelHeaderPicks);
                        applyAll(allKnownRows, registry, pathKey, excelHeaderPicks);
                        registry.saveToDisk();
                        List<JuchuHeaderMismatch> remaining = readMismatches(juchuFile, registry);
                        reloadSheetRows.run();
                        if (remaining.isEmpty()) {
                            statusLabel.setText(
                                    promoted > 0
                                            ? "採用列を "
                                                    + promoted
                                                    + " 件保存しました。不一致は解消されています。"
                                            : "採用列を適用しました。不一致は解消されています。");
                            mismatchBanner.setManaged(false);
                            mismatchBanner.setVisible(false);
                            if (mode == DialogMode.TRANSFER_PROMPT) {
                                outcome[0] = Result.FIXED;
                                stage.close();
                            }
                        } else {
                            statusLabel.setText(
                                    "不一致が "
                                            + remaining.size()
                                            + " 件残っています。採用列を見直して再度「適用」してください。");
                            mismatchBanner.setText("見出し不一致 " + remaining.size() + " 件");
                            mismatchBanner.setManaged(true);
                            mismatchBanner.setVisible(true);
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

        VBox center = new VBox(10, intro, headerRowRow, mismatchBanner, formMappingBox, statusLabel);
        center.setPadding(new Insets(12));
        VBox.setVgrow(formMappingBox, Priority.ALWAYS);

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
            String defaultPick =
                    defaultSelectedPickLabel(m, ctx.headerPicks(), registry, pathKey);
            JuchuSheetColumnLayout.ExcelHeaderPick resolvedDefault =
                    resolvePick(defaultPick, ctx.headerPicks());
            boolean crossColumnDefault =
                    resolvedDefault != null
                            && !resolvedDefault
                                    .columnLetter()
                                    .equalsIgnoreCase(m.columnLetter());
            boolean adoptionSaved = isAdoptionSaved(registry, pathKey, m.column());
            FixAction defaultAction;
            if (registry.isExcludedFromTransfer(pathKey, m.column())) {
                defaultAction = FixAction.EXCLUDE;
            } else if (!adoptionSaved
                    && (crossColumnDefault
                            || !JuchuSheetColumnLayout.headerMatches(
                                    m.column(), m.actualHeader(), registry, pathKey))) {
                defaultAction = FixAction.REDEFINE;
            } else if (JuchuSheetColumnLayout.headerMatches(
                    m.column(), m.actualHeader(), registry, pathKey)) {
                defaultAction = FixAction.SKIP;
            } else {
                defaultAction = FixAction.REDEFINE;
            }
            rows.add(
                    new KnownRow(
                            m,
                            defaultAction,
                            defaultPick,
                            headerTextForPickLabel(defaultPick, ctx.headerPicks())));
        }
        return rows;
    }

    private static List<FixAction> fixActionsForWizard() {
        return List.of(FixAction.REDEFINE, FixAction.SKIP, FixAction.EXCLUDE);
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
                                KnownRow.rowStatus(
                                        c.getValue(),
                                        registry,
                                        pathKey,
                                        headerPicksSupplier.get())));

        TableColumn<KnownRow, String> colFormItem = new TableColumn<>("フォーム項目");
        colFormItem.setCellValueFactory(
                c -> new SimpleStringProperty(c.getValue().formItem()));

        TableColumn<KnownRow, String> colPick = new TableColumn<>("採用する受注列");
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
                                                    if (newV != null && !newV.isBlank()) {
                                                        boundRow.setAction(FixAction.REDEFINE);
                                                    }
                                                });
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    boundRow = null;
                                    setGraphic(null);
                                } else {
                                    boundRow = getTableView().getItems().get(getIndex());
                                    combo.setItems(pickLabelItems);
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
                                    setGraphic(combo);
                                }
                            }
                        });

        TableColumn<KnownRow, FixAction> colAction = new TableColumn<>("転記");
        colAction.setCellValueFactory(c -> c.getValue().action);
        colAction.setCellFactory(
                col ->
                        new TableCell<KnownRow, FixAction>() {
                            private final ComboBox<FixAction> combo = new ComboBox<>();
                            private KnownRow boundRow;

                            {
                                combo.setMaxWidth(Double.MAX_VALUE);
                                enableComboPopupInTableCell(combo, this);
                                combo.setConverter(
                                        new javafx.util.StringConverter<>() {
                                            @Override
                                            public String toString(FixAction action) {
                                                if (action == null) {
                                                    return "";
                                                }
                                                if (action == FixAction.SKIP
                                                        && boundRow != null
                                                        && isAdoptionSaved(
                                                                registry,
                                                                pathKey,
                                                                boundRow.mismatch.column())) {
                                                    return "採用済";
                                                }
                                                return action.toString();
                                            }

                                            @Override
                                            public FixAction fromString(String string) {
                                                if (string == null || string.isBlank()) {
                                                    return FixAction.SKIP;
                                                }
                                                if ("採用済".equals(string.strip())) {
                                                    return FixAction.SKIP;
                                                }
                                                for (FixAction action : fixActionsForWizard()) {
                                                    if (action.toString().equals(string)) {
                                                        return action;
                                                    }
                                                }
                                                return FixAction.SKIP;
                                            }
                                        });
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
                                            FXCollections.observableArrayList(fixActionsForWizard()));
                                    combo.setValue(
                                            boundRow != null ? boundRow.getAction() : FixAction.SKIP);
                                    setGraphic(combo);
                                }
                            }
                        });

        table.getColumns().addAll(colStatus, colFormItem, colPick, colAction);
        return table;
    }

    private static void applyAll(
            List<KnownRow> knownRows,
            JuchuHeaderAliasRegistry registry,
            String pathKey,
            List<JuchuSheetColumnLayout.ExcelHeaderPick> headerPicks)
            throws Exception {
        validateBeforeApply(knownRows, headerPicks);
        boolean changed = applyKnownFixes(knownRows, registry, pathKey, headerPicks);
        if (changed) {
            registry.saveToDisk();
        }
    }

    /** 「変更なし」のままでも、未保存の採用列があれば REDEFINE へ昇格する。 */
    static int promoteRowsNeedingAdoptionPersist(
            List<KnownRow> rows,
            JuchuHeaderAliasRegistry registry,
            String pathKey,
            List<JuchuSheetColumnLayout.ExcelHeaderPick> picks) {
        int promoted = 0;
        for (KnownRow row : rows) {
            if (row.getAction() == FixAction.EXCLUDE) {
                continue;
            }
            if (needsAdoptionPersist(row, registry, pathKey, picks)) {
                row.setAction(FixAction.REDEFINE);
                promoted++;
            }
        }
        return promoted;
    }

    static boolean isAdoptionSaved(
            JuchuHeaderAliasRegistry registry,
            String pathKey,
            JuchuSheetColumnLayout.Col column) {
        if (registry == null || pathKey == null || column == null) {
            return false;
        }
        return registry
                .expectedPickLabelFor(pathKey, column)
                .map(label -> !label.isBlank())
                .orElse(false);
    }

    static boolean needsAdoptionPersist(
            KnownRow row,
            JuchuHeaderAliasRegistry registry,
            String pathKey,
            List<JuchuSheetColumnLayout.ExcelHeaderPick> picks) {
        if (row.getAction() == FixAction.EXCLUDE) {
            return false;
        }
        String label = row.getSelectedPickLabel();
        if (label == null || label.isBlank()) {
            return false;
        }
        String stripped = label.strip();
        if (registry == null || pathKey == null) {
            return resolvePick(stripped, picks) != null;
        }
        if (isAdoptionSaved(registry, pathKey, row.mismatch.column())) {
            var saved = registry.expectedPickLabelFor(pathKey, row.mismatch.column());
            return saved.isEmpty() || !saved.get().strip().equals(stripped);
        }
        JuchuSheetColumnLayout.ExcelHeaderPick pick = resolvePick(stripped, picks);
        String header = pick != null ? pick.headerText() : stripped;
        var savedExpected = registry.expectedOverrideFor(pathKey, row.mismatch.column());
        if (savedExpected.isEmpty()) {
            return pick != null || !row.mismatch.actualEmpty();
        }
        return !savedExpected.get().equals(header);
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
                                m.formItemDescription() + ": 採用する受注列を選んでください。");
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
        JuchuSheetColumnLayout.ExcelHeaderPick bestPick =
                findBestMatchingPick(mismatch, picks, registry, pathKey);
        if (bestPick != null) {
            return bestPick.displayLabel();
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

    /** フォーム項目の見出し候補と一致する受注列を推定。 */
    static JuchuSheetColumnLayout.ExcelHeaderPick findBestMatchingPick(
            JuchuHeaderMismatch mismatch,
            List<JuchuSheetColumnLayout.ExcelHeaderPick> picks,
            JuchuHeaderAliasRegistry registry,
            String pathKey) {
        if (picks == null || picks.isEmpty()) {
            return null;
        }
        Set<String> candidates = new LinkedHashSet<>();
        JuchuSheetColumnLayout.Col col = mismatch.column();
        candidates.add(col.primaryHeader());
        candidates.addAll(col.aliases());
        if (mismatch.expectedHeader() != null && !mismatch.expectedHeader().isBlank()) {
            candidates.add(mismatch.expectedHeader());
        }
        if (registry != null && pathKey != null) {
            registry.expectedOverrideFor(pathKey, col).ifPresent(candidates::add);
            candidates.addAll(registry.extraAliasesFor(pathKey, col));
        }
        for (JuchuSheetColumnLayout.ExcelHeaderPick pick : picks) {
            for (String candidate : candidates) {
                if (headerTextsMatch(pick.headerText(), candidate)) {
                    return pick;
                }
            }
        }
        return null;
    }

    /** 見出し文字列の一致（完全一致または意味のある部分一致）。 */
    static boolean headerTextsMatch(String pickHeader, String candidate) {
        if (pickHeader == null || candidate == null) {
            return false;
        }
        String normPick = JuchuSheetColumnLayout.normalizeHeader(pickHeader);
        String normCand = JuchuSheetColumnLayout.normalizeHeader(candidate);
        if (normPick.isEmpty() || normCand.isEmpty()) {
            return false;
        }
        if (normPick.equals(normCand)) {
            return true;
        }
        return (normPick.contains(normCand) || normCand.contains(normPick))
                && Math.min(normPick.length(), normCand.length()) >= 2;
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
                        row.formItem() + ": 採用する受注列を選んでください。");
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
