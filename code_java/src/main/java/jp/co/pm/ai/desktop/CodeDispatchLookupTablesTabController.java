package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicBoolean;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.AppPaths.DispatchLookupTableOverwriteResult;
import jp.co.pm.ai.desktop.io.CodeDispatchLookupTableIo;
import jp.co.pm.ai.desktop.io.CodeDispatchLookupTableIo.KeyValTable;

/**
 * 「材料・製品種類情報」: {@code code/} 配下のキー・値テーブルをタブ切替で編集する。
 */
public final class CodeDispatchLookupTablesTabController {

    private record FileSpec(String relativePath, String defaultHeaderLine, String tabLabel) {}

    private static final List<FileSpec> FILES =
            List.of(
                    new FileSpec(
                            CodeDispatchLookupTablesMerge.FILE_USED_RAW_ROLL,
                            CodeDispatchLookupTablesMerge.FILE_USED_RAW_ROLL.replace(".txt", ""),
                            "使用原反→ロール長(m)"),
                    new FileSpec(
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_ROLL,
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_ROLL.replace(".txt", ""),
                            "製品名→ロール長(m)"),
                    new FileSpec(
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_WIDTH,
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_WIDTH.replace(".txt", ""),
                            "製品名→製品幅(mm)"),
                    new FileSpec(
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_THICK,
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_THICK.replace(".txt", ""),
                            "製品名→厚み(mm)"),
                    new FileSpec(
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_LENGTH,
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_LENGTH.replace(".txt", ""),
                            "製品名→製品長(mm)"),
                    new FileSpec(
                            CodeDispatchLookupTablesMerge.FILE_USED_RAW_WIDTH,
                            "使用原反,原反幅",
                            "使用原反→原反幅(mm)"));

    private MainShellController shell;

    @FXML
    private Label hintLabel;

    @FXML
    private TabPane fileTabPane;

    private final Map<Tab, FilePanel> panelByTab = new ConcurrentHashMap<>();
    private final Map<String, FilePanel> panelByFilename = new LinkedHashMap<>();

    private RepoRestorePanel repoRestorePanel;

    private final AtomicBoolean planInputRollUnitNotifyPending = new AtomicBoolean(false);

    @FXML
    private void initialize() {
        hintLabel.setText(
                "サマリ Excel（PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK）と同一フォルダの材料・製品種類テーブルを編集します（UTF-8）。"
                        + " 段階1が正常終了したとき、plan_input_tasks の製品名・使用原反で不足キーのみ自動追記します。"
                        + " 値が空欄の行があると段階2・段階3は実行できません。");
        for (FileSpec spec : FILES) {
            Tab tab = new Tab(spec.tabLabel());
            FilePanel panel = new FilePanel(spec);
            tab.setContent(panel.root());
            panelByTab.put(tab, panel);
            panelByFilename.put(spec.relativePath(), panel);
            fileTabPane.getTabs().add(tab);
        }
        repoRestorePanel = new RepoRestorePanel();
        Tab restoreTab = new Tab("リポジトリから上書き");
        restoreTab.setContent(repoRestorePanel.root());
        fileTabPane.getTabs().add(restoreTab);
        fileTabPane
                .getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (obs, prev, sel) -> {
                            if (sel != null) {
                                FilePanel p = panelByTab.get(sel);
                                if (p != null) {
                                    p.ensureLoaded();
                                } else if (repoRestorePanel != null && sel.getContent() == repoRestorePanel.root()) {
                                    repoRestorePanel.refreshStatus();
                                }
                            }
                        });
        Platform.runLater(
                () -> {
                    Tab first = fileTabPane.getTabs().isEmpty() ? null : fileTabPane.getTabs().getFirst();
                    if (first != null) {
                        fileTabPane.getSelectionModel().select(first);
                        FilePanel p = panelByTab.get(first);
                        if (p != null) {
                            p.ensureLoaded();
                        }
                    }
                });
    }

    void bindShell(MainShellController mainShellController) {
        this.shell = mainShellController;
    }

    /** 全サブタブをディスクから再読込（段階1マージ直後など）。 */
    void reloadAllFromDisk() {
        for (FilePanel p : panelByTab.values()) {
            p.reloadFromDisk();
        }
        scheduleInvalidatePlanInputRollUnitHighlightCache();
    }

    int snapshotInnerTabSelectedIndex() {
        if (fileTabPane == null) {
            return -1;
        }
        return fileTabPane.getSelectionModel().getSelectedIndex();
    }

    void applyInnerTabSelectedIndex(int index) {
        if (fileTabPane == null || fileTabPane.getTabs().isEmpty()) {
            return;
        }
        int i = Math.max(0, Math.min(index, fileTabPane.getTabs().size() - 1));
        fileTabPane.getSelectionModel().select(i);
    }

    private Map<String, String> uiEnv() {
        return shell != null ? shell.snapshotUiEnv() : Map.of();
    }

    private void logLine(String line) {
        if (shell != null) {
            shell.appendLog(line);
        }
    }

    /** 保存・再読込でディスク上のルックアップ表が変わったあと、配台計画_タスク入力のロール長キャッシュを1回だけ無効化する。 */
    private void scheduleInvalidatePlanInputRollUnitHighlightCache() {
        if (shell == null) {
            return;
        }
        if (!planInputRollUnitNotifyPending.compareAndSet(false, true)) {
            return;
        }
        Platform.runLater(
                () -> {
                    planInputRollUnitNotifyPending.set(false);
                    shell.invalidatePlanInputRollUnitHighlightCache();
                });
    }

    private final class FilePanel {

        private final FileSpec spec;
        private final VBox root;
        private final Label pathLabel = new Label();
        private final TextField rowSearchField = new TextField();
        private final TableView<ObservableList<String>> table = new TableView<>();
        private final ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        private final FilteredList<ObservableList<String>> rowsFiltered;
        private volatile boolean loaded;

        FilePanel(FileSpec spec) {
            this.spec = spec;
            rowsFiltered = new FilteredList<>(rows, this::rowMatchesSearch);
            rowSearchField.setPromptText(
                    "キー・値のいずれかに含まれる文字列（部分一致）。空欄ですべて表示");
            HBox.setHgrow(rowSearchField, Priority.ALWAYS);
            rowSearchField
                    .textProperty()
                    .addListener((obs, prev, cur) -> applySearchPredicate());
            rows.addListener(
                    (ListChangeListener<ObservableList<String>>)
                            c -> {
                                while (c.next()) {
                                    if (c.wasAdded()) {
                                        for (ObservableList<String> r : c.getAddedSubList()) {
                                            hookRowForSearchFilter(r);
                                        }
                                    }
                                }
                            });
            TableColumn<ObservableList<String>, String> colKey = new TableColumn<>("キー");
            colKey.setPrefWidth(420);
            colKey.setCellValueFactory(
                    cd -> {
                        ObservableList<String> r = cd.getValue();
                        String v = r != null && !r.isEmpty() ? r.get(0) : "";
                        return new javafx.beans.property.SimpleStringProperty(v);
                    });
            colKey.setCellFactory(TextFieldTableCell.forTableColumn());
            colKey.setOnEditCommit(
                    ev -> {
                        ObservableList<String> r = ev.getRowValue();
                        if (r == null) {
                            return;
                        }
                        while (r.size() < 2) {
                            r.add("");
                        }
                        r.set(0, ev.getNewValue() != null ? ev.getNewValue() : "");
                        applySearchPredicate();
                    });
            TableColumn<ObservableList<String>, String> colVal = new TableColumn<>("値");
            colVal.setPrefWidth(160);
            colVal.setCellValueFactory(
                    cd -> {
                        ObservableList<String> r = cd.getValue();
                        String v = r != null && r.size() > 1 ? r.get(1) : "";
                        return new javafx.beans.property.SimpleStringProperty(v);
                    });
            colVal.setCellFactory(TextFieldTableCell.forTableColumn());
            colVal.setOnEditCommit(
                    ev -> {
                        ObservableList<String> r = ev.getRowValue();
                        if (r == null) {
                            return;
                        }
                        while (r.size() < 2) {
                            r.add("");
                        }
                        r.set(1, ev.getNewValue() != null ? ev.getNewValue() : "");
                        applySearchPredicate();
                    });
            table.getColumns().addAll(List.of(colKey, colVal));
            table.setItems(rowsFiltered);
            table.setEditable(true);
            table.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
            table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);

            Button reload = new Button("再読込");
            reload.setOnAction(e -> reloadFromDisk());
            Button save = new Button("保存");
            save.setOnAction(e -> saveToDisk());
            Button add = new Button("行追加");
            add.setOnAction(e -> rows.add(FXCollections.observableArrayList("", "")));
            Button remove = new Button("行削除");
            remove.setOnAction(
                    e -> {
                        var sel = table.getSelectionModel().getSelectedItems();
                        if (sel == null || sel.isEmpty()) {
                            return;
                        }
                        rows.removeAll(sel);
                    });
            HBox bar = new HBox(8, rowSearchField, reload, save, add, remove);
            bar.setAlignment(Pos.CENTER_LEFT);
            pathLabel.setWrapText(true);
            root = new VBox(8, pathLabel, bar, table);
            VBox.setVgrow(table, Priority.ALWAYS);
            root.setPadding(new Insets(0, 0, 4, 0));
        }

        VBox root() {
            return root;
        }

        void ensureLoaded() {
            if (!loaded) {
                reloadFromDisk();
            }
        }

        void reloadFromDisk() {
            boolean wasLoaded = loaded;
            Path path = resolvePath();
            pathLabel.setText(path.toString());
            try {
                KeyValTable t =
                        CodeDispatchLookupTableIo.readOrEmpty(path, spec.defaultHeaderLine());
                rows.clear();
                for (Map.Entry<String, String> e : t.rows().entrySet()) {
                    ObservableList<String> row =
                            FXCollections.observableArrayList(e.getKey(), e.getValue());
                    hookRowForSearchFilter(row);
                    rows.add(row);
                }
                applySearchPredicate();
                loaded = true;
                logLine("[code-lookup] 読込: " + path);
                if (wasLoaded) {
                    scheduleInvalidatePlanInputRollUnitHighlightCache();
                }
            } catch (IOException ex) {
                logLine("[code-lookup] 読込失敗: " + path + " → " + ex.getMessage());
            }
            table.refresh();
        }

        void saveToDisk() {
            Path path = resolvePath();
            try {
                LinkedHashMap<String, String> m = new LinkedHashMap<>();
                for (ObservableList<String> r : rows) {
                    if (r == null) {
                        continue;
                    }
                    String k = r.isEmpty() ? "" : r.get(0) != null ? r.get(0).strip() : "";
                    if (k.isEmpty()) {
                        continue;
                    }
                    String v = r.size() > 1 && r.get(1) != null ? r.get(1).strip() : "";
                    m.put(k, v);
                }
                KeyValTable cur = CodeDispatchLookupTableIo.readOrEmpty(path, spec.defaultHeaderLine());
                CodeDispatchLookupTableIo.write(path, new KeyValTable(cur.headerLine(), m));
                logLine("[code-lookup] 保存: " + path + " (" + m.size() + " 行)");
                scheduleInvalidatePlanInputRollUnitHighlightCache();
            } catch (IOException ex) {
                logLine("[code-lookup] 保存失敗: " + path + " → " + ex.getMessage());
                if (shell != null) {
                    shell.showErrorDialog("保存", "保存に失敗しました。\n" + ex.getMessage());
                }
            }
        }

        private Path resolvePath() {
            Map<String, String> ui = uiEnv();
            AppPaths.ensureDispatchLookupTableFromRepoIfMissing(ui, spec.relativePath());
            return AppPaths.dispatchLookupTablePath(ui, spec.relativePath());
        }

        private void applySearchPredicate() {
            rowsFiltered.setPredicate(this::rowMatchesSearch);
        }

        private String normalizedSearchQuery() {
            String t = rowSearchField.getText();
            return t != null ? t.trim().toLowerCase(Locale.ROOT) : "";
        }

        private boolean rowMatchesSearch(ObservableList<String> r) {
            if (r == null) {
                return false;
            }
            String q = normalizedSearchQuery();
            if (q.isEmpty()) {
                return true;
            }
            String key = r.isEmpty() || r.get(0) == null ? "" : r.get(0);
            String val = r.size() > 1 && r.get(1) != null ? r.get(1) : "";
            return key.toLowerCase(Locale.ROOT).contains(q)
                    || val.toLowerCase(Locale.ROOT).contains(q);
        }

        private void hookRowForSearchFilter(ObservableList<String> r) {
            if (r == null) {
                return;
            }
            Runnable ping = () -> Platform.runLater(this::applySearchPredicate);
            r.addListener((ListChangeListener<String>) c -> ping.run());
        }
    }

    /** リポジトリ {@code code/} 同梱から作業先へ上書き復元するサブタブ。 */
    private final class RepoRestorePanel {

        private record RestoreRow(String tabLabel, String filename, String workStatus, String repoStatus) {}

        private final VBox root;
        private final Label workDirLabel = new Label();
        private final TableView<RestoreRow> table = new TableView<>();
        private final ObservableList<RestoreRow> rows = FXCollections.observableArrayList();

        RepoRestorePanel() {
            workDirLabel.setWrapText(true);
            Label intro =
                    new Label(
                            "サマリ Excel と同一フォルダの材料・製品種類テーブルを、リポジトリ code/ 同梱の内容で上書きします。"
                                    + " 編集中の未保存変更は失われます。");
            intro.setWrapText(true);

            TableColumn<RestoreRow, String> colLabel = new TableColumn<>("編集タブ");
            colLabel.setPrefWidth(180);
            colLabel.setCellValueFactory(cd -> new javafx.beans.property.SimpleStringProperty(cd.getValue().tabLabel()));

            TableColumn<RestoreRow, String> colFile = new TableColumn<>("ファイル名");
            colFile.setPrefWidth(240);
            colFile.setCellValueFactory(cd -> new javafx.beans.property.SimpleStringProperty(cd.getValue().filename()));

            TableColumn<RestoreRow, String> colWork = new TableColumn<>("作業先");
            colWork.setPrefWidth(80);
            colWork.setCellValueFactory(cd -> new javafx.beans.property.SimpleStringProperty(cd.getValue().workStatus()));

            TableColumn<RestoreRow, String> colRepo = new TableColumn<>("リポジトリ源");
            colRepo.setPrefWidth(100);
            colRepo.setCellValueFactory(cd -> new javafx.beans.property.SimpleStringProperty(cd.getValue().repoStatus()));

            table.getColumns().addAll(List.of(colLabel, colFile, colWork, colRepo));
            table.setItems(rows);
            table.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
            table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);

            Button refresh = new Button("状態更新");
            refresh.setOnAction(e -> refreshStatus());
            Button overwriteSel = new Button("選択を上書き");
            overwriteSel.setOnAction(e -> overwriteSelected());
            Button overwriteAll = new Button("すべて上書き");
            overwriteAll.setOnAction(e -> overwriteAll());
            HBox bar = new HBox(8, refresh, overwriteSel, overwriteAll);
            bar.setAlignment(Pos.CENTER_LEFT);

            root = new VBox(8, intro, workDirLabel, bar, table);
            VBox.setVgrow(table, Priority.ALWAYS);
            root.setPadding(new Insets(0, 0, 4, 0));
        }

        VBox root() {
            return root;
        }

        void refreshStatus() {
            Map<String, String> ui = uiEnv();
            Path workDir =
                    AppPaths.dispatchLookupTablePath(ui, AppPaths.DISPATCH_LOOKUP_PRODUCT_THICK)
                            .getParent();
            workDirLabel.setText("作業先フォルダ: " + (workDir != null ? workDir : "（未設定）"));
            rows.clear();
            for (FileSpec spec : FILES) {
                Path target = AppPaths.dispatchLookupTablePath(ui, spec.relativePath());
                Optional<Path> source =
                        AppPaths.resolveBundledDispatchLookupTableSourceInRepo(ui, spec.relativePath());
                String workStatus = Files.isRegularFile(target) ? "あり" : "なし";
                String repoStatus = source.isPresent() ? "あり" : "なし";
                rows.add(new RestoreRow(spec.tabLabel(), spec.relativePath(), workStatus, repoStatus));
            }
            table.refresh();
        }

        private void overwriteSelected() {
            var sel = table.getSelectionModel().getSelectedItems();
            if (sel == null || sel.isEmpty()) {
                if (shell != null) {
                    shell.showErrorDialog("上書き", "上書きするファイルを選択してください。");
                }
                return;
            }
            if (!confirmOverwrite(sel.size() == 1 ? sel.getFirst().filename() : sel.size() + " 件")) {
                return;
            }
            applyOverwrite(sel.stream().map(RestoreRow::filename).toList());
        }

        private void overwriteAll() {
            if (!confirmOverwrite("すべて（6 件）")) {
                return;
            }
            applyOverwrite(AppPaths.dispatchLookupTableFilenames());
        }

        private boolean confirmOverwrite(String targetDescription) {
            Alert confirm = new Alert(AlertType.CONFIRMATION);
            confirm.setTitle("上書きの確認");
            confirm.setHeaderText(null);
            confirm.setContentText(
                    targetDescription
                            + " をリポジトリ同梱の内容で上書きします。\n"
                            + "作業先の未保存変更は失われます。続行しますか？");
            if (shell != null && shell.primaryStageForDialogs() != null) {
                confirm.initOwner(shell.primaryStageForDialogs());
            }
            Optional<ButtonType> ans = confirm.showAndWait();
            return ans.isPresent() && ans.get() == ButtonType.OK;
        }

        private void applyOverwrite(List<String> filenames) {
            Map<String, String> ui = uiEnv();
            int ok = 0;
            int ng = 0;
            for (String filename : filenames) {
                DispatchLookupTableOverwriteResult r =
                        AppPaths.overwriteDispatchLookupTableFromRepo(ui, filename);
                if (r.success()) {
                    ok++;
                    logLine(
                            "[code-lookup-restore] 上書き: "
                                    + r.targetPath()
                                    + " ← "
                                    + r.sourcePath());
                    FilePanel panel = panelByFilename.get(filename);
                    if (panel != null) {
                        panel.reloadFromDisk();
                    }
                } else {
                    ng++;
                    logLine(
                            "[code-lookup-restore] 失敗: "
                                    + filename
                                    + " → "
                                    + r.message());
                }
            }
            refreshStatus();
            scheduleInvalidatePlanInputRollUnitHighlightCache();
            if (ng > 0 && shell != null) {
                shell.showErrorDialog(
                        "上書き",
                        "成功 " + ok + " 件、失敗 " + ng + " 件。詳細は実行ログを確認してください。");
            } else if (ok > 0) {
                logLine("[code-lookup-restore] 完了: " + ok + " 件を上書きしました。");
            }
        }
    }
}
