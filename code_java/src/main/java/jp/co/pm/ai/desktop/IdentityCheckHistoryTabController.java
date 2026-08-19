package jp.co.pm.ai.desktop;

import java.awt.Desktop;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import javafx.application.Platform;
import javafx.beans.property.ReadOnlyStringWrapper;
import javafx.collections.FXCollections;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.OperatorUserPaths;
import jp.co.pm.ai.desktop.io.AladdinEntryDispatchPlanIdentityCheck;
import jp.co.pm.ai.desktop.io.IdentityCheckHistoryStore;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.ui.AladdinEntryIdentityCheckResultDialog;

/**
 * 同一化チェック履歴（操作者別・Excel＋加工計画 JSON）を閲覧する。
 */
public final class IdentityCheckHistoryTabController {

    private static final int JSON_DIALOG_MAX_COLUMNS = 40;

    private final AtomicInteger identityCheckGeneration = new AtomicInteger();

    private MainShellController shell;

    private boolean refreshing;

    @FXML
    private ComboBox<String> operatorCombo;

    @FXML
    private Button openExcelButton;

    @FXML
    private Button showPlanJsonButton;

    @FXML
    private Button runIdentityCheckButton;

    @FXML
    private Label pathLabel;

    @FXML
    private Label statusLabel;

    @FXML
    private Label detailLabel;

    @FXML
    private TableView<IdentityCheckHistoryStore.SnapshotRef> historyTable;

    @FXML
    private TableColumn<IdentityCheckHistoryStore.SnapshotRef, String> tsColumn;

    @FXML
    private TableColumn<IdentityCheckHistoryStore.SnapshotRef, String> resultColumn;

    @FXML
    private TableColumn<IdentityCheckHistoryStore.SnapshotRef, String> diffCountColumn;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        refresh();
    }

    /** メインタブ選択時に一覧を再読込する。 */
    public void onMainShellTabSelected() {
        refresh();
    }

    @FXML
    private void initialize() {
        tsColumn.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(formatTs(c.getValue().meta().savedAt())));
        resultColumn.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(resultLabel(c.getValue().meta().result())));
        diffCountColumn.setCellValueFactory(
                c ->
                        new ReadOnlyStringWrapper(
                                Integer.toString(Math.max(0, c.getValue().meta().diffCount()))));
        historyTable.setPlaceholder(new Label("この操作者の履歴はありません。"));
        historyTable
                .getSelectionModel()
                .selectedItemProperty()
                .addListener((obs, a, b) -> onSelectionChanged(b));
        operatorCombo
                .valueProperty()
                .addListener(
                        (obs, a, b) -> {
                            if (refreshing || b == null) {
                                return;
                            }
                            loadRows(b);
                        });
        updateActionButtons(null);
    }

    @FXML
    private void onRefreshAction() {
        refresh();
    }

    @FXML
    private void onOpenExcelAction() {
        IdentityCheckHistoryStore.SnapshotRef sel = historyTable.getSelectionModel().getSelectedItem();
        if (sel == null) {
            statusLabel.setText("行を選択してください。");
            return;
        }
        Path excel = sel.dir().resolve(IdentityCheckHistoryStore.EXCEL_FILE);
        if (!isSafeHistoryFile(excel)) {
            warnUser("Excel を開けません", "履歴フォルダ外、またはファイルがありません:\n" + excel);
            return;
        }
        openPathInBackground(excel, "Excel");
    }

    @FXML
    private void onShowPlanJsonAction() {
        IdentityCheckHistoryStore.SnapshotRef sel = historyTable.getSelectionModel().getSelectedItem();
        if (sel == null) {
            statusLabel.setText("行を選択してください。");
            return;
        }
        Path json = sel.dir().resolve(IdentityCheckHistoryStore.PLAN_JSON_FILE);
        if (!isSafeHistoryFile(json)) {
            warnUser("JSON を開けません", "履歴フォルダ外、またはファイルがありません:\n" + json);
            return;
        }
        String titleSuffix =
                sel.dir().getFileName() != null ? sel.dir().getFileName().toString() : "";
        statusLabel.setText("加工計画 JSON を読み込み中…");
        Thread t =
                new Thread(
                        () -> {
                            try {
                                JsonTableIo.ArrayTable table = JsonTableIo.loadArrayTable(json);
                                Platform.runLater(
                                        () -> {
                                            statusLabel.setText("加工計画 JSON を表示します。");
                                            showPlanTableDialog(table, titleSuffix);
                                        });
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () -> {
                                            statusLabel.setText("JSON を読めませんでした。");
                                            warnUser(
                                                    "JSON を読めませんでした",
                                                    ex.getMessage() != null
                                                            ? ex.getMessage()
                                                            : ex.toString());
                                        });
                            }
                        },
                        "identity-check-history-json");
        t.setDaemon(true);
        t.start();
    }

    @FXML
    private void onRunIdentityCheckAction() {
        IdentityCheckHistoryStore.SnapshotRef sel = historyTable.getSelectionModel().getSelectedItem();
        if (sel == null) {
            statusLabel.setText("行を選択してください。");
            return;
        }
        if (shell == null) {
            return;
        }
        Path excel = sel.dir().resolve(IdentityCheckHistoryStore.EXCEL_FILE);
        Path planJson = sel.dir().resolve(IdentityCheckHistoryStore.PLAN_JSON_FILE);
        if (!isSafeHistoryFile(excel) || !isSafeHistoryFile(planJson)) {
            warnUser(
                    "同一化チェック",
                    "履歴フォルダ外、またはファイルがありません:\nExcel: " + excel + "\nJSON: " + planJson);
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        int generation = identityCheckGeneration.incrementAndGet();
        statusLabel.setText("同一化チェック中…");
        Thread worker =
                new Thread(
                        () -> {
                            AladdinEntryDispatchPlanIdentityCheck.Result result =
                                    AladdinEntryDispatchPlanIdentityCheck.evaluateSnapshot(
                                            ui, sel.dir(), false);
                            Platform.runLater(
                                    () -> {
                                        if (generation != identityCheckGeneration.get()) {
                                            return;
                                        }
                                        finishIdentityCheck(result);
                                    });
                        },
                        "identity-check-history-rerun");
        worker.setDaemon(true);
        worker.start();
    }

    private void finishIdentityCheck(AladdinEntryDispatchPlanIdentityCheck.Result result) {
        if (result.excelPath().isPresent()) {
            shell.appendLog("[aladdin-identity-check] excel=" + result.excelPath().get());
        }
        if (result.planSourcePath().isPresent()) {
            shell.appendLog("[aladdin-identity-check] plan=" + result.planSourcePath().get());
        }
        shell.appendLog("[aladdin-identity-check] 履歴再実行 " + result.message());
        String identityResult = result.error() ? "error" : (result.identical() ? "ok" : "mismatch");
        String identityDetail =
                result.error()
                        ? (result.message() != null ? result.message() : "比較失敗")
                        : (result.identical()
                                ? "同一"
                                : (result.badgeText() != null ? result.badgeText() : "差異"));
        shell.recordOperatorAction("identity_check", identityResult, "履歴再実行 " + identityDetail);
        if (result.error()) {
            statusLabel.setText("同一化チェック失敗");
            shell.showWarningDialog("同一化チェック", result.message());
            return;
        }
        statusLabel.setText(
                result.identical()
                        ? "同一化チェック: 同一"
                        : "同一化チェック: " + (result.badgeText() != null ? result.badgeText() : "差異"));
        if (result.identical()) {
            shell.showInformationDialog("同一化チェック", result.message());
            return;
        }
        Window owner = historyTable.getScene() != null ? historyTable.getScene().getWindow() : null;
        AladdinEntryIdentityCheckResultDialog.show(owner, result);
    }

    private void refresh() {
        if (shell == null) {
            return;
        }
        refreshing = true;
        try {
            Map<String, String> ui = shell.snapshotUiEnv();
            String self = currentOperator(ui);
            List<String> names = new ArrayList<>(IdentityCheckHistoryStore.listOperatorDirNames(ui));
            if (!self.isBlank() && names.stream().noneMatch(n -> n.equalsIgnoreCase(self))) {
                names.add(0, self);
            }
            String selected = operatorCombo.getValue();
            operatorCombo.setItems(FXCollections.observableArrayList(names));
            String next =
                    selected != null && names.stream().anyMatch(n -> n.equalsIgnoreCase(selected))
                            ? selected
                            : pickIgnoreCase(names, self)
                                    .orElse(names.isEmpty() ? null : names.get(0));
            operatorCombo.setValue(next);
            var root = IdentityCheckHistoryStore.resolveRoot(ui);
            pathLabel.setText("保存先: " + root);
            if (!Files.isDirectory(root)) {
                historyTable.setItems(FXCollections.observableArrayList());
                statusLabel.setText("共有の同一化チェック履歴フォルダがありません（未作成または到達不能）。");
                onSelectionChanged(null);
                return;
            }
            if (next != null) {
                loadRows(next);
            } else {
                historyTable.setItems(FXCollections.observableArrayList());
                statusLabel.setText("履歴はありません。");
                onSelectionChanged(null);
            }
        } finally {
            refreshing = false;
        }
    }

    private void loadRows(String operator) {
        Map<String, String> ui = shell.snapshotUiEnv();
        List<IdentityCheckHistoryStore.SnapshotRef> rows =
                IdentityCheckHistoryStore.listNewestFirst(ui, operator);
        historyTable.setItems(FXCollections.observableArrayList(rows));
        statusLabel.setText(rows.isEmpty() ? "この操作者の履歴はありません。" : rows.size() + " 件");
        onSelectionChanged(historyTable.getSelectionModel().getSelectedItem());
    }

    private void onSelectionChanged(IdentityCheckHistoryStore.SnapshotRef sel) {
        updateActionButtons(sel);
        if (detailLabel == null) {
            return;
        }
        if (sel == null) {
            detailLabel.setText("行を選択すると保存時のソースパスなどを表示します。");
            return;
        }
        IdentityCheckHistoryStore.Meta m = sel.meta();
        String folder =
                sel.dir().getFileName() != null ? sel.dir().getFileName().toString() : "";
        detailLabel.setText(
                "フォルダ: "
                        + folder
                        + "\n結果: "
                        + resultLabel(m.result())
                        + " / 差異 "
                        + Math.max(0, m.diffCount())
                        + " 件"
                        + (m.badgeText() != null && !m.badgeText().isBlank()
                                ? "（" + m.badgeText() + "）"
                                : "")
                        + "\nExcel元: "
                        + nullToEmpty(m.excelSourcePath())
                        + "\n加工計画元: "
                        + nullToEmpty(m.planSourcePath()));
    }

    private void updateActionButtons(IdentityCheckHistoryStore.SnapshotRef sel) {
        boolean enable = sel != null;
        if (openExcelButton != null) {
            openExcelButton.setDisable(!enable);
        }
        if (showPlanJsonButton != null) {
            showPlanJsonButton.setDisable(!enable);
        }
        if (runIdentityCheckButton != null) {
            runIdentityCheckButton.setDisable(!enable);
        }
    }

    private void openPathInBackground(Path path, String kind) {
        statusLabel.setText(kind + " を開いています…");
        Thread t =
                new Thread(
                        () -> {
                            try {
                                if (!Desktop.isDesktopSupported()
                                        || !Desktop.getDesktop().isSupported(Desktop.Action.OPEN)) {
                                    Platform.runLater(
                                            () -> {
                                                statusLabel.setText("この環境ではファイルを開けません。");
                                                warnUser("開けません", "この環境では外部アプリ起動に対応していません。");
                                            });
                                    return;
                                }
                                Desktop.getDesktop().open(path.toFile());
                                Platform.runLater(() -> statusLabel.setText(kind + " を開きました。"));
                            } catch (Exception ex) {
                                Platform.runLater(
                                        () -> {
                                            statusLabel.setText(kind + " を開けませんでした。");
                                            warnUser(
                                                    kind + " を開けませんでした",
                                                    ex.getMessage() != null
                                                            ? ex.getMessage()
                                                            : ex.toString());
                                        });
                            }
                        },
                        "identity-check-history-open");
        t.setDaemon(true);
        t.start();
    }

    private void showPlanTableDialog(JsonTableIo.ArrayTable table, String titleSuffix) {
        TableView<List<String>> tv = new TableView<>();
        List<String> cols = table.columns() != null ? table.columns() : List.of();
        int colLimit = Math.min(cols.size(), JSON_DIALOG_MAX_COLUMNS);
        for (int i = 0; i < colLimit; i++) {
            final int colIndex = i;
            TableColumn<List<String>, String> col = new TableColumn<>(cols.get(i));
            col.setPrefWidth(120);
            col.setCellValueFactory(
                    c -> {
                        List<String> row = c.getValue();
                        String v =
                                row != null && colIndex < row.size() && row.get(colIndex) != null
                                        ? row.get(colIndex)
                                        : "";
                        return new ReadOnlyStringWrapper(v);
                    });
            tv.getColumns().add(col);
        }
        tv.setItems(
                FXCollections.observableArrayList(
                        table.rows() != null ? table.rows() : List.of()));
        tv.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        VBox.setVgrow(tv, Priority.ALWAYS);

        Label note =
                new Label(
                        cols.size() > JSON_DIALOG_MAX_COLUMNS
                                ? "列が多いため先頭 " + JSON_DIALOG_MAX_COLUMNS + " 列のみ表示しています（全 "
                                        + cols.size()
                                        + " 列）。"
                                : "");
        note.setWrapText(true);

        VBox content = new VBox(8, note, tv);
        content.setPadding(new Insets(8));
        VBox.setVgrow(tv, Priority.ALWAYS);

        ButtonType closeType = new ButtonType("閉じる", ButtonBar.ButtonData.CANCEL_CLOSE);
        Dialog<Void> dialog = new Dialog<>();
        dialog.setTitle("加工計画 JSON — " + nullToEmpty(titleSuffix));
        dialog.getDialogPane().getButtonTypes().add(closeType);
        dialog.getDialogPane().setContent(content);
        dialog.setResizable(true);
        dialog.initModality(Modality.WINDOW_MODAL);
        Window owner = historyTable.getScene() != null ? historyTable.getScene().getWindow() : null;
        if (owner != null) {
            dialog.initOwner(owner);
            dialog.setWidth(Math.max(640, owner.getWidth() * 0.8));
            dialog.setHeight(Math.max(420, owner.getHeight() * 0.75));
        } else {
            dialog.setWidth(900);
            dialog.setHeight(560);
        }
        dialog.showAndWait();
    }

    private boolean isSafeHistoryFile(Path file) {
        if (file == null || !Files.isRegularFile(file) || shell == null) {
            return false;
        }
        Path root = IdentityCheckHistoryStore.resolveRoot(shell.snapshotUiEnv()).toAbsolutePath().normalize();
        Path abs = file.toAbsolutePath().normalize();
        return abs.startsWith(root);
    }

    private void warnUser(String title, String message) {
        Alert a = new Alert(Alert.AlertType.WARNING);
        a.setTitle(title);
        a.setHeaderText(null);
        a.setContentText(message);
        Window owner = historyTable.getScene() != null ? historyTable.getScene().getWindow() : null;
        if (owner instanceof Stage stage) {
            a.initOwner(stage);
        }
        a.show();
    }

    private static String currentOperator(Map<String, String> ui) {
        String session = FactoryOperatorUserStore.sessionOperatorName();
        if (!session.isBlank()) {
            return OperatorUserPaths.sanitizeOperatorDirName(session);
        }
        return OperatorUserPaths.sanitizeOperatorDirName(OperatorUserPaths.resolveOperatorUser(ui));
    }

    private static java.util.Optional<String> pickIgnoreCase(List<String> names, String target) {
        if (target == null || target.isBlank() || names == null) {
            return java.util.Optional.empty();
        }
        return names.stream().filter(n -> n.equalsIgnoreCase(target)).findFirst();
    }

    static String formatTs(String ts) {
        if (ts == null || ts.isBlank()) {
            return "";
        }
        try {
            return OffsetDateTime.parse(ts)
                    .format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        } catch (DateTimeParseException ex) {
            return ts;
        }
    }

    static String resultLabel(String result) {
        if (result == null) {
            return "";
        }
        return switch (result) {
            case "ok" -> "同一";
            case "mismatch" -> "差異";
            case "error" -> "失敗";
            default -> result;
        };
    }

    private static String nullToEmpty(String s) {
        return s != null ? s : "";
    }
}
