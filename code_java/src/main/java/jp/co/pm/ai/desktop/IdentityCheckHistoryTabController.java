package jp.co.pm.ai.desktop;

import java.awt.Desktop;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javafx.beans.property.ReadOnlyStringWrapper;
import javafx.collections.FXCollections;
import javafx.fxml.FXML;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.ButtonType;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.OperatorUserPaths;
import jp.co.pm.ai.desktop.io.IdentityCheckHistoryStore;
import jp.co.pm.ai.desktop.io.JsonTableIo;

/**
 * 同一化チェック履歴（操作者別・Excel＋加工計画 JSON）を閲覧する。
 */
public final class IdentityCheckHistoryTabController {

    private MainShellController shell;

    @FXML
    private ComboBox<String> operatorCombo;

    @FXML
    private Label pathLabel;

    @FXML
    private Label statusLabel;

    @FXML
    private TableView<IdentityCheckHistoryStore.SnapshotRef> historyTable;

    @FXML
    private TableColumn<IdentityCheckHistoryStore.SnapshotRef, String> tsColumn;

    @FXML
    private TableColumn<IdentityCheckHistoryStore.SnapshotRef, String> resultColumn;

    @FXML
    private TableColumn<IdentityCheckHistoryStore.SnapshotRef, String> diffCountColumn;

    @FXML
    private TableColumn<IdentityCheckHistoryStore.SnapshotRef, String> folderColumn;

    @FXML
    private TableColumn<IdentityCheckHistoryStore.SnapshotRef, String> badgeColumn;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
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
        folderColumn.setCellValueFactory(
                c ->
                        new ReadOnlyStringWrapper(
                                c.getValue().dir().getFileName() != null
                                        ? c.getValue().dir().getFileName().toString()
                                        : ""));
        badgeColumn.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(nullToEmpty(c.getValue().meta().badgeText())));
        operatorCombo
                .valueProperty()
                .addListener(
                        (obs, a, b) -> {
                            if (b != null) {
                                loadRows(b);
                            }
                        });
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
        if (!Files.isRegularFile(excel)) {
            statusLabel.setText("Excel ファイルがありません: " + excel);
            return;
        }
        try {
            if (!Desktop.isDesktopSupported() || !Desktop.getDesktop().isSupported(Desktop.Action.OPEN)) {
                statusLabel.setText("この環境ではファイルを開けません。");
                return;
            }
            Desktop.getDesktop().open(excel.toFile());
            statusLabel.setText("Excel を開きました。");
        } catch (IOException ex) {
            statusLabel.setText(
                    "Excel を開けませんでした: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
    }

    @FXML
    private void onShowPlanJsonAction() {
        IdentityCheckHistoryStore.SnapshotRef sel = historyTable.getSelectionModel().getSelectedItem();
        if (sel == null) {
            statusLabel.setText("行を選択してください。");
            return;
        }
        Path json = sel.dir().resolve(IdentityCheckHistoryStore.PLAN_JSON_FILE);
        if (!Files.isRegularFile(json)) {
            statusLabel.setText("加工計画 JSON がありません: " + json);
            return;
        }
        try {
            JsonTableIo.ArrayTable table = JsonTableIo.loadArrayTable(json);
            showPlanTableDialog(table, sel.dir().getFileName().toString());
        } catch (IOException ex) {
            statusLabel.setText(
                    "JSON を読めませんでした: "
                            + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
    }

    private void refresh() {
        if (shell == null) {
            return;
        }
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
                        : (names.contains(self) ? self : (names.isEmpty() ? null : names.get(0)));
        operatorCombo.setValue(next);
        var root = IdentityCheckHistoryStore.resolveRoot(ui);
        pathLabel.setText("保存先: " + root);
        if (!Files.isDirectory(root)) {
            historyTable.setItems(FXCollections.observableArrayList());
            statusLabel.setText("共有の同一化チェック履歴フォルダがありません（未作成または到達不能）。");
            return;
        }
        if (next != null) {
            loadRows(next);
        } else {
            historyTable.setItems(FXCollections.observableArrayList());
            statusLabel.setText("履歴はありません。");
        }
    }

    private void loadRows(String operator) {
        Map<String, String> ui = shell.snapshotUiEnv();
        List<IdentityCheckHistoryStore.SnapshotRef> rows =
                IdentityCheckHistoryStore.listNewestFirst(ui, operator);
        historyTable.setItems(FXCollections.observableArrayList(rows));
        statusLabel.setText(rows.isEmpty() ? "この操作者の履歴はありません。" : rows.size() + " 件");
    }

    private void showPlanTableDialog(JsonTableIo.ArrayTable table, String titleSuffix) {
        TableView<List<String>> tv = new TableView<>();
        List<String> cols = table.columns() != null ? table.columns() : List.of();
        for (int i = 0; i < cols.size(); i++) {
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
        tv.setPrefSize(900, 480);

        Dialog<Void> dialog = new Dialog<>();
        dialog.setTitle("加工計画 JSON — " + nullToEmpty(titleSuffix));
        dialog.getDialogPane().getButtonTypes().add(ButtonType.CLOSE);
        dialog.getDialogPane().setContent(new VBox(8, tv));
        dialog.initModality(Modality.WINDOW_MODAL);
        Window owner = historyTable.getScene() != null ? historyTable.getScene().getWindow() : null;
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.showAndWait();
    }

    private static String currentOperator(Map<String, String> ui) {
        String session = FactoryOperatorUserStore.sessionOperatorName();
        if (!session.isBlank()) {
            return OperatorUserPaths.sanitizeOperatorDirName(session);
        }
        return OperatorUserPaths.sanitizeOperatorDirName(OperatorUserPaths.resolveOperatorUser(ui));
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
            default -> result;
        };
    }

    private static String nullToEmpty(String s) {
        return s != null ? s : "";
    }
}
