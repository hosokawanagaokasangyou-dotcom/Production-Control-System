package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.time.Instant;
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
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;

import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.OperatorActionLogStore;
import jp.co.pm.ai.desktop.config.OperatorUserPaths;

/**
 * 共有フォルダ上の操作者別操作ログを閲覧する。
 */
public final class OperatorActionLogTabController {

    private MainShellController shell;

    @FXML
    private ComboBox<String> operatorCombo;

    @FXML
    private Label pathLabel;

    @FXML
    private Label statusLabel;

    @FXML
    private TableView<OperatorActionLogStore.Entry> logTable;

    @FXML
    private TableColumn<OperatorActionLogStore.Entry, String> tsColumn;

    @FXML
    private TableColumn<OperatorActionLogStore.Entry, String> actionColumn;

    @FXML
    private TableColumn<OperatorActionLogStore.Entry, String> resultColumn;

    @FXML
    private TableColumn<OperatorActionLogStore.Entry, String> detailColumn;

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
        tsColumn.setCellValueFactory(c -> new ReadOnlyStringWrapper(formatTs(c.getValue().ts())));
        actionColumn.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(actionLabel(c.getValue().action())));
        resultColumn.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(resultLabel(c.getValue().result())));
        detailColumn.setCellValueFactory(c -> new ReadOnlyStringWrapper(c.getValue().detail()));
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

    private void refresh() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        String self = currentOperator(ui);
        List<String> names = new ArrayList<>(OperatorActionLogStore.listOperators(ui));
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
        var root = OperatorActionLogStore.resolveRoot(ui);
        pathLabel.setText("保存先: " + root);
        if (!Files.isDirectory(root)) {
            logTable.setItems(FXCollections.observableArrayList());
            statusLabel.setText("共有の操作ログフォルダがありません（未作成または到達不能）。");
            return;
        }
        if (next != null) {
            loadRows(next);
        } else {
            logTable.setItems(FXCollections.observableArrayList());
            statusLabel.setText("操作ログはありません。");
        }
    }

    private void loadRows(String operator) {
        Map<String, String> ui = shell.snapshotUiEnv();
        try {
            List<OperatorActionLogStore.Entry> rows =
                    OperatorActionLogStore.readOperator(ui, operator, Instant.now());
            logTable.setItems(FXCollections.observableArrayList(rows));
            statusLabel.setText(rows.isEmpty() ? "この操作者のログはありません。" : rows.size() + " 件");
        } catch (RuntimeException ex) {
            logTable.setItems(FXCollections.observableArrayList());
            statusLabel.setText(
                    "読み込めませんでした: " + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
        }
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
                    .format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm"));
        } catch (DateTimeParseException ex) {
            return ts;
        }
    }

    static String actionLabel(String action) {
        if (action == null) {
            return "";
        }
        return switch (action) {
            case "stage2_complete" -> "段階2完了";
            case "identity_check" -> "同一化チェック";
            case "excel_export" -> "Excel出力";
            case "close_warning" -> "終了警告";
            default -> action;
        };
    }

    static String resultLabel(String result) {
        if (result == null) {
            return "";
        }
        return switch (result) {
            case "ok" -> "成功";
            case "mismatch" -> "差異";
            case "error" -> "失敗";
            case "shown" -> "表示";
            default -> result;
        };
    }
}
