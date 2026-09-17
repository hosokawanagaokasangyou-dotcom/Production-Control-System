package jp.co.pm.ai.desktop.developer;

import java.nio.file.Files;
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
import javafx.scene.control.TextArea;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.OperatorUserPaths;
import jp.co.pm.ai.desktop.config.RuntimeErrorLogAnalyzer;

/**
 * remote_log を解析して実行時エラー行を一覧する。
 */
public final class DeveloperRuntimeErrorTabController {

    static final String ALL_OPERATORS = "（全員）";

    private MainShellController shell;
    private List<RuntimeErrorLogAnalyzer.Row> allRows = List.of();

    @FXML private ComboBox<String> operatorCombo;
    @FXML private Label pathLabel;
    @FXML private Label statusLabel;
    @FXML private TableView<RuntimeErrorLogAnalyzer.Row> errorTable;
    @FXML private TableColumn<RuntimeErrorLogAnalyzer.Row, String> tsColumn;
    @FXML private TableColumn<RuntimeErrorLogAnalyzer.Row, String> operatorColumn;
    @FXML private TableColumn<RuntimeErrorLogAnalyzer.Row, String> versionColumn;
    @FXML private TableColumn<RuntimeErrorLogAnalyzer.Row, String> factoryColumn;
    @FXML private TableColumn<RuntimeErrorLogAnalyzer.Row, String> eventColumn;
    @FXML private TableColumn<RuntimeErrorLogAnalyzer.Row, String> severityColumn;
    @FXML private TableColumn<RuntimeErrorLogAnalyzer.Row, String> excerptColumn;
    @FXML private TextArea detailArea;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        refresh();
    }

    public void onMainShellTabSelected() {
        refresh();
    }

    @FXML
    private void initialize() {
        tsColumn.setCellValueFactory(c -> new ReadOnlyStringWrapper(nz(c.getValue().ts())));
        operatorColumn.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(nz(c.getValue().operator())));
        versionColumn.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(nz(c.getValue().appVersion())));
        factoryColumn.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(nz(c.getValue().factory())));
        eventColumn.setCellValueFactory(c -> new ReadOnlyStringWrapper(nz(c.getValue().eventId())));
        severityColumn.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(severityLabel(c.getValue().severity())));
        excerptColumn.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(nz(c.getValue().excerpt())));
        operatorCombo
                .valueProperty()
                .addListener(
                        (obs, a, b) -> applyFilter());
        errorTable
                .getSelectionModel()
                .selectedItemProperty()
                .addListener((obs, a, b) -> showDetail(b));
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
        var root = AppPaths.resolveRemoteLogRoot(ui);
        pathLabel.setText("解析先: " + root);
        if (!Files.isDirectory(root)) {
            allRows = List.of();
            errorTable.setItems(FXCollections.observableArrayList());
            statusLabel.setText("共有の remote_log がありません（未作成または到達不能）。");
            return;
        }
        allRows = RuntimeErrorLogAnalyzer.scan(ui);
        String self = currentOperator(ui);
        List<String> names = new ArrayList<>();
        names.add(ALL_OPERATORS);
        names.addAll(RuntimeErrorLogAnalyzer.listOperators(ui));
        if (!self.isBlank() && names.stream().noneMatch(n -> n.equalsIgnoreCase(self))) {
            names.add(self);
        }
        String selected = operatorCombo.getValue();
        operatorCombo.setItems(FXCollections.observableArrayList(names));
        String next =
                selected != null && names.stream().anyMatch(n -> n.equalsIgnoreCase(selected))
                        ? selected
                        : ALL_OPERATORS;
        operatorCombo.setValue(next);
        applyFilter();
    }

    private void applyFilter() {
        String selected = operatorCombo.getValue();
        List<RuntimeErrorLogAnalyzer.Row> shown = new ArrayList<>();
        for (RuntimeErrorLogAnalyzer.Row row : allRows) {
            if (selected == null
                    || ALL_OPERATORS.equals(selected)
                    || selected.equalsIgnoreCase(row.operator())) {
                shown.add(row);
            }
        }
        errorTable.setItems(FXCollections.observableArrayList(shown));
        statusLabel.setText(shown.isEmpty() ? "該当する実行時エラーはありません。" : shown.size() + " 件");
        if (shown.isEmpty()) {
            detailArea.setText("");
        }
    }

    private void showDetail(RuntimeErrorLogAnalyzer.Row row) {
        if (detailArea == null) {
            return;
        }
        if (row == null) {
            detailArea.setText("");
            return;
        }
        detailArea.setText(
                "操作者: "
                        + nz(row.operator())
                        + "\n版: "
                        + nz(row.appVersion())
                        + "\n工場: "
                        + nz(row.factory())
                        + "\nホスト: "
                        + nz(row.host())
                        + "\nWindowsユーザー: "
                        + nz(row.osUser())
                        + "\n種別: "
                        + nz(row.eventId())
                        + "\n重要度: "
                        + severityLabel(row.severity())
                        + "\nファイル: "
                        + nz(row.sourcePath())
                        + "\n\n"
                        + nz(row.excerpt()));
    }

    private static String currentOperator(Map<String, String> ui) {
        String session = FactoryOperatorUserStore.sessionOperatorName();
        if (!session.isBlank()) {
            return OperatorUserPaths.sanitizeOperatorDirName(session);
        }
        return OperatorUserPaths.sanitizeOperatorDirName(OperatorUserPaths.resolveOperatorUser(ui));
    }

    static String severityLabel(String severity) {
        if ("error".equals(severity)) {
            return "エラー";
        }
        if ("warn".equals(severity)) {
            return "警告";
        }
        return severity != null ? severity : "";
    }

    private static String nz(String s) {
        return s != null ? s : "";
    }
}
