package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.ui.TabularCellHighlight;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.ui.TableViewColumnSettingsStrip;

/** Stage1 task-input preview table; layout {@code Stage1PreviewTab.fxml}. */
public final class Stage1PreviewTabController {

    public static final String DEFAULT_STAGE1_PREVIEW_SHEET = AppPaths.STAGE1_TASK_INPUT_PREVIEW_SHEET;

    private Stage ownerStage;

    private MainShellController shell;

    @FXML
    private Label hintLabel;

    @FXML
    private TextField pathField;

    @FXML
    private TextField sheetField;

    @FXML
    private Button fromEnvButton;

    @FXML
    private Button browseButton;

    @FXML
    private Button loadButton;

    @FXML
    private HBox columnStripHost;

    @FXML
    private TextField colWidthField;

    @FXML
    private TableView<ObservableList<String>> table;

    private final List<String> headersRef = new ArrayList<>();
    private ObservableList<ObservableList<String>> rows;
    private final AtomicBoolean suppressColumnOrderPersistence = new AtomicBoolean(false);
    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
            new AtomicReference<>(List.of());

    @FXML
    private void initialize() {
        pathField.setPromptText(
                "output/"
                        + AppPaths.STAGE1_TASK_INPUT_PREVIEW_FILENAME
                        + " (\u554f\u5408\u305b\u53d6\u8fbc\u5f8c\u30fb\u30bf\u30b9\u30af\u4e00\u89a7\u5316\u524d)");
        sheetField.setText(DEFAULT_STAGE1_PREVIEW_SHEET);
        sheetField.setPromptText("Excel sheet name");
        colWidthField.setText("112");

        hintLabel.setText(buildHintText());

        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        table.setEditable(false);
        table.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        rows = FXCollections.observableArrayList();
        table.setItems(rows);
    }

    private static String buildHintText() {
        return "\u554f\u5408\u305b xlsx \u3092\u8aad\u307f\u8fbc\u307f\u3001\u30d8\u30c3\u30c0\u30fc\u884c\u3068"
                + "\u5217\u540d\u3092\u6574\u3048\u305f\u76f4\u5f8c\uff08\u4f9d\u983cNO"
                + "\u304c\u3042\u308b\u884c\u306e\u307f\uff09\u3002"
                + " \u30de\u30b9\u30bf\u30fb\u914d\u53f0\u8a66\u884c\u9806\u4ed8\u4e0e\u524d\u306e"
                + " stage1_task_input_table.xlsx"
                + " \u30b7\u30fc\u30c8\u300c"
                + DEFAULT_STAGE1_PREVIEW_SHEET
                + "\u300d\u3092\u8868\u793a\u3057\u307e\u3059\u3002";
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        ownerStage = shell.getPrimaryStage();

        columnStripHost.getChildren().setAll(TableViewColumnSettingsStrip.create(table, this::applyDynamicColumnWidths, false));

        shell.acceptReloadAfterStage1Preview(
                () -> {
                    fillPathFromEnv();
                    sheetField.setText(DEFAULT_STAGE1_PREVIEW_SHEET);
                    loadButton.fire();
                });

        TableColumnOrderPersistence.installColumnLayoutWatcher(
                table,
                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW,
                suppressColumnOrderPersistence::get);

        javafx.application.Platform.runLater(
                () -> {
                    fillPathFromEnv();
                    if (!pathField.getText().isBlank()) {
                        loadButton.fire();
                    }
                });
    }

    @FXML
    private void onFromEnvButtonAction() {
        fillPathFromEnv();
    }

    @FXML
    private void onBrowseButtonAction() {
        FileChooser ch = new FileChooser();
        ch.setTitle(AppPaths.STAGE1_TASK_INPUT_PREVIEW_FILENAME);
        ch.getExtensionFilters()
                .addAll(
                        new FileChooser.ExtensionFilter("Excel", "*.xlsx", "*.xlsm"),
                        new FileChooser.ExtensionFilter("All", "*.*"));
        String cur = pathField.getText();
        if (cur != null && !cur.isBlank()) {
            try {
                Path p = Path.of(cur.trim());
                if (java.nio.file.Files.isRegularFile(p) && p.getParent() != null) {
                    ch.setInitialDirectory(p.getParent().toFile());
                }
            } catch (Exception ignored) {
            }
        }
        var f = ch.showOpenDialog(ownerStage);
        if (f != null) {
            pathField.setText(f.getAbsolutePath());
        }
    }

    @FXML
    private void onLoadButtonAction() {
        if (pathField.getText().isBlank()) {
            fillPathFromEnv();
        }
        Path path = Path.of(pathField.getText().trim());
        if (!java.nio.file.Files.isRegularFile(path)) {
            shell.appendLog("[stage1-preview] file not found: " + path);
            return;
        }
        String sheet = sheetField.getText().trim();
        if (sheet.isEmpty()) {
            sheet = DEFAULT_STAGE1_PREVIEW_SHEET;
        }
        try {
            PlanInputTabularIo.TabularSheet sh = PlanInputTabularIo.read(path, sheet);
            headersRef.clear();
            headersRef.addAll(sh.headers());
            rows.clear();
            for (List<String> line : sh.rows()) {
                ObservableList<String> r = FXCollections.observableArrayList(line);
                while (r.size() < headersRef.size()) {
                    r.add("");
                }
                while (r.size() > headersRef.size()) {
                    r.remove(r.size() - 1);
                }
                rows.add(r);
            }
            List<TableColumnOrderPersistence.ColumnSpec> lay =
                    TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.STAGE1_PREVIEW);
            persistedLayout.set(lay);
            TableColumnOrderPersistence.applyLogicalColumnOrder(
                    headersRef,
                    rows,
                    lay.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList());
            applyLoaded();
            shell.appendLog(
                    "[stage1-preview] loaded rows="
                            + rows.size()
                            + " cols="
                            + headersRef.size()
                            + " path="
                            + path);
        } catch (Exception ex) {
            shell.appendLog("[stage1-preview] load error: " + ex.getMessage());
        }
    }

    private void applyDynamicColumnWidths() {
        double w = 112;
        try {
            w = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
        } catch (NumberFormatException ignored) {
        }
        for (TableColumn<ObservableList<String>, ?> c : table.getColumns()) {
            c.setPrefWidth(w);
        }
    }

    private void rebuildColumns() {
        suppressColumnOrderPersistence.set(true);
        try {
            double colW = 112;
            try {
                colW = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
            } catch (NumberFormatException ignored) {
            }
            List<Double> widths =
                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                            headersRef, persistedLayout.get(), colW);
            table.getColumns().clear();
            for (int i = 0; i < headersRef.size(); i++) {
                final int idx = i;
                double prefW = i < widths.size() ? widths.get(i) : colW;
                String title =
                        !headersRef.get(i).isBlank()
                                ? headersRef.get(i)
                                : ("\u5217" + (i + 1));
                TableColumn<ObservableList<String>, String> col = new TableColumn<>(title);
                col.setEditable(false);
                col.setCellValueFactory(
                        cd -> {
                            ObservableList<String> row = cd.getValue();
                            String v =
                                    row != null && idx < row.size()
                                            ? row.get(idx)
                                            : "";
                            return new javafx.beans.property.SimpleStringProperty(v);
                        });
                col.setCellFactory(TabularCellHighlight.stage1DateHighlightCellFactory(title));
                col.setPrefWidth(prefW);
                table.getColumns().add(col);
            }
        } finally {
            suppressColumnOrderPersistence.set(false);
        }
    }

    private void applyLoaded() {
        rebuildColumns();
        table.refresh();
    }

    private void fillPathFromEnv() {
        Map<String, String> env = shell.snapshotUiEnv();
        if (env == null) {
            return;
        }
        pathField.setText(AppPaths.defaultStage1TaskInputPreviewPath(env).toString());
    }
}
