package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import org.controlsfx.control.spreadsheet.GridBase;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnReorderDialog;
import jp.co.pm.ai.desktop.ui.SpreadsheetColumnSettingsStrip;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;

/** Stage1 task-input preview table; layout {@code Stage1PreviewTab.fxml}. Uses ControlsFX {@link SpreadsheetView}. */
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
    private StackPane spreadsheetHost;

    private final SpreadsheetView spreadsheetView = new SpreadsheetView();

    private final List<String> headersRef = new ArrayList<>();
    private ObservableList<ObservableList<String>> rows;
    private final AtomicBoolean suppressColumnOrderPersistence = new AtomicBoolean(false);
    private final AtomicReference<List<TableColumnOrderPersistence.ColumnSpec>> persistedLayout =
            new AtomicReference<>(List.of());
    private final AtomicInteger headerColumnCount = new AtomicInteger(0);

    @FXML
    private void initialize() {
        pathField.setPromptText(
                "PM_AI_OUTPUT_DIR \u307e\u305f\u306f\u30ea\u30dd\u6839\u76f4\u4e0b/output/"
                        + AppPaths.STAGE1_TASK_INPUT_PREVIEW_FILENAME
                        + " (\u554f\u5408\u305b\u53d6\u8fbc\u5f8c\u30fb\u30bf\u30b9\u30af\u4e00\u89a7\u5316\u524d)");
        sheetField.setText(DEFAULT_STAGE1_PREVIEW_SHEET);
        sheetField.setPromptText("Excel sheet name");
        colWidthField.setText("112");

        hintLabel.setText(buildHintText());

        StackPane.setAlignment(spreadsheetView, Pos.CENTER_LEFT);
        spreadsheetHost.getChildren().add(spreadsheetView);
        VBox.setVgrow(spreadsheetHost, Priority.ALWAYS);

        rows = FXCollections.observableArrayList();
        spreadsheetView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        SpreadsheetThemeBridge.install(spreadsheetView);
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

        columnStripHost
                .getChildren()
                .setAll(
                        SpreadsheetColumnSettingsStrip.create(
                                this::applyDynamicColumnWidths,
                                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW,
                                headerColumnCount,
                                this::onLeadingColumnCountCommitted,
                                this::onReorderColumns));

        shell.acceptReloadAfterStage1Preview(
                () -> {
                    fillPathFromEnv();
                    sheetField.setText(DEFAULT_STAGE1_PREVIEW_SHEET);
                    loadButton.fire();
                });

        TableColumnOrderPersistence.installSpreadsheetColumnLayoutWatcher(
                spreadsheetView,
                TableColumnOrderPersistence.TableId.STAGE1_PREVIEW,
                suppressColumnOrderPersistence::get,
                () -> new ArrayList<>(headersRef));

        javafx.application.Platform.runLater(
                () -> {
                    if (pathField.getText().isBlank()) {
                        fillPathFromEnv();
                    }
                    if (!pathField.getText().isBlank()) {
                        loadButton.fire();
                    }
                });
    }

    private void onLeadingColumnCountCommitted(int n) {
        headerColumnCount.set(n);
        rebuildSpreadsheet();
    }

    private void onReorderColumns() {
        if (headersRef.isEmpty()) {
            shell.appendLog("[stage1-preview] \u5217\u304c\u3042\u308a\u307e\u305b\u3093\uff08\u5148\u306b\u8aad\u307f\u8fbc\u307f\uff09");
            return;
        }
        SpreadsheetColumnReorderDialog.show(ownerStage, new ArrayList<>(headersRef))
                .ifPresent(
                        perm -> {
                            List<String> oldHeaders = new ArrayList<>(headersRef);
                            List<String> titleOrder = perm.stream().map(oldHeaders::get).toList();
                            List<TableColumnOrderPersistence.ColumnSpec> lay = persistedLayout.get();
                            TableColumnOrderPersistence.applyLogicalColumnOrder(
                                    headersRef, rows, titleOrder);
                            double colW = 112;
                            try {
                                colW =
                                        Math.max(
                                                40,
                                                Double.parseDouble(colWidthField.getText().trim()));
                            } catch (NumberFormatException ignored) {
                            }
                            List<Double> widths =
                                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                                            headersRef, lay, colW);
                            List<TableColumnOrderPersistence.ColumnSpec> newLay = new ArrayList<>();
                            for (int i = 0; i < headersRef.size(); i++) {
                                newLay.add(
                                        new TableColumnOrderPersistence.ColumnSpec(
                                                headersRef.get(i), widths.get(i)));
                            }
                            persistedLayout.set(newLay);
                            TableColumnOrderPersistence.saveLayout(
                                    TableColumnOrderPersistence.TableId.STAGE1_PREVIEW, newLay);
                            rebuildSpreadsheet();
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
        for (var c : spreadsheetView.getColumns()) {
            c.setPrefWidth(w);
        }
    }

    private void rebuildSpreadsheet() {
        if (headersRef.isEmpty()) {
            spreadsheetView.setGrid(new GridBase(0, 0));
            return;
        }
        suppressColumnOrderPersistence.set(true);
        try {
            double colW = 112;
            try {
                colW = Math.max(40, Double.parseDouble(colWidthField.getText().trim()));
            } catch (NumberFormatException ignored) {
            }
            final List<Double> widths =
                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                            headersRef, persistedLayout.get(), colW);
            final double widthDefault = colW;

            GridBase grid = SpreadsheetTabularSupport.buildStage1PreviewGrid(headersRef, rows);
            spreadsheetView.setGrid(grid);

            javafx.application.Platform.runLater(
                    () -> {
                        SpreadsheetTabularSupport.applyColumnWidths(spreadsheetView, widths, widthDefault);
                        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicy(spreadsheetView);
                        SpreadsheetTabularSupport.applyFixedLeadingColumns(
                                spreadsheetView, headerColumnCount.get());
                        SpreadsheetTabularSupport.applyColumnFilters(spreadsheetView);
                    });
        } finally {
            suppressColumnOrderPersistence.set(false);
        }
    }

    private void applyLoaded() {
        rebuildSpreadsheet();
    }

    private void fillPathFromEnv() {
        Map<String, String> env = shell.snapshotUiEnv();
        if (env == null) {
            return;
        }
        pathField.setText(AppPaths.defaultStage1TaskInputPreviewPath(env).toString());
    }

    String snapshotStage1PreviewPath() {
        return pathField.getText() != null ? pathField.getText().trim() : "";
    }

    String snapshotStage1PreviewSheet() {
        return sheetField.getText() != null ? sheetField.getText().trim() : "";
    }

    void restoreDesktopSessionPaths(String path, String sheet) {
        if (path != null && !path.isBlank()) {
            pathField.setText(path.trim());
        }
        if (sheet != null && !sheet.isBlank()) {
            sheetField.setText(sheet.trim());
        }
    }

    void clearColumnFiltersAndSort() {
        SpreadsheetTabularSupport.clearAllFiltersAndSort(spreadsheetView);
    }
}
