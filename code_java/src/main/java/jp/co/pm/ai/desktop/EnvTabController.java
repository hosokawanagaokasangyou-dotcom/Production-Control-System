package jp.co.pm.ai.desktop;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;

import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.ui.FileChooserForEnvKey;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.ui.TableViewColumnSettingsStrip;

/**
 * Environment variables tab; columns/cell factories in code (FXML layout only). All Japanese UI strings use
 * \\u escapes so source stays portable across editors/OS encodings.
 */
public final class EnvTabController {

    private static final String ENV_HINT_TEXT =
            "OS \u74b0\u5883\u5909\u6570\u306f\u53c2\u7167\u3057\u307e\u305b\u3093\u3002\u3053\u306e\u30bf\u30d6\u3067\u96c6\u7d04\u3002"
                    + " \u521d\u671f\u5024: ui_ref_env_defaults.json + \u30ed\u30b8\u30c3\u30af\u8aac\u660e\u3002"
                    + " \u5b50\u30d7\u30ed\u30bb\u30b9: \u3053\u306e\u8868 + \u30e9\u30f3\u30c1\u30e3\u30fc\u306e TASK_INPUT_WORKBOOK"
                    + "\uff08\u30de\u30af\u30ed\u30d6\u30c3\u30af\u306f\u4efb\u610f\uff09\u2192 PYTHONUTF8 \u6700\u7d42\u56fa\u5b9a\u3002"
                    + " PM_AI_SKIP_WORKBOOK_ENV_SHEET \u304c\u7a7a\u306e\u3068\u304d\u306f 1 \u3068\u3057\u3066"
                    + "\u30de\u30af\u30ed\u300c\u8a2d\u5b9a_\u74b0\u5883\u5909\u6570\u300d\u30b7\u30fc\u30c8\u3092\u8aad\u307e\u306a\u3044\u3002"
                    + " \u30d5\u30a9\u30eb\u30c0\u578b\u306f\u300c\u30d5\u30a9\u30eb\u30c0...\u300d\u3001\u5404\u30d5\u30a1\u30a4\u30eb\u578b\u306f"
                    + "\u5909\u6570\u540d\u306b\u5fdc\u3058\u3066 JSON / Excel / CSV \u306e\u62e1\u5f35\u5b50\u3092\u8868\u793a\u3002";

    @FXML
    private Label hintLabel;

    @FXML
    private HBox columnStripHost;

    @FXML
    private TableView<EnvVarRow> envTable;

    @FXML
    private Button addRowButton;

    @FXML
    private Button delRowButton;

    private Stage ownerStage;
    private ObservableList<EnvVarRow> envRows;

    void bindShell(MainShellController shell) {
        this.ownerStage = shell.getPrimaryStage();
        this.envRows = shell.getEnvRows();
        hintLabel.setText(ENV_HINT_TEXT);
        addRowButton.setText("\u884c\u3092\u8ffd\u52a0");
        delRowButton.setText("\u884c\u3092\u524a\u9664");
        wireTable();
        addRowButton.setOnAction(
                e -> {
                    EnvVarRow r = new EnvVarRow();
                    r.setDescription("");
                    envRows.add(r);
                });
        delRowButton.setOnAction(
                e -> {
                    var sel = envTable.getSelectionModel().getSelectedItems();
                    if (!sel.isEmpty()) {
                        envRows.removeAll(sel);
                    } else if (!envRows.isEmpty()) {
                        envRows.remove(envRows.size() - 1);
                    }
                    if (envRows.isEmpty()) {
                        envRows.add(new EnvVarRow());
                    }
                });
    }

    private void wireTable() {
        envTable.setItems(envRows);
        envTable.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        envTable.setEditable(true);
        envTable.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);

        TableColumn<EnvVarRow, String> nameCol = new TableColumn<>("\u5909\u6570\u540d");
        nameCol.setCellValueFactory(cdf -> cdf.getValue().nameProperty());
        nameCol.setCellFactory(TextFieldTableCell.forTableColumn());
        nameCol.setOnEditCommit(
                e -> {
                    e.getRowValue().setName(e.getNewValue());
                    envTable.refresh();
                });
        nameCol.setPrefWidth(220);

        TableColumn<EnvVarRow, String> valueCol = new TableColumn<>("\u5024");
        valueCol.setCellValueFactory(cdf -> cdf.getValue().valueProperty());
        valueCol.setCellFactory(TextFieldTableCell.forTableColumn());
        valueCol.setOnEditCommit(
                e -> {
                    e.getRowValue().setValue(e.getNewValue());
                    envTable.refresh();
                });

        TableColumn<EnvVarRow, Void> folderCol = new TableColumn<>("\u9078\u629e");
        folderCol.setPrefWidth(120);
        folderCol.setSortable(false);
        folderCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final Button pickFolder =
                                    new Button("\u30d5\u30a9\u30eb\u30c0...");
                            private final Button pickFile =
                                    new Button("\u30d5\u30a1\u30a4\u30eb...");

                            {
                                pickFolder.setOnAction(
                                        ev -> {
                                            EnvVarRow row =
                                                    getTableRow() != null ? getTableRow().getItem() : null;
                                            if (row == null) {
                                                return;
                                            }
                                            DirectoryChooser dc = new DirectoryChooser();
                                            dc.setTitle(
                                                    "\u30d5\u30a9\u30eb\u30c0\u3092\u9078\u629e: "
                                                            + row.getName());
                                            String cur = row.getValue();
                                            if (cur != null && !cur.isBlank()) {
                                                try {
                                                    Path p = Path.of(cur.trim());
                                                    if (Files.isDirectory(p)) {
                                                        dc.setInitialDirectory(p.toFile());
                                                    } else {
                                                        Path par = p.getParent();
                                                        if (par != null && Files.isDirectory(par)) {
                                                            dc.setInitialDirectory(par.toFile());
                                                        }
                                                    }
                                                } catch (Exception ignored) {
                                                    // keep default initial directory
                                                }
                                            }
                                            File f = dc.showDialog(ownerStage);
                                            if (f != null) {
                                                row.setValue(f.getAbsolutePath());
                                                envTable.refresh();
                                            }
                                        });
                                pickFile.setOnAction(
                                        ev -> {
                                            EnvVarRow row =
                                                    getTableRow() != null ? getTableRow().getItem() : null;
                                            if (row == null) {
                                                return;
                                            }
                                            FileChooser fc = new FileChooser();
                                            fc.setTitle(
                                                    "\u30d5\u30a1\u30a4\u30eb\u3092\u9078\u629e: "
                                                            + row.getName());
                                            FileChooserForEnvKey.apply(fc, row.getName());
                                            String cur = row.getValue();
                                            if (cur != null && !cur.isBlank()) {
                                                try {
                                                    Path p = Path.of(cur.trim());
                                                    if (Files.isRegularFile(p)) {
                                                        fc.setInitialDirectory(
                                                                p.getParent() != null
                                                                        ? p.getParent().toFile()
                                                                        : null);
                                                        fc.setInitialFileName(p.getFileName().toString());
                                                    } else if (Files.isDirectory(p)) {
                                                        fc.setInitialDirectory(p.toFile());
                                                    } else {
                                                        Path par = p.getParent();
                                                        if (par != null && Files.isDirectory(par)) {
                                                            fc.setInitialDirectory(par.toFile());
                                                        }
                                                    }
                                                } catch (Exception ignored) {
                                                    // keep defaults
                                                }
                                            }
                                            File file = fc.showOpenDialog(ownerStage);
                                            if (file != null) {
                                                row.setValue(file.getAbsolutePath());
                                                envTable.refresh();
                                            }
                                        });
                            }

                            @Override
                            protected void updateItem(Void item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    setGraphic(null);
                                    return;
                                }
                                TableRow<EnvVarRow> tr = getTableRow();
                                EnvVarRow row = tr != null ? tr.getItem() : null;
                                if (row == null) {
                                    int i = getIndex();
                                    if (i >= 0 && i < getTableView().getItems().size()) {
                                        row = getTableView().getItems().get(i);
                                    }
                                }
                                String key = row != null && row.getName() != null ? row.getName() : "";
                                if (row != null && AppPaths.isFolderPathEnvKey(key)) {
                                    setGraphic(pickFolder);
                                } else if (row != null && AppPaths.isFilePathEnvKey(key)) {
                                    setGraphic(pickFile);
                                } else {
                                    setGraphic(null);
                                }
                            }
                        });

        TableColumn<EnvVarRow, String> descCol =
                new TableColumn<>(
                        "\u8aac\u660e\uff08\u30b7\u30fc\u30c8+\u30ed\u30b8\u30c3\u30af\uff09");
        descCol.setCellValueFactory(cdf -> cdf.getValue().descriptionProperty());
        descCol.setPrefWidth(420);
        descCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final Text text = new Text();

                            {
                                text.wrappingWidthProperty().bind(col.widthProperty().subtract(16));
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setGraphic(null);
                                } else {
                                    text.setText(item);
                                    setGraphic(text);
                                }
                            }
                        });

        envTable.getColumns().setAll(nameCol, valueCol, folderCol, descCol);
        var envLayout =
                TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.ENV_VARS);
        if (!envLayout.isEmpty()) {
            TableColumnOrderPersistence.applyOrderToTableColumns(
                    envTable,
                    envLayout.stream().map(TableColumnOrderPersistence.ColumnSpec::title).toList());
            TableColumnOrderPersistence.applyWidthsToTableColumns(envTable, envLayout, 112);
        }
        TableColumnOrderPersistence.installColumnLayoutWatcher(
                envTable, TableColumnOrderPersistence.TableId.ENV_VARS, () -> false);

        Runnable resetEnvColumns =
                () -> {
                    nameCol.setPrefWidth(220);
                    valueCol.setPrefWidth(280);
                    folderCol.setPrefWidth(120);
                    descCol.setPrefWidth(420);
                };

        HBox strip = TableViewColumnSettingsStrip.create(envTable, resetEnvColumns, false);
        columnStripHost.getChildren().setAll(strip);

        VBox.setVgrow(envTable, Priority.ALWAYS);
    }
}
