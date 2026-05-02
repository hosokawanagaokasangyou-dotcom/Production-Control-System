package jp.co.pm.ai.desktop;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;

import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
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

/** Environment variables tab; columns and cell factories in code (FXML is layout only). */
public final class EnvTabController {

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

        TableColumn<EnvVarRow, String> nameCol = new TableColumn<>("?????");
        nameCol.setCellValueFactory(cdf -> cdf.getValue().nameProperty());
        nameCol.setCellFactory(TextFieldTableCell.forTableColumn());
        nameCol.setOnEditCommit(
                e -> {
                    e.getRowValue().setName(e.getNewValue());
                    envTable.refresh();
                });
        nameCol.setPrefWidth(220);

        TableColumn<EnvVarRow, String> valueCol = new TableColumn<>("?l");
        valueCol.setCellValueFactory(cdf -> cdf.getValue().valueProperty());
        valueCol.setCellFactory(TextFieldTableCell.forTableColumn());
        valueCol.setOnEditCommit(
                e -> {
                    e.getRowValue().setValue(e.getNewValue());
                    envTable.refresh();
                });

        TableColumn<EnvVarRow, Void> folderCol = new TableColumn<>("?I??");
        folderCol.setPrefWidth(120);
        folderCol.setSortable(false);
        folderCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final Button pickFolder = new Button("?t?H???_...");
                            private final Button pickFile = new Button("?t?@?C??...");

                            {
                                pickFolder.setOnAction(
                                        ev -> {
                                            EnvVarRow row =
                                                    getTableRow() != null ? getTableRow().getItem() : null;
                                            if (row == null) {
                                                return;
                                            }
                                            DirectoryChooser dc = new DirectoryChooser();
                                            dc.setTitle("?t?H???_??I??: " + row.getName());
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
                                            fc.setTitle("?t?@?C????I??: " + row.getName());
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

        TableColumn<EnvVarRow, String> descCol = new TableColumn<>("?????i?V?[?g+???W?b?N?j");
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
        var envLayout = TableColumnOrderPersistence.loadLayout(TableColumnOrderPersistence.TableId.ENV_VARS);
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
