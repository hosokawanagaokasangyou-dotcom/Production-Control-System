package jp.co.pm.ai.desktop;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;

import javafx.application.Platform;
import javafx.collections.ListChangeListener;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.PasswordField;
import javafx.scene.control.TextField;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.geometry.Pos;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.StringConverter;

import org.controlsfx.control.table.TableFilter;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.crypto.GeminiCredentialsV2Crypto;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.FileChooserForEnvKey;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.ui.TableHeaderColumnStyle;
import jp.co.pm.ai.desktop.ui.TableViewColumnSettingsStrip;

/**
 * Environment variables tab; columns/cell factories in code (FXML layout only). All Japanese UI strings use
 * \\u escapes so source stays portable across editors/OS encodings.
 */
public final class EnvTabController {

    /**
     * Avoids null {@code newValue} from the table cell converter (would clear the row as blank on commit).
     */
    private static final StringConverter<String> ENV_TABLE_STRING =
            new StringConverter<>() {
                @Override
                public String toString(String object) {
                    return object != null ? object : "";
                }

                @Override
                public String fromString(String string) {
                    return string != null ? string : "";
                }
            };

    private static EnvVarRow rowForEditCommit(TableColumn.CellEditEvent<EnvVarRow, String> e) {
        int row = e.getTablePosition().getRow();
        var items = e.getTableView().getItems();
        if (row >= 0 && row < items.size()) {
            EnvVarRow at = items.get(row);
            if (at != null) {
                return at;
            }
        }
        return e.getRowValue();
    }

    private static final String ENV_HINT_TEXT =
            "OS 環境変数は参照しません。このタブで集約。"
                    + " 初期値: ui_ref_env_defaults.json + ロジック説明。"
                    + " 子プロセス: この表 + メイン実行タブのマクロブックパス（任意）"
                    + "→ PYTHONUTF8 最終固定。"
                    + " PM_AI_SKIP_WORKBOOK_ENV_SHEET が空のときは 1 として"
                    + "マクロ「設定_環境変数」シートを読まない。"
                    + " フォルダ型は「フォルダ...」、各ファイル型は"
                    + "変数名に応じて JSON / Excel / CSV の拡張子を表示。";

    @FXML
    private Label hintLabel;

    @FXML
    private Label userDirLabel;

    @FXML
    private Label envSearchLabel;

    @FXML
    private TextField envSearchField;

    @FXML
    private HBox columnStripHost;

    @FXML
    private TableView<EnvVarRow> envTable;

    @FXML
    private Button addRowButton;

    @FXML
    private Button delRowButton;

    @FXML
    private Button addMissingEnvVarsButton;

    @FXML
    private Button resetEnvDefaultsButton;

    @FXML
    private Button encryptGeminiCredentialsButton;

    private Stage ownerStage;
    private MainShellController shell;
    private ObservableList<EnvVarRow> envRows;

    private final AtomicInteger headerColumnCount = new AtomicInteger(0);

    private TableFilter<EnvVarRow> envTableFilter;

    private FilteredList<EnvVarRow> envRowsFiltered;

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.ownerStage = shell.getPrimaryStage();
        this.envRows = shell.getEnvRows();
        if (userDirLabel != null) {
            userDirLabel.setText(System.getProperty("user.dir", "."));
        }
        hintLabel.setText(ENV_HINT_TEXT);
        if (envSearchLabel != null) {
            envSearchLabel.setText("検索");
        }
        if (envSearchField != null) {
            envSearchField.setPromptText("変数名・値・説明で絞り込み");
        }
        addRowButton.setText("行を追加");
        delRowButton.setText("行を削除");
        addMissingEnvVarsButton.setText("不足している環境変数を追加");
        resetEnvDefaultsButton.setText("環境変数を初期化");
        if (encryptGeminiCredentialsButton != null) {
            encryptGeminiCredentialsButton.setText("Gemini API キーを暗号化保存");
        }
        wireTable();
    }

    @FXML
    private void onAddRowButtonAction() {
        EnvVarRow r = new EnvVarRow();
        r.setDescription("");
        envRows.add(r);
    }

    @FXML
    private void onResetEnvDefaultsButtonAction() {
        if (shell != null) {
            shell.confirmAndResetEnvRowsToDefaults();
        }
    }

    @FXML
    private void onAddMissingEnvVarsButtonAction() {
        if (shell != null) {
            shell.addMissingReferenceEnvRows();
        }
    }

    @FXML
    private void onDelRowButtonAction() {
        var sel = envTable.getSelectionModel().getSelectedItems();
        if (!sel.isEmpty()) {
            envRows.removeAll(sel);
        } else if (!envRows.isEmpty()) {
            envRows.remove(envRows.size() - 1);
        }
        if (envRows.isEmpty()) {
            envRows.add(new EnvVarRow());
        }
    }

    private void wireTable() {
        envRowsFiltered = new FilteredList<>(envRows);
        envRowsFiltered.setPredicate(this::rowMatchesEnvSearch);
        envTable.setItems(envRowsFiltered);
        envTable.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        envTable.setEditable(true);
        envTable.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);

        TableColumn<EnvVarRow, String> nameCol = new TableColumn<>("変数名");
        nameCol.setCellValueFactory(cdf -> cdf.getValue().nameProperty());
        nameCol.setCellFactory(
                col ->
                        new TextFieldTableCell<EnvVarRow, String>(ENV_TABLE_STRING) {
                            @Override
                            public void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                TableHeaderColumnStyle.applyBodyCellTint(
                                        this, envTable, col, headerColumnCount::get);
                            }
                        });
        nameCol.setOnEditCommit(
                e -> {
                    EnvVarRow row = rowForEditCommit(e);
                    if (row != null) {
                        row.setName(e.getNewValue());
                    }
                });
        nameCol.setPrefWidth(220);
        nameCol.setReorderable(true);

        TableColumn<EnvVarRow, String> valueCol = new TableColumn<>("値");
        valueCol.setCellValueFactory(cdf -> cdf.getValue().valueProperty());
        valueCol.setCellFactory(
                col ->
                        new TextFieldTableCell<EnvVarRow, String>(ENV_TABLE_STRING) {
                            @Override
                            public void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                TableHeaderColumnStyle.applyBodyCellTint(
                                        this, envTable, col, headerColumnCount::get);
                            }
                        });
        valueCol.setOnEditCommit(
                e -> {
                    EnvVarRow row = rowForEditCommit(e);
                    if (row != null) {
                        row.setValue(e.getNewValue());
                    }
                });
        valueCol.setReorderable(true);

        TableColumn<EnvVarRow, Void> folderCol = new TableColumn<>("選択");
        folderCol.setPrefWidth(190);
        folderCol.setSortable(false);
        folderCol.setReorderable(false);
        folderCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final Button pickFolder =
                                    new Button("フォルダ...");
                            private final Button openFolder =
                                    new Button("開く");
                            private final Button pickFile =
                                    new Button("ファイル...");
                            private final HBox folderActions = new HBox(6);

                            {
                                folderActions.getChildren().addAll(pickFolder, openFolder);
                                openFolder.setOnAction(
                                        ev -> {
                                            EnvVarRow row =
                                                    getTableRow() != null ? getTableRow().getItem() : null;
                                            if (row == null) {
                                                int i = getIndex();
                                                if (i >= 0 && i < getTableView().getItems().size()) {
                                                    row = getTableView().getItems().get(i);
                                                }
                                            }
                                            openFolderLocationForRow(row);
                                        });
                                pickFolder.setOnAction(
                                        ev -> {
                                            EnvVarRow row =
                                                    getTableRow() != null ? getTableRow().getItem() : null;
                                            if (row == null) {
                                                return;
                                            }
                                            DirectoryChooser dc = new DirectoryChooser();
                                            dc.setTitle(
                                                    "フォルダを選択: "
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
                                                    "ファイルを選択: "
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
                                    TableHeaderColumnStyle.applyBodyCellTint(
                                            this, envTable, col, headerColumnCount::get);
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
                                if (row != null && AppPaths.isFilePathEnvKey(key)) {
                                    setGraphic(pickFile);
                                } else if (row != null && AppPaths.isFolderPathEnvKey(key)) {
                                    setGraphic(folderActions);
                                } else {
                                    setGraphic(null);
                                }
                                TableHeaderColumnStyle.applyBodyCellTint(
                                        this, envTable, col, headerColumnCount::get);
                            }
                        });

        TableColumn<EnvVarRow, String> descCol =
                new TableColumn<>(
                        "説明（シート+ロジック）");
        descCol.setCellValueFactory(cdf -> cdf.getValue().descriptionProperty());
        descCol.setPrefWidth(420);
        descCol.setReorderable(true);
        descCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            /** Labeled control respects themed {@code -fx-text-fill}; {@link javafx.scene.text.Text} does not. */
                            private final Label descLabel = new Label();

                            {
                                descLabel.setWrapText(true);
                                descLabel.setAlignment(Pos.TOP_LEFT);
                                descLabel.getStyleClass().add("pm-env-var-description");
                                descLabel.prefWidthProperty().bind(col.widthProperty().subtract(16));
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setGraphic(null);
                                } else {
                                    descLabel.setText(item);
                                    setGraphic(descLabel);
                                }
                                TableHeaderColumnStyle.applyBodyCellTint(
                                        this, envTable, col, headerColumnCount::get);
                            }
                        });

        envTable.getColumns().setAll(nameCol, valueCol, folderCol, descCol);
        envTableFilter = TableFilter.forTableView(envTable).apply();
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
                    folderCol.setPrefWidth(190);
                    descCol.setPrefWidth(420);
                };

        HBox strip =
                TableViewColumnSettingsStrip.create(
                        envTable,
                        resetEnvColumns,
                        false,
                        TableColumnOrderPersistence.TableId.ENV_VARS,
                        headerColumnCount,
                        () ->
                                ColumnVisibilitySupport.openTableViewColumnVisibilityDialog(
                                        ownerStage,
                                        TableColumnOrderPersistence.TableId.ENV_VARS,
                                        envTable));
        columnStripHost.getChildren().setAll(strip);

        VBox.setVgrow(envTable, Priority.ALWAYS);

        javafx.application.Platform.runLater(
                () ->
                        ColumnVisibilitySupport.applyColumnVisibilityToTableView(
                                envTable,
                                TableColumnOrderPersistence.loadColumnVisibility(
                                        TableColumnOrderPersistence.TableId.ENV_VARS,
                                        envTable.getColumns().size())));

        for (EnvVarRow r : envRows) {
            hookEnvRowForSearchFilter(r);
        }
        envRows.addListener(
                (ListChangeListener<EnvVarRow>)
                        c -> {
                            while (c.next()) {
                                if (c.wasAdded()) {
                                    for (EnvVarRow r : c.getAddedSubList()) {
                                        hookEnvRowForSearchFilter(r);
                                    }
                                }
                            }
                        });
        if (envSearchField != null) {
            envSearchField
                    .textProperty()
                    .addListener(
                            (o, a, b) -> {
                                applyEnvSearchPredicate();
                            });
        }
        applyEnvSearchPredicate();
    }

    private void hookEnvRowForSearchFilter(EnvVarRow r) {
        if (r == null) {
            return;
        }
        // Defer so FilteredList predicate does not run in the same pulse as table edit commit (avoids blank cells).
        Runnable ping = () -> Platform.runLater(this::applyEnvSearchPredicate);
        r.nameProperty().addListener((o, a, b) -> ping.run());
        r.valueProperty().addListener((o, a, b) -> ping.run());
        r.descriptionProperty().addListener((o, a, b) -> ping.run());
    }

    private void applyEnvSearchPredicate() {
        if (envRowsFiltered == null) {
            return;
        }
        envRowsFiltered.setPredicate(this::rowMatchesEnvSearch);
    }

    private String normalizedEnvSearchQuery() {
        if (envSearchField == null) {
            return "";
        }
        String t = envSearchField.getText();
        return t != null ? t.trim().toLowerCase(Locale.ROOT) : "";
    }

    private boolean rowMatchesEnvSearch(EnvVarRow r) {
        if (r == null) {
            return false;
        }
        String q = normalizedEnvSearchQuery();
        if (q.isEmpty()) {
            return true;
        }
        return containsNormalized(r.getName(), q)
                || containsNormalized(r.getValue(), q)
                || containsNormalized(r.getDescription(), q);
    }

    private static boolean containsNormalized(String s, String qLower) {
        if (s == null || s.isEmpty()) {
            return false;
        }
        return s.toLowerCase(Locale.ROOT).contains(qLower);
    }

    void clearColumnFiltersAndSort() {
        if (envSearchField != null) {
            envSearchField.clear();
        }
        if (envTableFilter != null) {
            envTableFilter.resetAllFilters();
        }
        envTable.getSortOrder().clear();
        applyEnvSearchPredicate();
    }

    @FXML
    private void onClearColumnFiltersAction() {
        clearColumnFiltersAndSort();
    }

    @FXML
    private void onEncryptGeminiCredentialsAction() {
        if (shell == null || ownerStage == null) {
            return;
        }
        Path target = resolveGeminiCredentialsJsonOutputPath();
        Dialog<String> dialog = new Dialog<>();
        dialog.initOwner(ownerStage);
        dialog.setTitle("Gemini 認証 JSON を暗号化保存");
        Label hint =
                new Label(
                        "暗号化は planning_core（Python）と互換の形式です。保存先は GEMINI_CREDENTIALS_JSON が優先、"
                                + "未設定時はリポジトリ code 配下の gemini_credentials.encrypted.json です。");
        hint.setWrapText(true);
        Label pathLab = new Label(target.toString());
        pathLab.setWrapText(true);
        PasswordField pf = new PasswordField();
        pf.setPromptText("Gemini API キー（平文）");
        VBox box = new VBox(8, hint, new Label("保存先:"), pathLab, new Label("API キー:"), pf);
        dialog.getDialogPane().setContent(box);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.setResultConverter(
                bt -> {
                    if (bt != ButtonType.OK) {
                        return null;
                    }
                    String t = pf.getText();
                    return t != null ? t.strip() : "";
                });
        Optional<String> opt = dialog.showAndWait();
        if (opt.isEmpty()) {
            return;
        }
        String apiKey = opt.get();
        if (apiKey.isEmpty()) {
            alertFolderOpen(ownerStage, AlertType.WARNING, "API キーが空です。");
            return;
        }
        try {
            GeminiCredentialsV2Crypto.writeEncryptedCredentials(target, apiKey);
            alertFolderOpen(
                    ownerStage,
                    AlertType.INFORMATION,
                    "暗号化 JSON を書き込みました。\n" + target);
        } catch (IllegalArgumentException | IOException | GeneralSecurityException ex) {
            String msg =
                    ex.getMessage() != null && !ex.getMessage().isBlank()
                            ? ex.getMessage()
                            : ex.getClass().getSimpleName();
            alertFolderOpen(ownerStage, AlertType.ERROR, "書き込みに失敗しました: " + msg);
        }
    }

    /**
     * {@link AppPaths#KEY_GEMINI_CREDENTIALS_JSON} の値があればそのパス。空ならリポジトリ {@code
     * code/gemini_credentials.encrypted.json}。
     */
    private Path resolveGeminiCredentialsJsonOutputPath() {
        if (envRows != null) {
            for (EnvVarRow r : envRows) {
                String k = r.getName() != null ? r.getName().strip() : "";
                if (AppPaths.KEY_GEMINI_CREDENTIALS_JSON.equals(k)) {
                    String v = r.getValue();
                    if (v != null && !v.isBlank()) {
                        return Path.of(v.strip()).toAbsolutePath().normalize();
                    }
                    break;
                }
            }
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        return AppPaths.resolveRepoRoot(ui)
                .resolve("code")
                .resolve("gemini_credentials.encrypted.json")
                .toAbsolutePath()
                .normalize();
    }

    /**
     * Opens the folder path from a folder-type env row (Explorer / Finder / xdg-open).
     */
    private void openFolderLocationForRow(EnvVarRow row) {
        if (ownerStage == null) {
            return;
        }
        String raw = row != null ? row.getValue() : null;
        if (raw == null || raw.isBlank()) {
            alertFolderOpen(ownerStage, AlertType.INFORMATION, "値が空のため開けません。");
            return;
        }
        Path p;
        try {
            p = Path.of(raw.trim()).toAbsolutePath().normalize();
        } catch (Exception ex) {
            alertFolderOpen(ownerStage, AlertType.WARNING, "パスが無効です。");
            return;
        }
        try {
            Path dir = null;
            if (Files.isDirectory(p)) {
                dir = p;
            } else if (Files.isRegularFile(p)) {
                dir = p.getParent();
            } else if (!Files.exists(p)) {
                Path par = p.getParent();
                if (par != null && Files.isDirectory(par)) {
                    dir = par;
                }
            }
            if (dir == null || !Files.isDirectory(dir)) {
                alertFolderOpen(
                        ownerStage,
                        AlertType.WARNING,
                        "開けるフォルダが見つかりません。");
                return;
            }
            if (!Desktop.isDesktopSupported() || !Desktop.getDesktop().isSupported(Desktop.Action.OPEN)) {
                alertFolderOpen(
                        ownerStage,
                        AlertType.WARNING,
                        "この環境では外部フォルダ開きに対応していません。");
                return;
            }
            Desktop.getDesktop().open(dir.toFile());
        } catch (Exception ex) {
            String msg =
                    ex.getMessage() != null && !ex.getMessage().isBlank()
                            ? ex.getMessage()
                            : "開けませんでした。";
            alertFolderOpen(ownerStage, AlertType.WARNING, msg);
        }
    }

    private static void alertFolderOpen(Stage owner, AlertType type, String message) {
        Alert a = new Alert(type);
        a.initOwner(owner);
        a.setHeaderText(null);
        a.setContentText(message);
        a.show();
    }
}
