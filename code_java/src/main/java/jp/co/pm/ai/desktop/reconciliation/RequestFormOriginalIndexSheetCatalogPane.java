package jp.co.pm.ai.desktop.reconciliation;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javafx.application.Platform;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.ContextMenu;
import javafx.scene.control.Label;
import javafx.scene.control.MenuItem;
import javafx.scene.control.OverrunStyle;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.input.KeyCode;
import javafx.scene.input.MouseButton;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;

/** 依頼書原本「目次」シートの一覧表示（{@link ReconciliationApp} 内タブ用）。 */
public final class RequestFormOriginalIndexSheetCatalogPane {

    private static final String HINT_SHORT =
            "転記・照合では目次の値を優先します。行を選んで「原本を開く」またはダブルクリック／Enter で Excel を読み取り専用で開きます。";
    private static final String HINT_DETAIL =
            "転記・照合で優先される目次の値を一覧表示します。"
                    + " 行をダブルクリック、Enter、または「原本を開く」で原本 xlsm を Excel で開きます（読み取り専用）。"
                    + " 依頼シートとの相違は「一括照合」画面のプレビュー横バナーでも確認できます。";
    private static final String ROW_OPEN_HINT = "ダブルクリックまたは Enter で原本を Excel 読み取り専用で開きます";

    private RequestFormOriginalIndexSheetCatalogPane() {}

    public static final class DisplayRow {
        private final SimpleStringProperty sourceFileName = new SimpleStringProperty();
        private final SimpleStringProperty iraiNo = new SimpleStringProperty();
        private final SimpleStringProperty orderRequestDate = new SimpleStringProperty();
        private final SimpleStringProperty responseDate = new SimpleStringProperty();
        private final SimpleStringProperty inputDate = new SimpleStringProperty();
        private final SimpleStringProperty deliveryDate = new SimpleStringProperty();
        private final SimpleStringProperty deliveryRemarks = new SimpleStringProperty();
        private final SimpleStringProperty contractDate = new SimpleStringProperty();
        private final SimpleStringProperty contractNo = new SimpleStringProperty();
        private final SimpleStringProperty contractRemarks = new SimpleStringProperty();
        private final SimpleStringProperty sourcePath = new SimpleStringProperty();

        static DisplayRow from(RequestFormOriginalIndexSheetCatalog.Row row) {
            DisplayRow d = new DisplayRow();
            d.sourceFileName.set(row.sourceFileName());
            d.iraiNo.set(row.iraiNo());
            d.orderRequestDate.set(row.orderRequestDate());
            d.responseDate.set(row.responseDate());
            d.inputDate.set(row.inputDate());
            d.deliveryDate.set(row.deliveryDate());
            d.deliveryRemarks.set(row.deliveryRemarks());
            d.contractDate.set(row.contractDate());
            d.contractNo.set(row.contractNo());
            d.contractRemarks.set(row.contractRemarks());
            d.sourcePath.set(row.sourcePath());
            return d;
        }

        public String getSourceFileName() {
            return sourceFileName.get();
        }

        public SimpleStringProperty sourceFileNameProperty() {
            return sourceFileName;
        }

        public String getIraiNo() {
            return iraiNo.get();
        }

        public SimpleStringProperty iraiNoProperty() {
            return iraiNo;
        }

        public String getOrderRequestDate() {
            return orderRequestDate.get();
        }

        public SimpleStringProperty orderRequestDateProperty() {
            return orderRequestDate;
        }

        public String getResponseDate() {
            return responseDate.get();
        }

        public SimpleStringProperty responseDateProperty() {
            return responseDate;
        }

        public String getInputDate() {
            return inputDate.get();
        }

        public SimpleStringProperty inputDateProperty() {
            return inputDate;
        }

        public String getDeliveryDate() {
            return deliveryDate.get();
        }

        public SimpleStringProperty deliveryDateProperty() {
            return deliveryDate;
        }

        public String getDeliveryRemarks() {
            return deliveryRemarks.get();
        }

        public SimpleStringProperty deliveryRemarksProperty() {
            return deliveryRemarks;
        }

        public String getContractDate() {
            return contractDate.get();
        }

        public SimpleStringProperty contractDateProperty() {
            return contractDate;
        }

        public String getContractNo() {
            return contractNo.get();
        }

        public SimpleStringProperty contractNoProperty() {
            return contractNo;
        }

        public String getContractRemarks() {
            return contractRemarks.get();
        }

        public SimpleStringProperty contractRemarksProperty() {
            return contractRemarks;
        }

        public String getSourcePath() {
            return sourcePath.get();
        }

        public SimpleStringProperty sourcePathProperty() {
            return sourcePath;
        }
    }

    public static VBox build(Supplier<Map<String, String>> uiEnv, Consumer<String> log) {
        Consumer<String> logFn = log != null ? log : s -> {};
        BooleanProperty loading = new SimpleBooleanProperty(false);

        Label title = new Label("依頼書原本 目次");
        title.getStyleClass().add("settings-card-title");

        Label note = new Label(HINT_SHORT);
        note.getStyleClass().add("paper-main-subtitle");
        note.setWrapText(false);
        note.setTextOverrun(OverrunStyle.ELLIPSIS);
        note.setTooltip(new Tooltip(HINT_DETAIL));
        note.setMaxWidth(Double.MAX_VALUE);

        Label statusLabel = new Label("読込中...");
        statusLabel.getStyleClass().add("index-sheet-catalog-count");
        statusLabel.setMinWidth(Region.USE_PREF_SIZE);

        Label actionLabel = new Label();
        actionLabel.getStyleClass().add("paper-main-subtitle");
        actionLabel.setWrapText(false);
        actionLabel.setTextOverrun(OverrunStyle.ELLIPSIS);
        actionLabel.setMaxWidth(Double.MAX_VALUE);
        HBox.setHgrow(actionLabel, Priority.SOMETIMES);

        Label sourceDirLabel = new Label();
        sourceDirLabel.getStyleClass().addAll("paper-main-subtitle", "index-sheet-catalog-path");
        sourceDirLabel.setWrapText(false);
        sourceDirLabel.setTextOverrun(OverrunStyle.ELLIPSIS);
        sourceDirLabel.setMaxWidth(320);
        sourceDirLabel.setMinWidth(80);

        TextField filterField = new TextField();
        filterField.setPromptText("依頼NO / 納期 / 投入日 / 契約No / 原本ファイル");
        filterField.setPrefWidth(280);
        filterField.setMaxWidth(Double.MAX_VALUE);
        HBox.setHgrow(filterField, Priority.ALWAYS);

        ObservableList<DisplayRow> backing = FXCollections.observableArrayList();
        FilteredList<DisplayRow> filtered = new FilteredList<>(backing, r -> true);
        int[] warningCount = {0};
        boolean[] loadFailed = {false};

        Runnable refreshCount =
                () -> {
                    if (loading.get() || loadFailed[0]) {
                        return;
                    }
                    statusLabel.getStyleClass().remove("excel-grid-label-error");
                    statusLabel.setText(
                            formatCountLabel(filtered.size(), backing.size(), warningCount[0]));
                    if (warningCount[0] > 0) {
                        statusLabel.setTooltip(new Tooltip("警告の詳細は実行ログを参照してください。"));
                    } else {
                        statusLabel.setTooltip(null);
                    }
                };

        Button btnClearFilter = new Button("解除");
        btnClearFilter.getStyleClass().add("btn-copy");
        btnClearFilter.setTooltip(new Tooltip("絞り込みを解除します"));
        btnClearFilter.setDisable(true);
        btnClearFilter.setOnAction(e -> filterField.clear());
        filterField
                .textProperty()
                .addListener(
                        (obs, old, text) -> {
                            filtered.setPredicate(
                                    row -> matchesFilter(row, text != null ? text : ""));
                            btnClearFilter.setDisable(text == null || text.isBlank());
                            refreshCount.run();
                        });

        TableView<DisplayRow> table = new TableView<>(filtered);
        table.getStyleClass().add("index-sheet-catalog-table");
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);
        table.setPlaceholder(new Label("目次行がありません。「再読込」で原本フォルダを走査してください。"));
        table.setFixedCellSize(28);
        table.getColumns().setAll(buildColumns());
        VBox.setVgrow(table, Priority.ALWAYS);
        table.setMinHeight(240);

        Button btnOpen = new Button("原本を開く");
        btnOpen.getStyleClass().add("btn-reload");
        btnOpen.setTooltip(new Tooltip("選択行の原本 xlsm を Excel で開きます（読み取り専用）"));
        btnOpen.setOnAction(
                e -> openSourceFile(table.getSelectionModel().getSelectedItem(), actionLabel::setText, logFn));
        btnOpen.disableProperty()
                .bind(loading.or(table.getSelectionModel().selectedItemProperty().isNull()));

        table.setRowFactory(
                tv -> {
                    TableRow<DisplayRow> row = new TableRow<>();
                    MenuItem openItem = new MenuItem("原本を開く（読み取り専用）");
                    openItem.setOnAction(
                            e -> openSourceFile(row.getItem(), actionLabel::setText, logFn));
                    ContextMenu menu = new ContextMenu(openItem);
                    row.emptyProperty()
                            .addListener(
                                    (o, was, empty) -> {
                                        row.setContextMenu(empty ? null : menu);
                                        row.setTooltip(empty ? null : new Tooltip(ROW_OPEN_HINT));
                                    });
                    row.setOnMouseClicked(
                            e -> {
                                if (e.getClickCount() != 2
                                        || row.isEmpty()
                                        || e.getButton() != MouseButton.PRIMARY) {
                                    return;
                                }
                                openSourceFile(row.getItem(), actionLabel::setText, logFn);
                            });
                    return row;
                });
        table.setOnKeyPressed(
                e -> {
                    if (e.getCode() != KeyCode.ENTER) {
                        return;
                    }
                    openSourceFile(
                            table.getSelectionModel().getSelectedItem(), actionLabel::setText, logFn);
                    e.consume();
                });

        Button btnReload = new Button("再読込");
        btnReload.getStyleClass().add("btn-reload");
        btnReload.setTooltip(
                new Tooltip("PM_AI_REQUEST_FORM_ORIGINAL_DIR 配下の xlsm から「目次」シートを再読込（読み取り専用）"));
        btnReload.disableProperty().bind(loading);

        Runnable reload =
                () -> {
                    if (loading.get()) {
                        return;
                    }
                    loading.set(true);
                    loadFailed[0] = false;
                    warningCount[0] = 0;
                    statusLabel.getStyleClass().remove("excel-grid-label-error");
                    statusLabel.setText("目次シートを読込中...");
                    actionLabel.setText("");
                    Map<String, String> ui = uiEnv.get();
                    Path dir =
                            AppPaths.resolveRequestFormOriginalDir(ui).toAbsolutePath().normalize();
                    sourceDirLabel.setText("原本フォルダ: " + dir);
                    sourceDirLabel.setTooltip(new Tooltip(dir.toString()));
                    Thread t =
                            new Thread(
                                    () -> {
                                        List<String> warnings = new ArrayList<>();
                                        try {
                                            List<RequestFormOriginalIndexSheetCatalog.Row> loaded =
                                                    RequestFormOriginalIndexSheetCatalog.loadAll(
                                                            ui,
                                                            warnings,
                                                            (processed, total) ->
                                                                    Platform.runLater(
                                                                            () -> {
                                                                                if (loading.get()) {
                                                                                    statusLabel
                                                                                            .setText(
                                                                                                    "読込中... ("
                                                                                                            + processed
                                                                                                            + " / "
                                                                                                            + total
                                                                                                            + " ファイル)");
                                                                                }
                                                                            }));
                                            Platform.runLater(
                                                    () -> {
                                                        backing.setAll(
                                                                loaded.stream()
                                                                        .map(DisplayRow::from)
                                                                        .toList());
                                                        warningCount[0] = warnings.size();
                                                        loading.set(false);
                                                        refreshCount.run();
                                                        if (!warnings.isEmpty()) {
                                                            for (String w : warnings) {
                                                                logFn.accept(
                                                                        "[index-sheet-catalog] "
                                                                                + w);
                                                            }
                                                        }
                                                    });
                                        } catch (Exception ex) {
                                            Platform.runLater(
                                                    () -> {
                                                        backing.clear();
                                                        warningCount[0] = 0;
                                                        loadFailed[0] = true;
                                                        loading.set(false);
                                                        statusLabel.setText(
                                                                "読込失敗: "
                                                                        + (ex.getMessage() != null
                                                                                ? ex.getMessage()
                                                                                : ex.toString()));
                                                        statusLabel
                                                                .getStyleClass()
                                                                .add("excel-grid-label-error");
                                                    });
                                            logFn.accept(
                                                    "[index-sheet-catalog] " + ex.getMessage());
                                        }
                                    },
                                    "index-sheet-catalog-reload");
                    t.setDaemon(true);
                    t.start();
                };
        btnReload.setOnAction(e -> reload.run());

        HBox actions =
                new HBox(
                        8,
                        btnReload,
                        btnOpen,
                        filterField,
                        btnClearFilter,
                        statusLabel,
                        actionLabel,
                        sourceDirLabel);
        actions.getStyleClass().add("index-sheet-catalog-toolbar");
        actions.setAlignment(Pos.CENTER_LEFT);

        VBox root = new VBox(8);
        root.getStyleClass().addAll("form-tab-container", "index-sheet-catalog");
        root.setFillWidth(true);
        root.setPadding(Insets.EMPTY);
        root.getChildren().addAll(title, note, actions, table);

        reload.run();
        return root;
    }

    static String formatCountLabel(int shown, int total, int warnings) {
        String base;
        if (shown == total) {
            base = "登録 " + total + " 件";
        } else {
            base = "表示 " + shown + " / 登録 " + total + " 件";
        }
        if (warnings > 0) {
            return base + "（警告 " + warnings + " 件）";
        }
        return base;
    }

    static boolean matchesFilter(DisplayRow row, String text) {
        if (text == null || text.isBlank()) {
            return true;
        }
        if (row == null) {
            return false;
        }
        String q = text.strip().toLowerCase(Locale.ROOT);
        return containsIgnoreCase(row.getIraiNo(), q)
                || containsIgnoreCase(row.getSourceFileName(), q)
                || containsIgnoreCase(row.getContractNo(), q)
                || containsIgnoreCase(row.getDeliveryDate(), q)
                || containsIgnoreCase(row.getInputDate(), q)
                || containsIgnoreCase(row.getOrderRequestDate(), q)
                || containsIgnoreCase(row.getResponseDate(), q)
                || containsIgnoreCase(row.getContractDate(), q)
                || containsIgnoreCase(row.getDeliveryRemarks(), q)
                || containsIgnoreCase(row.getContractRemarks(), q);
    }

    private static boolean containsIgnoreCase(String value, String q) {
        return value != null && value.toLowerCase(Locale.ROOT).contains(q);
    }

    static void openSourceFile(DisplayRow row, Consumer<String> statusUpdater, Consumer<String> log) {
        if (row == null) {
            return;
        }
        String pathText = row.getSourcePath();
        if (pathText == null || pathText.isBlank()) {
            if (statusUpdater != null) {
                statusUpdater.accept("原本ファイルのパスがありません。");
            }
            return;
        }
        Path path = Path.of(pathText);
        try {
            DesktopFileOpener.openFileReadOnly(path);
            if (statusUpdater != null) {
                statusUpdater.accept("開きました: " + path.getFileName());
            }
        } catch (Exception ex) {
            String msg =
                    "ファイルを開けません: "
                            + pathText
                            + (ex.getMessage() != null ? " — " + ex.getMessage() : "");
            if (statusUpdater != null) {
                statusUpdater.accept(msg);
            }
            if (log != null) {
                log.accept("[index-sheet-catalog] " + msg);
            }
        }
    }

    private static List<TableColumn<DisplayRow, String>> buildColumns() {
        List<TableColumn<DisplayRow, String>> cols = new ArrayList<>();
        cols.add(col("加工依頼NO", "iraiNo", 80, 96, "index-sheet-col-key"));
        cols.add(col("納期", "deliveryDate", 64, 80, "index-sheet-col-delivery"));
        cols.add(col("投入日", "inputDate", 64, 80, "index-sheet-col-date"));
        cols.add(col("契約No", "contractNo", 80, 110, "index-sheet-col-key"));
        cols.add(col("発注依頼日", "orderRequestDate", 64, 88, "index-sheet-col-date"));
        cols.add(col("回答日", "responseDate", 64, 80, "index-sheet-col-date"));
        cols.add(col("契約日", "contractDate", 64, 80, "index-sheet-col-date"));
        cols.add(col("納期備考", "deliveryRemarks", 60, 100, "index-sheet-col-remarks"));
        cols.add(col("契約備考", "contractRemarks", 60, 100, "index-sheet-col-remarks"));
        cols.add(col("原本ファイル", "sourceFileName", 120, 180, "index-sheet-col-file"));
        return cols;
    }

    private static TableColumn<DisplayRow, String> col(
            String title, String prop, double minWidth, double prefWidth, String cellStyle) {
        TableColumn<DisplayRow, String> c = new TableColumn<>(title);
        c.setCellValueFactory(new PropertyValueFactory<>(prop));
        c.setMinWidth(minWidth);
        c.setPrefWidth(prefWidth);
        c.setReorderable(true);
        c.getStyleClass().add(cellStyle);
        c.setCellFactory(
                col -> {
                    TableCell<DisplayRow, String> cell =
                            new TableCell<>() {
                                @Override
                                protected void updateItem(String item, boolean empty) {
                                    super.updateItem(item, empty);
                                    if (empty || item == null || item.isBlank()) {
                                        setText(null);
                                        setTooltip(null);
                                        return;
                                    }
                                    setText(item);
                                    setTooltip(new Tooltip(item));
                                }
                            };
                    cell.getStyleClass().add(cellStyle);
                    return cell;
                });
        return c;
    }
}
