package jp.co.pm.ai.desktop.reconciliation;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Supplier;

import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;

import org.controlsfx.control.table.TableFilter;

/** 依頼書原本「目次」シートの一覧表示（{@link ReconciliationApp} 内タブ用）。 */
public final class RequestFormOriginalIndexSheetCatalogPane {

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

        Label statusLabel = new Label("読込中...");
        statusLabel.setWrapText(true);
        statusLabel.setStyle("-fx-font-size: 11px; -fx-font-weight: bold;");

        Label sourceDirLabel = new Label();
        sourceDirLabel.setWrapText(true);
        sourceDirLabel.setStyle("-fx-font-size: 10px;");

        TextField filterField = new TextField();
        filterField.setPromptText("依頼NO / 原本ファイル / 契約No で絞り込み");
        filterField.setPrefWidth(280);

        ObservableList<DisplayRow> backing = FXCollections.observableArrayList();
        FilteredList<DisplayRow> filtered = new FilteredList<>(backing, r -> true);
        filterField
                .textProperty()
                .addListener(
                        (obs, old, text) ->
                                filtered.setPredicate(
                                        row -> matchesFilter(row, text != null ? text : "")));

        TableView<DisplayRow> table = new TableView<>(filtered);
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        table.setPlaceholder(new Label("目次行がありません。「再読込」で原本フォルダを走査してください。"));
        VBox.setVgrow(table, Priority.ALWAYS);
        table.setMinHeight(240);
        table.setPrefHeight(480);
        table.getColumns().setAll(buildColumns());
        TableFilter.forTableView(table).apply();
        table.setRowFactory(
                tv -> {
                    TableRow<DisplayRow> row = new TableRow<>();
                    row.setOnMouseClicked(
                            e -> {
                                if (e.getClickCount() != 2 || row.isEmpty()) {
                                    return;
                                }
                                openSourceFile(row.getItem(), statusLabel::setText, logFn);
                            });
                    return row;
                });

        Runnable reload =
                () -> {
                    statusLabel.setText("目次シートを読込中...");
                    Map<String, String> ui = uiEnv.get();
                    sourceDirLabel.setText(
                            "原本フォルダ: "
                                    + AppPaths.resolveRequestFormOriginalDir(ui)
                                            .toAbsolutePath()
                                            .normalize());
                    Thread t =
                            new Thread(
                                    () -> {
                                        List<String> warnings = new ArrayList<>();
                                        try {
                                            List<RequestFormOriginalIndexSheetCatalog.Row> loaded =
                                                    RequestFormOriginalIndexSheetCatalog.loadAll(
                                                            ui, warnings);
                                            Platform.runLater(
                                                    () -> {
                                                        backing.setAll(
                                                                loaded.stream()
                                                                        .map(DisplayRow::from)
                                                                        .toList());
                                                        if (!warnings.isEmpty()) {
                                                            statusLabel.setText(
                                                                    "登録 "
                                                                            + loaded.size()
                                                                            + " 件（警告 "
                                                                            + warnings.size()
                                                                            + " 件）");
                                                            for (String w : warnings) {
                                                                logFn.accept(
                                                                        "[index-sheet-catalog] "
                                                                                + w);
                                                            }
                                                        } else {
                                                            statusLabel.setText(
                                                                    "登録 " + loaded.size() + " 件");
                                                        }
                                                    });
                                        } catch (Exception ex) {
                                            Platform.runLater(
                                                    () -> {
                                                        backing.clear();
                                                        statusLabel.setText(
                                                                "読込失敗: "
                                                                        + (ex.getMessage() != null
                                                                                ? ex.getMessage()
                                                                                : ex.toString()));
                                                    });
                                            logFn.accept(
                                                    "[index-sheet-catalog] "
                                                            + ex.getMessage());
                                        }
                                    },
                                    "index-sheet-catalog-reload");
                    t.setDaemon(true);
                    t.start();
                };

        Button btnReload = new Button("再読込");
        btnReload.getStyleClass().add("btn-reload");
        btnReload.setTooltip(
                new Tooltip(
                        "PM_AI_REQUEST_FORM_ORIGINAL_DIR 配下の xlsm から「目次」シートを再読込（読み取り専用）"));
        btnReload.setOnAction(e -> reload.run());

        Label title = new Label("依頼書原本 目次シート");
        title.getStyleClass().add("paper-main-title");

        Label note =
                new Label(
                        "転記・照合で優先される目次の値を一覧表示します。"
                                + " 行をダブルクリックすると原本 xlsm を Excel で開きます（読み取り専用）。"
                                + " 依頼シートとの相違は「一括照合」画面のプレビュー横バナーでも確認できます。");
        note.setWrapText(true);
        note.setStyle("-fx-font-size: 10px;");

        HBox actions = new HBox(8, btnReload, filterField);
        actions.setAlignment(Pos.CENTER_LEFT);

        VBox root = new VBox(10);
        root.getStyleClass().add("form-tab-container");
        root.setFillWidth(true);
        root.setPadding(new Insets(8, 12, 12, 12));
        root.getChildren().addAll(title, note, actions, statusLabel, sourceDirLabel, table);

        reload.run();
        return root;
    }

    private static boolean matchesFilter(DisplayRow row, String text) {
        if (text.isBlank()) {
            return true;
        }
        String q = text.strip().toLowerCase();
        return containsIgnoreCase(row.getIraiNo(), q)
                || containsIgnoreCase(row.getSourceFileName(), q)
                || containsIgnoreCase(row.getContractNo(), q);
    }

    private static boolean containsIgnoreCase(String value, String q) {
        return value != null && value.toLowerCase().contains(q);
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
        cols.add(col("原本ファイル", "sourceFileName", 200));
        cols.add(col("加工依頼NO", "iraiNo", 88));
        cols.add(col("発注依頼日", "orderRequestDate", 88));
        cols.add(col("回答日", "responseDate", 72));
        cols.add(col("投入日", "inputDate", 72));
        cols.add(col("納期", "deliveryDate", 72));
        cols.add(col("納期備考", "deliveryRemarks", 120));
        cols.add(col("契約日", "contractDate", 72));
        cols.add(col("契約No", "contractNo", 100));
        cols.add(col("契約備考", "contractRemarks", 120));
        return cols;
    }

    private static TableColumn<DisplayRow, String> col(String title, String prop, double width) {
        TableColumn<DisplayRow, String> c = new TableColumn<>(title);
        c.setCellValueFactory(new PropertyValueFactory<>(prop));
        c.setMinWidth(width);
        c.setPrefWidth(width);
        c.setReorderable(true);
        return c;
    }
}
