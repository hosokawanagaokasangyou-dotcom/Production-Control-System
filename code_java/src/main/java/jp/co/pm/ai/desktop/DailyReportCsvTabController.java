package jp.co.pm.ai.desktop;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;

import org.controlsfx.control.table.TableFilter;

import jp.co.pm.ai.desktop.reconciliation.KonanDailyReportLookup;
import jp.co.pm.ai.desktop.reconciliation.KonanDailyReportLookup.DailyReportCsvTable;
import jp.co.pm.ai.desktop.ui.ColumnVisibilitySupport;
import jp.co.pm.ai.desktop.ui.SourceExtensionErrorOverlay;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence;
import jp.co.pm.ai.desktop.ui.TableColumnOrderPersistence.ColumnSpec;
import jp.co.pm.ai.desktop.ui.TableViewColumnSettingsStrip;

/** 加工日報 CSV（加工日報発行問合せ）の閲覧タブ。 */
public final class DailyReportCsvTabController {

    private static final String HINT_TEXT =
            "PM_AI_DAILY_REPORT_CSV_PATH で単一ファイルを指定するか、"
                    + " PM_AI_DAILY_REPORT_SOURCE_DIR 配下の最新 "
                    + "加工日報発行問合せ_*.csv を表示します。"
                    + " 最新ファイルの拡張子が .csv 以外のときはエラーとし表を暗転します。"
                    + " 先頭3行はメタ情報、4行目が見出しです。"
                    + " 列ヘッダをクリックで並べ替え、ヘッダ右のフィルタアイコンで列単位の絞り込みができます。";

    @FXML
    private Button refreshButton;

    @FXML
    private Label statusLabel;

    @FXML
    private Label sourceLabel;

    @FXML
    private Label metaLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private TextField searchField;

    @FXML
    private HBox columnStripHost;

    @FXML
    private StackPane tableHost;

    @FXML
    private TableView<Map<String, String>> table;

    private ObservableList<Map<String, String>> rows;

    private FilteredList<Map<String, String>> rowsFiltered;

    private MainShellController shell;

    private TableFilter<Map<String, String>> tableFilter;

    private List<String> currentHeaders = List.of();

    private final AtomicBoolean suppressColumnPersistence = new AtomicBoolean(false);

    private final AtomicReference<List<ColumnSpec>> persistedLayout = new AtomicReference<>(List.of());

    private boolean tableChromeInitialized;

    @FXML
    private void initialize() {
        hintLabel.setText(HINT_TEXT);
        rows = FXCollections.observableArrayList();
        rowsFiltered = new FilteredList<>(rows, this::rowMatchesSearch);
        table.setItems(rowsFiltered);
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        table.setPlaceholder(new Label("データがありません。「再読込」で加工日報 CSV を読み直してください。"));
        installTableScrollLayout();
        if (searchField != null) {
            searchField
                    .textProperty()
                    .addListener((obs, old, text) -> applySearchPredicate());
        }
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        installTableChromeOnce();
        reloadAsync();
    }

    @FXML
    private void onRefreshButtonAction() {
        reloadAsync();
    }

    @FXML
    private void onClearColumnFiltersAction() {
        if (searchField != null) {
            searchField.clear();
        }
        if (tableFilter != null) {
            tableFilter.resetAllFilters();
        }
        if (table != null) {
            table.getSortOrder().clear();
        }
        applySearchPredicate();
        updateStatusLabel();
    }

    private void installTableChromeOnce() {
        if (tableChromeInitialized || table == null || columnStripHost == null) {
            return;
        }
        tableChromeInitialized = true;
        TableColumnOrderPersistence.installColumnLayoutWatcher(
                table,
                TableColumnOrderPersistence.TableId.DAILY_REPORT_CSV,
                suppressColumnPersistence::get);
        Runnable resetColumns = this::resetColumnWidthsToDefault;
        columnStripHost
                .getChildren()
                .setAll(
                        TableViewColumnSettingsStrip.create(
                                table,
                                resetColumns,
                                false,
                                TableColumnOrderPersistence.TableId.DAILY_REPORT_CSV,
                                null,
                                () ->
                                        ColumnVisibilitySupport.openTableViewColumnVisibilityDialog(
                                                shell != null ? shell.getPrimaryStage() : null,
                                                TableColumnOrderPersistence.TableId.DAILY_REPORT_CSV,
                                                table)));
    }

    private void installTableScrollLayout() {
        if (table == null || tableHost == null) {
            return;
        }
        VBox.setVgrow(tableHost, Priority.ALWAYS);
        table.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
        table.prefWidthProperty().bind(tableHost.widthProperty());
        table.prefHeightProperty().bind(tableHost.heightProperty());
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
    }

    private void reloadAsync() {
        if (shell == null) {
            statusLabel.setText("シェル未接続");
            return;
        }
        statusLabel.setText("読込中...");
        refreshButton.setDisable(true);
        Map<String, String> ui = shell.snapshotUiEnv();
        Thread t =
                new Thread(
                        () -> {
                            try {
                                DailyReportCsvTable loaded =
                                        KonanDailyReportLookup.readLatestTable(ui);
                                Platform.runLater(() -> applyLoaded(loaded));
                            } catch (Exception ex) {
                                String msg =
                                        ex.getMessage() != null ? ex.getMessage() : ex.toString();
                                Platform.runLater(
                                        () -> {
                                            rows.clear();
                                            table.getColumns().clear();
                                            currentHeaders = List.of();
                                            tableFilter = null;
                                            statusLabel.setText("読込失敗: " + msg);
                                            sourceLabel.setText("");
                                            metaLabel.setText("");
                                            refreshButton.setDisable(false);
                                            if (msg.contains("拡張子が不正")) {
                                                SourceExtensionErrorOverlay.show(tableHost, msg);
                                            } else {
                                                SourceExtensionErrorOverlay.clear(tableHost);
                                            }
                                        });
                            }
                        },
                        "daily-report-csv-reload");
        t.setDaemon(true);
        t.start();
    }

    private void applyLoaded(DailyReportCsvTable loaded) {
        List<String> headers =
                loaded.headers() != null ? List.copyOf(loaded.headers()) : List.of();
        boolean headersChanged = !Objects.equals(currentHeaders, headers);
        if (headersChanged) {
            persistedLayout.set(
                    TableColumnOrderPersistence.loadLayout(
                            TableColumnOrderPersistence.TableId.DAILY_REPORT_CSV));
            rebuildColumns(headers);
            currentHeaders = headers;
            reinstallTableFilter();
            Platform.runLater(
                    () ->
                            ColumnVisibilitySupport.applyColumnVisibilityToTableView(
                                    table,
                                    TableColumnOrderPersistence.loadColumnVisibility(
                                            TableColumnOrderPersistence.TableId.DAILY_REPORT_CSV,
                                            table.getColumns().size())));
        }
        rows.setAll(loaded.rows());
        applySearchPredicate();
        sourceLabel.setText("参照ファイル: " + loaded.sourcePath());
        metaLabel.setText(formatMetaLines(loaded.metaLines()));
        refreshButton.setDisable(false);
        SourceExtensionErrorOverlay.clear(tableHost);
        updateStatusLabel();
    }

    private void reinstallTableFilter() {
        tableFilter = TableFilter.forTableView(table).apply();
    }

    private void applySearchPredicate() {
        if (rowsFiltered != null) {
            rowsFiltered.setPredicate(this::rowMatchesSearch);
        }
        updateStatusLabel();
    }

    private void updateStatusLabel() {
        if (statusLabel == null) {
            return;
        }
        int total = rows != null ? rows.size() : 0;
        int shown = rowsFiltered != null ? rowsFiltered.size() : total;
        if (shown == total) {
            statusLabel.setText("データ行 " + total + " 件");
        } else {
            statusLabel.setText("表示 " + shown + " / 全 " + total + " 件");
        }
    }

    static boolean rowMatchesSearch(Map<String, String> row, String query) {
        String q = query != null ? query.strip().toLowerCase(Locale.ROOT) : "";
        if (q.isEmpty()) {
            return true;
        }
        if (row == null) {
            return false;
        }
        for (String value : row.values()) {
            if (value != null && value.toLowerCase(Locale.ROOT).contains(q)) {
                return true;
            }
        }
        return false;
    }

    private boolean rowMatchesSearch(Map<String, String> row) {
        String query = searchField != null ? searchField.getText() : "";
        return rowMatchesSearch(row, query);
    }

    private static String formatMetaLines(List<String> metaLines) {
        if (metaLines == null || metaLines.isEmpty()) {
            return "";
        }
        List<String> parts = new ArrayList<>();
        for (int i = 0; i < metaLines.size(); i++) {
            String line = metaLines.get(i) != null ? metaLines.get(i).strip() : "";
            parts.add("メタ" + (i + 1) + ": " + (line.isEmpty() ? "（空）" : line));
        }
        return String.join(" / ", parts);
    }

    private void rebuildColumns(List<String> headers) {
        suppressColumnPersistence.set(true);
        try {
            table.getColumns().clear();
            if (headers == null || headers.isEmpty()) {
                return;
            }
            List<Double> widths =
                    TableColumnOrderPersistence.resolveWidthsForHeaders(
                            headers, persistedLayout.get(), 112);
            for (int i = 0; i < headers.size(); i++) {
                String header = headers.get(i);
                String title = header != null && !header.isBlank() ? header : "列";
                TableColumn<Map<String, String>, String> col = new TableColumn<>(title);
                col.setCellValueFactory(
                        cd -> {
                            Map<String, String> row = cd.getValue();
                            String v = row != null ? row.getOrDefault(header, "") : "";
                            return new SimpleStringProperty(v != null ? v : "");
                        });
                double width =
                        i < widths.size() && widths.get(i) != null && widths.get(i) > 0
                                ? widths.get(i)
                                : defaultColumnWidth(title);
                col.setMinWidth(72);
                col.setPrefWidth(width);
                col.setReorderable(true);
                col.setSortable(true);
                table.getColumns().add(col);
            }
            List<String> titles =
                    headers.stream()
                            .map(h -> h != null && !h.isBlank() ? h : "列")
                            .toList();
            TableColumnOrderPersistence.applyOrderToTableColumns(table, titles);
            TableColumnOrderPersistence.applyWidthsToTableColumns(
                    table, persistedLayout.get(), 112);
        } finally {
            suppressColumnPersistence.set(false);
        }
    }

    private void resetColumnWidthsToDefault() {
        suppressColumnPersistence.set(true);
        try {
            for (TableColumn<Map<String, String>, ?> col : table.getColumns()) {
                String title = col.getText();
                double width = defaultColumnWidth(title != null ? title : "列");
                col.setMinWidth(72);
                col.setPrefWidth(width);
            }
        } finally {
            suppressColumnPersistence.set(false);
        }
    }

    private static double defaultColumnWidth(String title) {
        int len = title != null ? title.length() : 4;
        return Math.min(Math.max(len * 9 + 36, 80), 200);
    }
}
