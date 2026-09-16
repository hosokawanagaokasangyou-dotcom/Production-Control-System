package jp.co.pm.ai.desktop.kouchin;

import java.awt.Desktop;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.OverrunStyle;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.MouseButton;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.BorderPane;
import javafx.stage.FileChooser;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.MainShellTabId;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.ui.ButtonAttentionGlow;
import jp.co.pm.ai.kouchin.verify.BothResult;
import jp.co.pm.ai.kouchin.verify.FactoryId;
import jp.co.pm.ai.kouchin.verify.FileDiscovery;
import jp.co.pm.ai.kouchin.verify.Judge;
import jp.co.pm.ai.kouchin.verify.KouchinDiscovery;
import jp.co.pm.ai.kouchin.verify.KouchinPaths;
import jp.co.pm.ai.kouchin.verify.RecordA;
import jp.co.pm.ai.kouchin.verify.RecordB;
import jp.co.pm.ai.kouchin.verify.VerifyResult;
import jp.co.pm.ai.kouchin.verify.VerifyRunSupport;
import jp.co.pm.ai.kouchin.verify.VerifyService;

/**
 * 後加工工賃の検証タブ。POI/CSV/UNC 一覧は {@link #onMainShellTabSelected} でワーカーに載せる。
 * UI 更新は {@link Platform#runLater} のみ。
 */
public class KouchinVerifyTabController {

    @FXML private BorderPane root;
    @FXML private Label diffHintLabel;
    @FXML private Label dropTargetLabel;
    @FXML private Button runKokubuButton;
    @FXML private Button runKonanButton;
    @FXML private Button runBothButton;
    @FXML private Button importCsvButton;
    @FXML private Button openExcelButton;
    @FXML private Button openFolderButton;
    @FXML private Button copyMailButton;
    @FXML private Button gotoRdpButton;
    @FXML private Label statusLabel;
    @FXML private Label kpiLabel;
    @FXML private ComboBox<String> judgeFilterCombo;
    @FXML private TextField searchField;
    @FXML private TableView<KouchinDiscovery.Row> discoveryTable;
    @FXML private TableView<ResultLine> resultTable;
    @FXML private TextArea mailArea;
    @FXML private TextArea logArea;

    private MainShellController shell;
    private KouchinHostTabController host;
    private boolean selected;
    private boolean discoveryLoaded;
    private final AtomicInteger discoveryGeneration = new AtomicInteger();
    private VerifyRunSupport.Written lastWritten;
    private BothResult lastBoth;
    private ButtonAttentionGlow openExcelGlow;
    private final List<ResultLine> allResultLines = new ArrayList<>();

    public static final class ResultLine {
        private final String kind;
        private final String key;
        private final Double a1;
        private final Double a2;
        private final Double a3;
        private final Double diff;
        private final Double report;
        private final String judge;
        private final String note;

        public ResultLine(String kind, String key, Double a1, Double a2, Double a3, Double diff,
                           Double report, String judge, String note) {
            this.kind = kind;
            this.key = key;
            this.a1 = a1;
            this.a2 = a2;
            this.a3 = a3;
            this.diff = diff;
            this.report = report;
            this.judge = judge;
            this.note = note;
        }

        public String getKind() { return kind; }
        public String getKey() { return key; }
        public Double getA1() { return a1; }
        public String getA1Text() { return fmt(a1); }
        public String getA2Text() { return fmt(a2); }
        public String getA3Text() { return fmt(a3); }
        public String getDiffText() { return fmt(diff); }
        public String getReportText() { return fmt(report); }
        public String getJudge() { return judge; }
        public String getNote() { return note; }

        private static String fmt(Double v) {
            return v == null ? "" : String.format(Locale.US, "%,.0f", v);
        }
    }

    @FXML
    private void initialize() {
        if (statusLabel != null) {
            statusLabel.setText("まだ検証していません");
        }
        setupDiscoveryTable();
        setupResultTable();
        if (judgeFilterCombo != null) {
            judgeFilterCombo.getItems().setAll("すべて", Judge.MISMATCH, Judge.NEXT_MONTH, Judge.ONLY_1,
                    Judge.ONLY_2, Judge.ONLY_3, Judge.PREV_ADJUST, Judge.PREV_GAP, Judge.MANUAL_1,
                    Judge.MANUAL_2, Judge.MATCH);
            judgeFilterCombo.getSelectionModel().selectFirst();
            judgeFilterCombo.valueProperty().addListener((o, a, b) -> applyResultFilter());
        }
        if (searchField != null) {
            searchField.textProperty().addListener((o, a, b) -> applyResultFilter());
        }
        installDropHandlers();
        if (openExcelButton != null) {
            openExcelGlow = new ButtonAttentionGlow(openExcelButton);
        }
        refreshRunEnabled();
    }

    public void bindShell(MainShellController shell, KouchinHostTabController host) {
        this.shell = shell;
        this.host = host;
        refreshDropTargetLabel();
        refreshRunEnabled();
    }

    public void onMainShellTabSelected() {
        selected = true;
        refreshDropTargetLabel();
        refreshRunEnabled();
        if (!discoveryLoaded) {
            if (statusLabel != null && lastBoth == null) {
                statusLabel.setText("検出中…");
            }
            reloadDiscovery();
        }
    }

    public void onMainShellTabDeselected() {
        selected = false;
    }

    public void reloadDiscovery() {
        if (!selected && discoveryLoaded) {
            discoveryLoaded = false;
        }
        if (!selected) {
            return;
        }
        discoveryLoaded = true;
        if (statusLabel != null && lastBoth == null) {
            statusLabel.setText("検出中…");
        }
        Map<String, String> ui = shell == null ? Map.of() : shell.snapshotUiEnv();
        FactorySite site = shell == null ? FactorySite.KOKUBU : shell.currentFactorySite();
        FactoryId first = site == FactorySite.KONAN ? FactoryId.KONAN : FactoryId.KOKUBU;
        int gen = discoveryGeneration.incrementAndGet();
        Thread t = new Thread(() -> {
            List<KouchinDiscovery.Row> rows;
            String error = null;
            try {
                rows = new ArrayList<>(KouchinDiscovery.scan(first, KouchinPaths.fromEnv(ui)));
            } catch (RuntimeException e) {
                rows = List.of();
                error = e.getMessage();
            }
            List<KouchinDiscovery.Row> result = rows;
            String err = error;
            Platform.runLater(() -> applyDiscoveryResult(gen, result, err));
        }, "kouchin-discovery");
        t.setDaemon(true);
        t.start();
    }

    private void applyDiscoveryResult(int gen, List<KouchinDiscovery.Row> rows, String error) {
        if (gen != discoveryGeneration.get()) {
            return;
        }
        if (discoveryTable != null) {
            discoveryTable.getItems().setAll(rows);
        }
        if (statusLabel != null && lastBoth == null) {
            statusLabel.setText(error == null
                    ? "検出完了。未実行なら「まだ検証していません」。"
                    : "検出失敗: " + error);
        }
        refreshDropTargetLabel();
        refreshRunEnabled();
    }

    @FXML
    private void onRunKokubu() {
        runVerify(FactoryId.KOKUBU, false);
    }

    @FXML
    private void onRunKonan() {
        runVerify(FactoryId.KONAN, false);
    }

    @FXML
    private void onRunBoth() {
        runVerify(null, true);
    }

    void importCsvPaths(List<Path> files) {
        copyDropped(files);
    }

    @FXML
    private void onImportCsv() {
        Window w = dropTargetLabel == null ? null : dropTargetLabel.getScene() == null
                ? null : dropTargetLabel.getScene().getWindow();
        FileChooser fc = new FileChooser();
        fc.setTitle("①東レCSVを取り込む");
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("CSV", "*.csv"));
        List<File> files = fc.showOpenMultipleDialog(w);
        if (files == null || files.isEmpty()) {
            return;
        }
        copyDropped(files.stream().map(File::toPath).toList());
    }

    @FXML
    private void onOpenExcel() {
        List<Path> files = VerifyRunSupport.excelFilesToOpen(lastWritten);
        if (files.isEmpty()) {
            appendLog("開く Excel がありません（書込成功側が候補）");
            return;
        }
        int opened = 0;
        for (Path p : files) {
            try {
                DesktopFileOpener.openFile(p);
                opened++;
            } catch (Exception e) {
                appendLog("Excelを開けません: " + p.getFileName() + " " + e.getMessage());
            }
        }
        if (opened > 0) {
            setStatus("開いた: " + files.stream()
                    .map(p -> p.getFileName().toString())
                    .collect(Collectors.joining(" / ")));
            stopOpenExcelGlow();
        }
    }

    @FXML
    private void onOpenFolder() {
        Path dir = lastWritten == null || shell == null
                ? null
                : VerifyRunSupport.preferredOpenDir(lastWritten, shell.currentFactorySite());
        if (dir == null) {
            appendLog("出力フォルダがありません");
            return;
        }
        try {
            Desktop.getDesktop().open(dir.toFile());
            setStatus("開いた: " + dir.toAbsolutePath());
        } catch (Exception e) {
            appendLog("フォルダを開けません: " + e.getMessage());
        }
    }

    @FXML
    private void onCopyMail() {
        if (mailArea == null || mailArea.getText() == null || mailArea.getText().isBlank()) {
            return;
        }
        ClipboardContent c = new ClipboardContent();
        c.putString(mailArea.getText());
        Clipboard.getSystemClipboard().setContent(c);
        appendLog("報告メールをコピーしました");
    }

    @FXML
    private void onGotoRdp() {
        if (shell != null) {
            shell.selectMainShellTab(MainShellTabId.REMOTE_DESKTOP);
        }
    }

    private void runVerify(FactoryId one, boolean both) {
        if (shell == null) {
            return;
        }
        if (host != null && host.hasUnappliedSourceEdits()) {
            setStatus("参照先が未適用のため実行できません");
            return;
        }
        String label = both ? "後加工工賃 まとめて検証中…"
                : one == FactoryId.KONAN ? "後加工工賃 湖南工場 検証中…" : "後加工工賃 国分工場 検証中…";
        if (!shell.tryBeginKouchinRun(label)) {
            appendLog("他の実行中のため開始できません");
            return;
        }
        stopOpenExcelGlow();
        refreshRunEnabled();
        Map<String, String> ui = shell.snapshotUiEnv();
        KouchinPaths paths = KouchinPaths.fromEnv(ui);
        AtomicBoolean cancel = shell.kouchinCancelRequested();
        Task<Void> task = new Task<>() {
            @Override
            protected Void call() throws Exception {
                FileDiscovery.invalidateListingCache();
                BothResult bothResult;
                if (both) {
                    bothResult = VerifyService.runBoth(paths);
                } else if (one == FactoryId.KONAN) {
                    VerifyResult konan = VerifyService.run(FactoryId.KONAN, paths);
                    bothResult = new BothResult(null, konan, null, null, null);
                } else {
                    VerifyResult kokubu = VerifyService.run(FactoryId.KOKUBU, paths);
                    bothResult = new BothResult(kokubu, null, null, null, null);
                }
                if (cancel.get()) {
                    throw new IllegalStateException("中断されました");
                }
                VerifyRunSupport.Written written = VerifyRunSupport.writeBoth(
                        bothResult.kokubu(), bothResult.konan(), bothResult, ui, cancel);
                lastBoth = bothResult;
                lastWritten = written;
                return null;
            }
        };
        task.setOnSucceeded(e -> {
            shell.endKouchinRun();
            showResults();
            refreshRunEnabled();
        });
        task.setOnFailed(e -> {
            shell.endKouchinRun();
            Throwable err = task.getException();
            setStatus("失敗: " + (err == null ? "" : err.getMessage()));
            appendLog(String.valueOf(err == null ? "" : err.getMessage()));
            stopOpenExcelGlow();
            refreshRunEnabled();
        });
        Thread t = new Thread(task, "kouchin-verify");
        t.setDaemon(true);
        t.start();
    }

    private void showResults() {
        BothResult both = lastBoth;
        VerifyRunSupport.Written written = lastWritten;
        allResultLines.clear();
        if (both != null && both.kokubu() != null) {
            addResult(both.kokubu());
        }
        if (both != null && both.konan() != null) {
            addResult(both.konan());
        }
        applyResultFilter();
        if (mailArea != null && both != null) {
            mailArea.setText(both.unifiedMail() == null
                    ? (both.kokubu() != null || both.konan() != null
                            ? jp.co.pm.ai.kouchin.verify.UnifiedMailBuilder.buildText(
                                    both.kokubu() == null ? null : both.kokubu().mail(),
                                    both.konan() == null ? null : both.konan().mail())
                            : "")
                    : both.unifiedMail());
        }
        StringBuilder kpi = new StringBuilder();
        appendKpi(kpi, both == null ? null : both.kokubu());
        appendKpi(kpi, both == null ? null : both.konan());
        if (kpiLabel != null) {
            kpiLabel.setText(kpi.toString());
        }
        StringBuilder st = new StringBuilder("検証完了。");
        if (written != null) {
            appendOutcome(st, "Excel", written.xlsx());
            appendOutcome(st, "メールtxt", written.mailTxt());
            appendOutcome(st, "HTML", written.mailHtml());
        }
        if (both != null && both.kokubuError() != null) {
            st.append(" 国分失敗: ").append(both.kokubuError());
        }
        if (both != null && both.konanError() != null) {
            st.append(" 湖南失敗: ").append(both.konanError());
        }
        setStatus(st.toString());
        appendLog(st.toString());
        refreshOpenExcelGlow();
        reloadDiscovery();
    }

    private void appendKpi(StringBuilder kpi, VerifyResult r) {
        if (r == null) {
            return;
        }
        if (kpi.length() > 0) {
            kpi.append("  /  ");
        }
        kpi.append(r.profile().label())
                .append(" A要確認").append(r.requiredCheckA())
                .append(" B要確認").append(r.requiredCheckB())
                .append(" 報告過不足").append(r.num("報告する過不足"))
                .append(" 警告").append(r.warnings().size());
        if (r.checkD() != null) {
            kpi.append(" D要修正").append(r.checkD().errorCount());
        }
        if (r.checkC() != null) {
            kpi.append(" C判定").append(r.checkC().isSkipped() ? "スキップ" : ("要確認" + r.checkC().needCheckCount()));
        }
    }

    private void appendOutcome(StringBuilder st, String label, jp.co.pm.ai.kouchin.verify.DualWriteFiles.WriteOutcome o) {
        if (o == null) {
            return;
        }
        st.append(" ").append(label).append("成功").append(o.succeeded().size());
        if (o.anyFailed()) {
            st.append(" 失敗 ").append(String.join(" / ", o.failures()));
        }
    }

    private void addResult(VerifyResult r) {
        for (RecordA rec : r.recordsA()) {
            allResultLines.add(new ResultLine("A:" + r.profile().label(), rec.keiyaku(), rec.amount1(),
                    rec.amount2(), null, rec.diff(), rec.reportAmount(), rec.judge(), rec.note()));
        }
        for (RecordB rec : r.recordsB()) {
            allResultLines.add(new ResultLine("B:" + r.profile().label(), rec.irai(), null, rec.amount2(),
                    rec.amount3(), rec.diff(), null, rec.judge(), rec.note()));
        }
        for (String w : r.warnings()) {
            appendLog(r.profile().label() + " 警告: " + w);
        }
    }

    private void applyResultFilter() {
        if (resultTable == null) {
            return;
        }
        String judge = judgeFilterCombo == null ? "すべて" : judgeFilterCombo.getValue();
        String q = searchField == null || searchField.getText() == null
                ? "" : searchField.getText().trim().toUpperCase(Locale.ROOT);
        List<ResultLine> filtered = allResultLines.stream()
                .filter(r -> judge == null || "すべて".equals(judge) || judge.equals(r.getJudge()))
                .filter(r -> q.isEmpty()
                        || (r.getKey() != null && r.getKey().toUpperCase(Locale.ROOT).contains(q))
                        || (r.getNote() != null && r.getNote().toUpperCase(Locale.ROOT).contains(q)))
                .collect(Collectors.toList());
        resultTable.getItems().setAll(filtered);
    }

    private void copyDropped(List<Path> files) {
        if (shell == null) {
            return;
        }
        Path dest = KouchinPaths.fromEnv(shell.snapshotUiEnv()).torayCsvDir();
        Task<KouchinTorayCsvDropSupport.Outcome> task = new Task<>() {
            @Override
            protected KouchinTorayCsvDropSupport.Outcome call() {
                return KouchinTorayCsvDropSupport.copyCsvFiles(files, dest, shell.kouchinCancelRequested());
            }
        };
        task.setOnSucceeded(e -> {
            KouchinTorayCsvDropSupport.Outcome o = task.getValue();
            for (String w : o.warnings()) {
                appendLog(w);
            }
            for (String err : o.errors()) {
                appendLog(err);
            }
            if (o.copiedAny()) {
                FileDiscovery.invalidateListingCache();
                appendLog("同名上書きで保存: " + o.copied().size() + "件 → " + dest);
                reloadDiscovery();
            } else {
                appendLog("取り込みなし（検出は更新していません）");
            }
        });
        Thread t = new Thread(task, "kouchin-csv-drop");
        t.setDaemon(true);
        t.start();
    }

    private void installDropHandlers() {
        if (dropTargetLabel == null) {
            return;
        }
        dropTargetLabel.setOnDragOver(this::onDragOver);
        dropTargetLabel.setOnDragDropped(this::onDragDropped);
        if (discoveryTable != null) {
            discoveryTable.setOnDragOver(this::onDragOver);
            discoveryTable.setOnDragDropped(this::onDragDropped);
        }
    }

    private void onDragOver(DragEvent e) {
        Dragboard db = e.getDragboard();
        if (db.hasFiles()) {
            e.acceptTransferModes(TransferMode.COPY);
        }
        e.consume();
    }

    private void onDragDropped(DragEvent e) {
        Dragboard db = e.getDragboard();
        boolean ok = db.hasFiles();
        if (ok) {
            copyDropped(db.getFiles().stream().map(File::toPath).toList());
        }
        e.setDropCompleted(ok);
        e.consume();
    }

    private void refreshDropTargetLabel() {
        if (dropTargetLabel == null) {
            return;
        }
        Map<String, String> ui = shell == null ? Map.of() : shell.snapshotUiEnv();
        Path dest = KouchinPaths.fromEnv(ui).torayCsvDir();
        dropTargetLabel.setText("①東レCSVの取り込み先: " + dest.toAbsolutePath());
    }

    void refreshRunEnabled() {
        boolean unapplied = host != null && host.hasUnappliedSourceEdits();
        boolean busy = shell != null && (shell.isKouchinRunBusy() || shell.isPlanningPipelineStageRunning());
        boolean en = !unapplied && !busy;
        if (runKokubuButton != null) {
            runKokubuButton.setDisable(!en);
        }
        if (runKonanButton != null) {
            runKonanButton.setDisable(!en);
        }
        if (runBothButton != null) {
            runBothButton.setDisable(!en);
        }
        FactorySite site = shell == null ? FactorySite.KOKUBU : shell.currentFactorySite();
        if (runKokubuButton != null) {
            runKokubuButton.setStyle(site == FactorySite.KOKUBU ? "-fx-font-weight: bold;" : "");
        }
        if (runKonanButton != null) {
            runKonanButton.setStyle(site == FactorySite.KONAN ? "-fx-font-weight: bold;" : "");
        }
    }

    private void setupDiscoveryTable() {
        if (discoveryTable == null) {
            return;
        }
        discoveryTable.getColumns().clear();
        TableColumn<KouchinDiscovery.Row, String> roleCol = new TableColumn<>("区分");
        roleCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue() == null ? "" : cd.getValue().role()));
        roleCol.setPrefWidth(120);
        TableColumn<KouchinDiscovery.Row, String> pathCol = new TableColumn<>("ファイル");
        pathCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue() == null ? "" : cd.getValue().path()));
        pathCol.setPrefWidth(420);
        pathCol.setMinWidth(180);
        TableColumn<KouchinDiscovery.Row, String> ymCol = new TableColumn<>("対象月");
        ymCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue() == null ? "" : cd.getValue().ym()));
        ymCol.setPrefWidth(120);
        TableColumn<KouchinDiscovery.Row, String> noteCol = new TableColumn<>("備考");
        noteCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue() == null ? "" : cd.getValue().note()));
        noteCol.setPrefWidth(280);
        discoveryTable.getColumns().addAll(roleCol, pathCol, ymCol, noteCol);
        discoveryTable.setRowFactory(tv -> {
            TableRow<KouchinDiscovery.Row> row = new TableRow<>();
            row.setOnMouseClicked(e -> {
                if (e.getClickCount() != 2
                        || e.getButton() != MouseButton.PRIMARY
                        || row.isEmpty()) {
                    return;
                }
                openDiscoveryRow(row.getItem());
            });
            return row;
        });
        pathCol.setCellFactory(c -> new TableCell<>() {
            private final Tooltip tip = new Tooltip();

            {
                setTextOverrun(OverrunStyle.LEADING_ELLIPSIS);
                tip.setWrapText(true);
                tip.setMaxWidth(720);
            }

            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                setText(empty ? null : item);
                KouchinDiscovery.Row row = empty ? null : getTableRow().getItem();
                setStyle(row != null && row.missing() ? "-fx-text-fill: #c00000;" : "");
                if (row == null || row.fullPath() == null || row.fullPath().isBlank()) {
                    setTooltip(null);
                } else {
                    tip.setText(row.fullPath().replace('\\', '/'));
                    setTooltip(tip);
                }
            }
        });
    }

    static boolean shouldGlowOpenExcel(VerifyRunSupport.Written written, FactorySite site) {
        if (written == null || site == null) {
            return false;
        }
        return !VerifyRunSupport.excelFilesToOpen(written).isEmpty();
    }

    private void refreshOpenExcelGlow() {
        if (openExcelButton == null) {
            return;
        }
        if (openExcelGlow == null) {
            openExcelGlow = new ButtonAttentionGlow(openExcelButton);
        }
        FactorySite site = shell == null ? FactorySite.KOKUBU : shell.currentFactorySite();
        if (shouldGlowOpenExcel(lastWritten, site)) {
            openExcelGlow.ensureActive();
        } else {
            openExcelGlow.stop();
        }
    }

    private void stopOpenExcelGlow() {
        ButtonAttentionGlow.stopAll(openExcelGlow);
    }

    static Path openableDiscoveryFile(KouchinDiscovery.Row row) {
        if (row == null || row.missing()) {
            return null;
        }
        String full = row.fullPath();
        if (full == null || full.isBlank()) {
            return null;
        }
        Path p = Path.of(full);
        if (!Files.isRegularFile(p)) {
            return null;
        }
        return p.toAbsolutePath().normalize();
    }

    private void openDiscoveryRow(KouchinDiscovery.Row row) {
        Path p = openableDiscoveryFile(row);
        if (p == null) {
            appendLog("開けるファイルがありません");
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            setStatus("開いた: " + p.toAbsolutePath());
        } catch (Exception e) {
            appendLog("ファイルを開けません: " + e.getMessage());
        }
    }

    private void setupResultTable() {
        if (resultTable == null) {
            return;
        }
        resultTable.getColumns().clear();
        colR("区分", "kind", 90);
        colR("契約NO/依頼NO", "key", 140);
        colR("①金額", "a1Text", 90);
        colR("②金額", "a2Text", 90);
        colR("③金額", "a3Text", 90);
        colR("差額", "diffText", 90);
        colR("報告計上額", "reportText", 90);
        colR("判定", "judge", 110);
        colR("備考", "note", 360);
    }

    private void colR(String title, String prop, int w) {
        TableColumn<ResultLine, String> c = new TableColumn<>(title);
        c.setCellValueFactory(new PropertyValueFactory<>(prop));
        c.setPrefWidth(w);
        resultTable.getColumns().add(c);
    }

    private void setStatus(String text) {
        if (statusLabel != null) {
            statusLabel.setText(text);
        }
    }

    private void appendLog(String line) {
        if (logArea == null || line == null || line.isBlank()) {
            return;
        }
        Platform.runLater(() -> logArea.appendText(line + System.lineSeparator()));
        if (shell != null) {
            shell.appendLog("[kouchin] " + line);
        }
    }
}
