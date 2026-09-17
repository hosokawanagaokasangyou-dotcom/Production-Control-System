package jp.co.pm.ai.desktop.kouchin;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CompletionException;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Function;
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
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;
import jp.co.pm.ai.desktop.ui.ButtonAttentionGlow;
import jp.co.pm.ai.desktop.ui.JapanDateTimeDisplay;
import jp.co.pm.ai.kouchin.verify.BothResult;
import jp.co.pm.ai.kouchin.verify.FactoryId;
import jp.co.pm.ai.kouchin.verify.FileDiscovery;
import jp.co.pm.ai.kouchin.verify.Judge;
import jp.co.pm.ai.kouchin.verify.KouchinDiscovery;
import jp.co.pm.ai.kouchin.verify.KouchinOutputDirs;
import jp.co.pm.ai.kouchin.verify.KouchinPaths;
import jp.co.pm.ai.kouchin.verify.RecordA;
import jp.co.pm.ai.kouchin.verify.RecordB;
import jp.co.pm.ai.kouchin.verify.UnifiedMailBuilder;
import jp.co.pm.ai.kouchin.verify.VerifyOutputAccess;
import jp.co.pm.ai.kouchin.verify.VerifyResult;
import jp.co.pm.ai.kouchin.verify.VerifyRunSupport;
import jp.co.pm.ai.kouchin.verify.VerifyService;
import jp.co.pm.ai.kouchin.verify.VerifySourceAccess;

/**
 * 後加工工賃の検証タブ。POI/CSV/UNC 一覧は {@link #onMainShellTabSelected} でワーカーに載せる。
 * UI 更新は {@link Platform#runLater} のみ。
 */
public class KouchinVerifyTabController {

    private static final DateTimeFormatter FILE_MODIFIED_AT =
            DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss", Locale.JAPAN);

    @FXML private BorderPane root;
    @FXML private Label diffHintLabel;
    @FXML private Label dropTargetLabel;
    @FXML private Button runKokubuButton;
    @FXML private Button runKonanButton;
    @FXML private Button runBothButton;
    @FXML private Label outputWriteBlockBadge;
    @FXML private Button importCsvButton;
    @FXML private Button openKokubuExcelButton;
    @FXML private Button openKonanExcelButton;
    @FXML private Button openKokubuSource2Button;
    @FXML private Button openKonanSource2Button;
    @FXML private Button openFolderButton;
    @FXML private Button copyMailButton;
    @FXML private Button gotoRdpButton;
    @FXML private Label statusLabel;
    @FXML private Label kpiLabel;
    @FXML private Label skipBBanner;
    @FXML private Label skipBResultBanner;
    @FXML private Label targetYmReasonLabel;
    @FXML private ComboBox<String> judgeFilterCombo;
    @FXML private TextField searchField;
    @FXML private TableView<DiscoveryLine> discoveryTable;
    @FXML private TableView<ResultLine> resultTable;
    @FXML private TextArea mailArea;
    @FXML private TextArea logArea;

    private MainShellController shell;
    private KouchinHostTabController host;
    private boolean selected;
    private boolean discoveryLoaded;
    private boolean pendingSourceReload;
    private final AtomicInteger discoveryGeneration = new AtomicInteger();
    private VerifyRunSupport.Written lastWritten;
    private BothResult lastBoth;
    private List<KouchinDiscovery.Row> lastKokubuDiscovery = List.of();
    private List<KouchinDiscovery.Row> lastKonanDiscovery = List.of();
    private String lastKokubuBlock = "検出未完了";
    private String lastKonanBlock = "検出未完了";
    private boolean lastOutputWritable;
    private boolean kokubuVerified;
    private boolean konanVerified;
    private ButtonAttentionGlow runKokubuGlow;
    private ButtonAttentionGlow runKonanGlow;
    private ButtonAttentionGlow runBothGlow;
    private ButtonAttentionGlow openKokubuExcelGlow;
    private ButtonAttentionGlow openKonanExcelGlow;
    private final List<ResultLine> allResultLines = new ArrayList<>();

    public static final class DiscoveryLine {
        static final String FACTORY_SHARED = "国分/湖南共通";

        private final String factory;
        private final KouchinDiscovery.Row row;
        private final VerifySourceAccess.FileAccess access;
        private final String modifiedAt;
        private final boolean openable;

        DiscoveryLine(String factory, KouchinDiscovery.Row row, VerifySourceAccess.FileAccess access) {
            this.factory = factory == null ? "" : factory;
            this.row = row;
            this.access = access == null ? VerifySourceAccess.FileAccess.of(null) : access;
            this.modifiedAt = fileModifiedAtText(row);
            this.openable = openableDiscoveryFile(row) != null;
        }

        static DiscoveryLine of(String factory, KouchinDiscovery.Row row) {
            return new DiscoveryLine(factory, row, VerifySourceAccess.FileAccess.ofRow(row));
        }

        public String getFactory() { return factory; }
        public String getRole() { return row == null ? "" : row.role(); }
        public String getPath() { return row == null ? "" : row.path(); }
        public String getPathCss() { return isMissing() ? "pm-kouchin-missing-path" : ""; }
        public String getYm() { return row == null ? "" : row.ym(); }
        public String getModifiedAt() { return modifiedAt; }
        public boolean canOpen() { return openable; }
        public String getNote() { return row == null ? "" : row.note(); }
        public String getReadStatus() { return access.readLabel(); }
        public String getWriteStatus() { return access.writeLabel(); }
        public String getReadCss() { return access.readCss(); }
        public String getWriteCss() { return access.writeCss(); }
        public String getReadHint() { return access.readHint(); }
        public String getWriteHint() { return access.writeHint(); }
        public boolean isMissing() { return row != null && row.missing(); }
        public KouchinDiscovery.Row source() { return row; }
        public VerifySourceAccess.FileAccess access() { return access; }
    }

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

    private record VerifyTaskOutcome(BothResult both, VerifyRunSupport.Written written) {}

    static List<DiscoveryLine> buildDiscoveryLines(
            List<KouchinDiscovery.Row> kokubu, List<KouchinDiscovery.Row> konan) {
        Map<String, VerifySourceAccess.FileAccess> memo = new HashMap<>();
        List<DiscoveryLine> lines = new ArrayList<>();
        if (kokubu != null) {
            for (KouchinDiscovery.Row r : kokubu) {
                lines.add(new DiscoveryLine(factoryLabelFor("国分", r), r, accessFor(r, memo)));
            }
        }
        if (konan != null) {
            for (KouchinDiscovery.Row r : konan) {
                lines.add(new DiscoveryLine(factoryLabelFor("湖南", r), r, accessFor(r, memo)));
            }
        }
        return lines;
    }

    private static VerifySourceAccess.FileAccess accessFor(
            KouchinDiscovery.Row row, Map<String, VerifySourceAccess.FileAccess> memo) {
        if (row == null || row.missing()) {
            return VerifySourceAccess.FileAccess.ofRow(row);
        }
        String key = row.fullPath();
        if (key == null || key.isBlank()) {
            return VerifySourceAccess.FileAccess.ofRow(row);
        }
        return memo.computeIfAbsent(key, k -> VerifySourceAccess.FileAccess.ofRow(row));
    }

    static String factoryLabelFor(String site, KouchinDiscovery.Row row) {
        if (row != null && KouchinDiscovery.ROLE_3.equals(row.role())) {
            return DiscoveryLine.FACTORY_SHARED;
        }
        return site == null ? "" : site;
    }

    static boolean factoryMatches(String wanted, String lineFactory) {
        if (wanted == null || lineFactory == null) {
            return false;
        }
        if (wanted.equals(lineFactory)) {
            return true;
        }
        return DiscoveryLine.FACTORY_SHARED.equals(lineFactory)
                && ("国分".equals(wanted) || "湖南".equals(wanted));
    }

    private static Function<KouchinDiscovery.Row, VerifySourceAccess.FileAccess> accessLookup(
            List<DiscoveryLine> lines, String factory) {
        Map<String, VerifySourceAccess.FileAccess> byRole = new HashMap<>();
        if (lines != null) {
            for (DiscoveryLine line : lines) {
                if (line != null && factoryMatches(factory, line.getFactory()) && line.source() != null) {
                    byRole.put(line.getRole(), line.access());
                }
            }
        }
        return row -> {
            if (row == null) {
                return VerifySourceAccess.FileAccess.of(null);
            }
            VerifySourceAccess.FileAccess acc = byRole.get(row.role());
            return acc == null ? VerifySourceAccess.FileAccess.ofRow(row) : acc;
        };
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
        if (runKokubuButton != null) {
            runKokubuGlow = new ButtonAttentionGlow(runKokubuButton);
        }
        if (runKonanButton != null) {
            runKonanGlow = new ButtonAttentionGlow(runKonanButton);
        }
        if (runBothButton != null) {
            runBothGlow = new ButtonAttentionGlow(runBothButton);
        }
        if (openKokubuExcelButton != null) {
            openKokubuExcelGlow = new ButtonAttentionGlow(openKokubuExcelButton);
        }
        if (openKonanExcelButton != null) {
            openKonanExcelGlow = new ButtonAttentionGlow(openKonanExcelButton);
        }
        refreshRunEnabled();
        refreshOpenExcelGlow();
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
        if (!discoveryLoaded || pendingSourceReload) {
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
        if (!shouldReloadDiscovery(selected, discoveryLoaded)) {
            pendingSourceReload = true;
            discoveryLoaded = false;
            return;
        }
        pendingSourceReload = false;
        discoveryLoaded = true;
        if (statusLabel != null && lastBoth == null) {
            statusLabel.setText("検出中…");
        }
        Map<String, String> ui = shell == null ? Map.of() : shell.snapshotUiEnv();
        KouchinPaths paths = KouchinPaths.fromEnv(ui);
        int gen = discoveryGeneration.incrementAndGet();
        Thread t = new Thread(() -> {
            CompletableFuture<List<KouchinDiscovery.Row>> kokubuFut = CompletableFuture.supplyAsync(
                    () -> new ArrayList<>(KouchinDiscovery.scan(FactoryId.KOKUBU, paths)));
            CompletableFuture<List<KouchinDiscovery.Row>> konanFut = CompletableFuture.supplyAsync(
                    () -> new ArrayList<>(KouchinDiscovery.scan(FactoryId.KONAN, paths)));
            List<KouchinDiscovery.Row> kokubu = List.of();
            List<KouchinDiscovery.Row> konan = List.of();
            String error = null;
            try {
                kokubu = kokubuFut.join();
            } catch (CompletionException e) {
                Throwable c = e.getCause() == null ? e : e.getCause();
                error = c.getMessage();
                kokubu = List.of();
            }
            try {
                konan = konanFut.join();
            } catch (CompletionException e) {
                Throwable c = e.getCause() == null ? e : e.getCause();
                String k = c.getMessage();
                error = error == null ? k : error + " / " + k;
                konan = List.of();
            }
            List<KouchinDiscovery.Row> kokubuRows = kokubu;
            List<KouchinDiscovery.Row> konanRows = konan;
            String err = error;
            boolean outputWritable = VerifyOutputAccess.anyOutputWritable(KouchinOutputDirs.resolveAll(ui));
            List<DiscoveryLine> lines = buildDiscoveryLines(kokubuRows, konanRows);
            String kokubuBlock = VerifySourceAccess.blockReason(kokubuRows, accessLookup(lines, "国分"));
            String konanBlock = VerifySourceAccess.blockReason(konanRows, accessLookup(lines, "湖南"));
            Platform.runLater(() -> applyDiscoveryResult(
                    gen, kokubuRows, konanRows, err, outputWritable, lines, kokubuBlock, konanBlock));
        }, "kouchin-discovery");
        t.setDaemon(true);
        t.start();
    }

    private void applyDiscoveryResult(
            int gen,
            List<KouchinDiscovery.Row> kokubu,
            List<KouchinDiscovery.Row> konan,
            String error,
            boolean outputWritable,
            List<DiscoveryLine> lines,
            String kokubuBlock,
            String konanBlock) {
        if (gen != discoveryGeneration.get()) {
            return;
        }
        lastKokubuDiscovery = kokubu == null ? List.of() : List.copyOf(kokubu);
        lastKonanDiscovery = konan == null ? List.of() : List.copyOf(konan);
        lastKokubuBlock = kokubuBlock;
        lastKonanBlock = konanBlock;
        boolean busy = shell != null && shell.isKouchinRunBusy();
        if (!busy) {
            lastOutputWritable = outputWritable;
        }
        if (discoveryTable != null) {
            discoveryTable.getItems().setAll(lines == null ? List.of() : lines);
        }
        refreshTargetYmReason();
        if (statusLabel != null && lastBoth == null) {
            statusLabel.setText(error == null ? discoveryStatusText() : "検出失敗: " + error);
        }
        refreshDropTargetLabel();
        if (!busy) {
            refreshRunEnabled();
        }
    }

    private String discoveryStatusText() {
        StringBuilder sb = new StringBuilder("検出完了。");
        if (!lastOutputWritable) {
            sb.append(" 結果Excelを書き込めません。");
        }
        if (lastKokubuBlock != null) {
            sb.append(" 国分検証不可: ").append(lastKokubuBlock).append("。");
        }
        if (lastKonanBlock != null) {
            sb.append(" 湖南検証不可: ").append(lastKonanBlock).append("。");
        }
        if (lastOutputWritable && lastKokubuBlock == null && lastKonanBlock == null) {
            sb.append(" 検証できます。");
        }
        return sb.toString();
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
    private void onOpenKokubuSource2() {
        openSource2For(FactorySite.KOKUBU);
    }

    @FXML
    private void onOpenKonanSource2() {
        openSource2For(FactorySite.KONAN);
    }

    private void openSource2For(FactorySite site) {
        List<KouchinDiscovery.Row> discovery = site == FactorySite.KONAN ? lastKonanDiscovery : lastKokubuDiscovery;
        VerifyResult verified = lastBoth == null
                ? null
                : (site == FactorySite.KONAN ? lastBoth.konan() : lastBoth.kokubu());
        Path p = source2FileToOpen(verified, discovery);
        if (p == null) {
            appendLog((site == FactorySite.KONAN ? "湖南" : "国分") + "の対象月②がありません");
            return;
        }
        try {
            DesktopFileOpener.openFile(p);
            setStatus("読み書き可で開いた: " + p.getFileName());
        } catch (Exception e) {
            appendLog("②を開けません: " + p.getFileName() + " " + e.getMessage());
        }
    }

    @FXML
    private void onOpenKokubuExcel() {
        openExcelFor(FactorySite.KOKUBU);
    }

    @FXML
    private void onOpenKonanExcel() {
        openExcelFor(FactorySite.KONAN);
    }

    private void openExcelFor(FactorySite site) {
        Path p = VerifyRunSupport.excelFileToOpen(lastWritten, site);
        if (p == null) {
            appendLog((site == FactorySite.KONAN ? "湖南" : "国分") + "のExcelがありません（書込成功側が候補）");
            return;
        }
        try {
            DesktopFileOpener.openFileReadOnly(p);
            setStatus("読み取り専用で開いた: " + p.getFileName());
            stopOpenExcelGlow(site);
        } catch (Exception e) {
            appendLog("Excelを開けません: " + p.getFileName() + " " + e.getMessage());
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
        if (both) {
            if (lastKokubuBlock != null || lastKonanBlock != null) {
                setStatus("関連ファイルにアクセスできないためまとめて検証できません");
                return;
            }
        } else if (one == FactoryId.KONAN) {
            if (lastKonanBlock != null) {
                setStatus("湖南の関連ファイルにアクセスできないため検証できません");
                return;
            }
        } else if (lastKokubuBlock != null) {
            setStatus("国分の関連ファイルにアクセスできないため検証できません");
            return;
        }
        if (!lastOutputWritable) {
            setStatus("結果Excelを書き込めないため検証できません");
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
        Task<VerifyTaskOutcome> task = new Task<>() {
            @Override
            protected VerifyTaskOutcome call() throws Exception {
                FileDiscovery.invalidateListingCache();
                BothResult bothResult;
                if (both) {
                    bothResult = VerifyService.runBoth(paths);
                } else if (one == FactoryId.KONAN) {
                    VerifyResult konan = VerifyService.run(FactoryId.KONAN, paths);
                    bothResult = new BothResult(
                            null, konan, null, null, UnifiedMailBuilder.buildText(null, konan.mail()));
                } else {
                    VerifyResult kokubu = VerifyService.run(FactoryId.KOKUBU, paths);
                    bothResult = new BothResult(
                            kokubu, null, null, null, UnifiedMailBuilder.buildText(kokubu.mail(), null));
                }
                if (cancel.get()) {
                    throw new IllegalStateException("中断されました");
                }
                VerifyRunSupport.Written written = VerifyRunSupport.writeBoth(
                        bothResult.kokubu(), bothResult.konan(), bothResult, ui, cancel);
                return new VerifyTaskOutcome(bothResult, written);
            }
        };
        task.setOnSucceeded(e -> {
            shell.endKouchinRun();
            VerifyTaskOutcome outcome = task.getValue();
            lastBoth = outcome == null ? null : outcome.both();
            lastWritten = outcome == null ? null : outcome.written();
            if (lastBoth != null) {
                if (lastBoth.kokubu() != null) {
                    kokubuVerified = true;
                }
                if (lastBoth.konan() != null) {
                    konanVerified = true;
                }
            }
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
        applySkipBBanner(both);
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
        String skipBanner = skipBBannerText(both);
        if (!skipBanner.isBlank()) {
            st.append(" 【検証Bスキップ】③月次実績に対象月データがありません。");
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
                .append(" A要確認").append(r.requiredCheckA());
        if (r.skippedB()) {
            kpi.append(" Bスキップ");
        } else {
            kpi.append(" B要確認").append(r.requiredCheckB());
        }
        kpi.append(" 報告過不足").append(r.num("報告する過不足"))
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
        if (r.skippedB()) {
            String w = r.warnings().stream()
                    .filter(x -> x.contains("【検証Bスキップ】"))
                    .findFirst()
                    .orElse(r.str("検証Bスキップ理由"));
            allResultLines.add(new ResultLine("警告:" + r.profile().label(), "—",
                    null, null, null, null, null, "検証Bスキップ", w));
        }
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
                ? "" : searchField.getText();
        List<ResultLine> filtered = allResultLines.stream()
                .filter(r -> matchesResultFilter(r, judge, q))
                .collect(Collectors.toList());
        resultTable.getItems().setAll(filtered);
    }

    static boolean matchesResultFilter(ResultLine r, String judge, String qRaw) {
        if (r == null) {
            return false;
        }
        if (judge != null && !"すべて".equals(judge) && !judge.equals(r.getJudge())) {
            return false;
        }
        String q = qRaw == null ? "" : qRaw.trim().toUpperCase(Locale.ROOT);
        if (q.isEmpty()) {
            return true;
        }
        return containsIgnoreCase(r.getKey(), q)
                || containsIgnoreCase(r.getNote(), q)
                || containsIgnoreCase(r.getKind(), q)
                || containsIgnoreCase(r.getJudge(), q);
    }

    private static boolean containsIgnoreCase(String value, String q) {
        return value != null && value.toUpperCase(Locale.ROOT).contains(q);
    }

    private void copyDropped(List<Path> files) {
        if (shell == null) {
            return;
        }
        Path dest = KouchinPaths.fromEnv(shell.snapshotUiEnv()).torayCsvDir();
        Task<KouchinTorayCsvDropSupport.Outcome> task = new Task<>() {
            @Override
            protected KouchinTorayCsvDropSupport.Outcome call() {
                AtomicBoolean cancel = shell.isKouchinRunBusy() ? shell.kouchinCancelRequested() : null;
                return KouchinTorayCsvDropSupport.copyCsvFiles(files, dest, cancel);
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
                markUnverified();
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
        boolean base = !unapplied && !busy && lastOutputWritable;
        boolean kokubuReady = lastKokubuBlock == null;
        boolean konanReady = lastKonanBlock == null;
        if (runKokubuButton != null) {
            runKokubuButton.setDisable(!(base && kokubuReady));
        }
        if (runKonanButton != null) {
            runKonanButton.setDisable(!(base && konanReady));
        }
        if (runBothButton != null) {
            runBothButton.setDisable(!(base && kokubuReady && konanReady));
        }
        if (outputWriteBlockBadge != null) {
            boolean show = !lastOutputWritable && discoveryLoaded;
            outputWriteBlockBadge.setVisible(show);
            outputWriteBlockBadge.setManaged(show);
            if (show) {
                outputWriteBlockBadge.setTooltip(new Tooltip("結果Excelの出力先に書き込めません。Excelで開いているファイルを閉じるか、フォルダ権限を確認してください。"));
            }
        }
        FactorySite site = shell == null ? FactorySite.KOKUBU : shell.currentFactorySite();
        if (runKokubuButton != null) {
            runKokubuButton.setStyle(site == FactorySite.KOKUBU ? "-fx-font-weight: bold;" : "");
        }
        if (runKonanButton != null) {
            runKonanButton.setStyle(site == FactorySite.KONAN ? "-fx-font-weight: bold;" : "");
        }
        refreshRunGlow();
        refreshOpenSource2();
    }

    private void refreshOpenSource2() {
        if (openKokubuSource2Button != null) {
            VerifyResult verified = lastBoth == null ? null : lastBoth.kokubu();
            openKokubuSource2Button.setDisable(source2FileToOpen(verified, lastKokubuDiscovery) == null);
        }
        if (openKonanSource2Button != null) {
            VerifyResult verified = lastBoth == null ? null : lastBoth.konan();
            openKonanSource2Button.setDisable(source2FileToOpen(verified, lastKonanDiscovery) == null);
        }
    }

    private void setupDiscoveryTable() {
        if (discoveryTable == null) {
            return;
        }
        discoveryTable.getColumns().clear();
        TableColumn<DiscoveryLine, Void> openCol = new TableColumn<>("開く");
        openCol.setPrefWidth(64);
        openCol.setMinWidth(56);
        openCol.setMaxWidth(80);
        openCol.setSortable(false);
        openCol.setReorderable(false);
        openCol.setCellFactory(c -> new TableCell<>() {
            private final Button btn = new Button("開く");

            {
                btn.setFocusTraversable(false);
                btn.setMaxHeight(22);
                btn.setStyle("-fx-padding: 1 8 1 8; -fx-font-size: 11px;");
                btn.setOnAction(e -> {
                    DiscoveryLine line = getTableRow() == null ? null : getTableRow().getItem();
                    if (line != null && line.canOpen()) {
                        openDiscoveryRow(line.source());
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
                DiscoveryLine line = getTableRow() == null ? null : getTableRow().getItem();
                boolean on = line != null && line.canOpen();
                btn.setDisable(!on);
                btn.setTooltip(new Tooltip(on ? "ファイルを開く" : "開けるファイルがありません"));
                setGraphic(btn);
            }
        });
        TableColumn<DiscoveryLine, String> factoryCol = new TableColumn<>("工場");
        factoryCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue() == null ? "" : cd.getValue().getFactory()));
        factoryCol.setPrefWidth(128);
        factoryCol.setMinWidth(110);
        TableColumn<DiscoveryLine, String> roleCol = new TableColumn<>("区分");
        roleCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue() == null ? "" : cd.getValue().getRole()));
        roleCol.setPrefWidth(120);
        TableColumn<DiscoveryLine, String> pathCol = new TableColumn<>("ファイル");
        pathCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue() == null ? "" : cd.getValue().getPath()));
        pathCol.setPrefWidth(280);
        pathCol.setMinWidth(140);
        TableColumn<DiscoveryLine, String> modifiedCol = new TableColumn<>("ファイル更新日時");
        modifiedCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue() == null ? "" : cd.getValue().getModifiedAt()));
        modifiedCol.setPrefWidth(150);
        modifiedCol.setMinWidth(130);
        TableColumn<DiscoveryLine, String> ymCol = new TableColumn<>("対象月");
        ymCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue() == null ? "" : cd.getValue().getYm()));
        ymCol.setPrefWidth(110);
        TableColumn<DiscoveryLine, String> noteCol = new TableColumn<>("備考");
        noteCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue() == null ? "" : cd.getValue().getNote()));
        noteCol.setPrefWidth(140);
        discoveryTable.getColumns().add(openCol);
        discoveryTable.getColumns().add(factoryCol);
        discoveryTable.getColumns().add(roleCol);
        discoveryTable.getColumns().add(pathCol);
        discoveryTable.getColumns().add(modifiedCol);
        discoveryTable.getColumns().add(ymCol);
        addAccessColumn("読取", DiscoveryLine::getReadStatus, DiscoveryLine::getReadCss, 64);
        addAccessColumn("書込", DiscoveryLine::getWriteStatus, DiscoveryLine::getWriteCss, 64);
        discoveryTable.getColumns().add(noteCol);
        discoveryTable.setRowFactory(tv -> {
            TableRow<DiscoveryLine> row = new TableRow<>();
            row.setOnMouseClicked(e -> {
                if (e.getClickCount() != 2
                        || e.getButton() != MouseButton.PRIMARY
                        || row.isEmpty()) {
                    return;
                }
                DiscoveryLine line = row.getItem();
                openDiscoveryRow(line == null ? null : line.source());
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
                DiscoveryLine line = empty ? null : getTableRow().getItem();
                getStyleClass().remove("pm-kouchin-missing-path");
                if (line != null && line.isMissing()) {
                    getStyleClass().add(line.getPathCss());
                }
                KouchinDiscovery.Row src = line == null ? null : line.source();
                if (src == null || src.fullPath() == null || src.fullPath().isBlank()) {
                    setTooltip(null);
                } else {
                    tip.setText(src.fullPath().replace('\\', '/'));
                    setTooltip(tip);
                }
            }
        });
    }

    private void addAccessColumn(
            String title,
            Function<DiscoveryLine, String> text,
            Function<DiscoveryLine, String> css,
            int width) {
        TableColumn<DiscoveryLine, String> col = new TableColumn<>(title);
        col.setCellValueFactory(cd -> new SimpleStringProperty(
                cd.getValue() == null ? "" : text.apply(cd.getValue())));
        col.setPrefWidth(width);
        col.setCellFactory(c -> new TableCell<>() {
            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                getStyleClass().removeAll(
                        "pm-kouchin-access-ok",
                        "pm-kouchin-access-ng",
                        "pm-kouchin-access-na",
                        "pm-kouchin-access-warn");
                if (empty) {
                    setText(null);
                    setTooltip(null);
                    return;
                }
                DiscoveryLine line = getTableRow() == null ? null : getTableRow().getItem();
                if (line == null) {
                    setText(item);
                    setTooltip(null);
                    return;
                }
                setText(text.apply(line));
                getStyleClass().add(css.apply(line));
                String tip = "読取".equals(title) ? line.getReadHint() : line.getWriteHint();
                setTooltip(new Tooltip(tip));
            }
        });
        discoveryTable.getColumns().add(col);
    }

    static boolean shouldGlowOpenExcel(VerifyRunSupport.Written written, FactorySite site) {
        return VerifyRunSupport.excelFileToOpen(written, site) != null;
    }

    static String skipBBannerText(BothResult both) {
        if (both == null) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        appendSkipBanner(sb, both.kokubu());
        appendSkipBanner(sb, both.konan());
        return sb.toString().trim();
    }

    private static void appendSkipBanner(StringBuilder sb, VerifyResult r) {
        if (r == null || !r.skippedB()) {
            return;
        }
        if (sb.length() > 0) {
            sb.append("\n\n");
        }
        String warning = r.warnings().stream()
                .filter(w -> w.contains("【検証Bスキップ】"))
                .findFirst()
                .orElse(r.str("検証Bスキップ理由"));
        if (warning == null || warning.isBlank()) {
            warning = "③月次実績に対象月のデータがありません。検証Bは実施していません。";
        }
        sb.append("【").append(r.profile().label()).append("】 ").append(warning);
    }

    private void applySkipBBanner(BothResult both) {
        String text = skipBBannerText(both);
        boolean show = !text.isBlank();
        applySkipBBanner(skipBBanner, text, show);
        applySkipBBanner(skipBResultBanner, text, show);
    }

    private static void applySkipBBanner(Label banner, String text, boolean show) {
        if (banner == null) {
            return;
        }
        banner.setText(show ? text : "");
        banner.setVisible(show);
        banner.setManaged(show);
    }

    static Path source2FileToOpen(VerifyResult verified, List<KouchinDiscovery.Row> discovery) {
        if (verified != null) {
            Path fromVerify = openablePath(verified.str("②パス"));
            if (fromVerify != null) {
                return fromVerify;
            }
        }
        if (discovery != null) {
            for (KouchinDiscovery.Row row : discovery) {
                if (row != null && row.role() != null && row.role().startsWith("②")) {
                    Path found = openableDiscoveryFile(row);
                    if (found != null) {
                        return found;
                    }
                }
            }
        }
        return null;
    }

    static Path openablePath(String full) {
        if (full == null || full.isBlank()) {
            return null;
        }
        Path p = Path.of(full);
        if (!Files.isRegularFile(p)) {
            return null;
        }
        return p.toAbsolutePath().normalize();
    }

    static boolean shouldGlowRunButton(boolean enabled, boolean verified) {
        return enabled && !verified;
    }

    static boolean shouldGlowRunBoth(boolean enabled, boolean kokubuDone, boolean konanDone) {
        return enabled && (!kokubuDone || !konanDone);
    }

    void markUnverified() {
        kokubuVerified = false;
        konanVerified = false;
        refreshRunGlow();
    }

    private void refreshRunGlow() {
        applyRunGlow(runKokubuButton, runKokubuGlow,
                shouldGlowRunButton(enabled(runKokubuButton), kokubuVerified));
        applyRunGlow(runKonanButton, runKonanGlow,
                shouldGlowRunButton(enabled(runKonanButton), konanVerified));
        applyRunGlow(runBothButton, runBothGlow,
                shouldGlowRunBoth(enabled(runBothButton), kokubuVerified, konanVerified));
    }

    private static boolean enabled(Button button) {
        return button != null && !button.isDisable();
    }

    private static void applyRunGlow(Button button, ButtonAttentionGlow glow, boolean on) {
        if (button == null || glow == null) {
            return;
        }
        if (on) {
            glow.startIfIdle();
        } else {
            glow.stop();
        }
    }

    private void refreshOpenExcelGlow() {
        applyOpenExcelGlow(FactorySite.KOKUBU, openKokubuExcelButton);
        applyOpenExcelGlow(FactorySite.KONAN, openKonanExcelButton);
    }

    private void applyOpenExcelGlow(FactorySite site, Button button) {
        if (button == null) {
            return;
        }
        ButtonAttentionGlow glow = site == FactorySite.KONAN ? openKonanExcelGlow : openKokubuExcelGlow;
        if (glow == null) {
            glow = new ButtonAttentionGlow(button);
            if (site == FactorySite.KONAN) {
                openKonanExcelGlow = glow;
            } else {
                openKokubuExcelGlow = glow;
            }
        }
        boolean openable = shouldGlowOpenExcel(lastWritten, site);
        button.setDisable(!openable);
        if (openable) {
            glow.ensureActive();
        } else {
            glow.stop();
        }
    }

    private void stopOpenExcelGlow() {
        ButtonAttentionGlow.stopAll(openKokubuExcelGlow, openKonanExcelGlow);
    }

    private void stopOpenExcelGlow(FactorySite site) {
        if (site == FactorySite.KONAN) {
            ButtonAttentionGlow.stopAll(openKonanExcelGlow);
        } else {
            ButtonAttentionGlow.stopAll(openKokubuExcelGlow);
        }
    }

    static boolean shouldReloadDiscovery(boolean selected, boolean alreadyLoaded) {
        return true;
    }

    static String targetYmReasonText(
            List<KouchinDiscovery.Row> kokubu, List<KouchinDiscovery.Row> konan) {
        KouchinDiscovery.Row toray = firstToray(kokubu);
        if (toray == null) {
            toray = firstToray(konan);
        }
        String rule = "対象月はカレンダーではなく、①東レCSVのファイル名（RVSHEETyyyymm.csv）で決まります。"
                + "フォルダ内で年月が最新のCSVを使い、②③もその月に合わせます。国分と湖南は同じ①です。";
        if (toray == null || toray.missing()) {
            return rule + " いまは該当CSVが見つかりません。";
        }
        String name = toray.fullPath() == null || toray.fullPath().isBlank()
                ? toray.path()
                : Path.of(toray.fullPath()).getFileName().toString();
        String ym = toray.ym() == null || toray.ym().isBlank() ? "年月不明" : toray.ym();
        return rule + " いまは " + name + " → " + ym + "。";
    }

    private static KouchinDiscovery.Row firstToray(List<KouchinDiscovery.Row> rows) {
        if (rows == null) {
            return null;
        }
        for (KouchinDiscovery.Row row : rows) {
            if (row != null && row.role() != null && row.role().startsWith("①")) {
                return row;
            }
        }
        return null;
    }

    private void refreshTargetYmReason() {
        if (targetYmReasonLabel == null) {
            return;
        }
        targetYmReasonLabel.setText(targetYmReasonText(lastKokubuDiscovery, lastKonanDiscovery));
    }

    static Path openableDiscoveryFile(KouchinDiscovery.Row row) {
        if (row == null || row.missing()) {
            return null;
        }
        return openablePath(row.fullPath());
    }

    static String fileModifiedAtText(KouchinDiscovery.Row row) {
        Path p = openableDiscoveryFile(row);
        if (p == null) {
            return "";
        }
        try {
            return FILE_MODIFIED_AT.format(
                    Files.getLastModifiedTime(p).toInstant().atZone(JapanDateTimeDisplay.JST));
        } catch (IOException e) {
            return "";
        }
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
