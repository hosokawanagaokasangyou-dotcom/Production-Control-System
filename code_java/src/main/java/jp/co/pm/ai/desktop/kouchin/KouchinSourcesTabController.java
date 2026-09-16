package jp.co.pm.ai.desktop.kouchin;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.stage.DirectoryChooser;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.kouchin.verify.FileDiscovery;
import jp.co.pm.ai.kouchin.verify.JudgmentCsv;
import jp.co.pm.ai.kouchin.verify.KouchinPaths;

/**
 * 参照先。正本は環境変数タブと同じグローバル {@code uiEnvRows}。
 */
public class KouchinSourcesTabController {

    private record FolderSpec(String key, String label, String defaultPath, boolean derivedFromBase) {}

    @FXML private Label judgmentPathLabel;
    @FXML private GridPane folderGrid;
    @FXML private Button applyButton;
    @FXML private Label statusLabel;

    private MainShellController shell;
    private KouchinHostTabController host;
    private final Map<String, TextField> fields = new LinkedHashMap<>();
    private Map<String, String> applied = Map.of();
    private boolean built;

    @FXML
    private void initialize() {
        // POI/UNC は選択時のみ
    }

    public void bindShell(MainShellController shell, KouchinHostTabController host) {
        this.shell = shell;
        this.host = host;
    }

    public void onMainShellTabSelected() {
        ensureGrid();
        if (!hasUnappliedEdits()) {
            loadFromEnv();
        }
        refreshJudgmentLabel();
        refreshStatus();
    }

    public void onMainShellTabDeselected() {}

    public boolean hasUnappliedEdits() {
        if (fields.isEmpty()) {
            return false;
        }
        return !currentValues().equals(applied);
    }

    @FXML
    private void onApply() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path original = AppPaths.resolveRequestFormOriginalDir(ui);
        for (Map.Entry<String, TextField> e : fields.entrySet()) {
            String v = e.getValue().getText() == null ? "" : e.getValue().getText().trim();
            if (v.isEmpty()) {
                continue;
            }
            if (v.toLowerCase().endsWith(".lnk")) {
                setStatus(".lnk は使えません: " + e.getKey());
                return;
            }
            Path p = Path.of(v);
            if (!p.isAbsolute()) {
                setStatus("絶対パスのみ指定してください: " + e.getKey());
                return;
            }
            if (original != null && p.normalize().startsWith(original.toAbsolutePath().normalize())) {
                setStatus("依頼書原本フォルダへは向けられません");
                return;
            }
            shell.setUiEnvRowValue(e.getKey(), v);
        }
        FileDiscovery.invalidateListingCache();
        applied = currentValues();
        setStatus("適用しました（グローバル session-state）");
        refreshJudgmentLabel();
        if (host != null) {
            host.onSourcesApplied();
            if (host.verifyTab() != null) {
                host.verifyTab().refreshRunEnabled();
            }
            if (host.trendTab() != null) {
                host.trendTab().refreshRunEnabled();
            }
        }
    }

    @FXML
    private void onResetUnchanged() {
        for (FolderSpec spec : specs()) {
            TextField tf = fields.get(spec.key());
            if (tf != null) {
                tf.setText(spec.defaultPath());
            }
        }
        refreshStatus();
    }

    @FXML
    private void onAlignFromBase() {
        TextField base = fields.get(AppPaths.KEY_PM_AI_KOUCHIN_BASE_DIR);
        String b = base == null || base.getText() == null || base.getText().isBlank()
                ? AppPaths.DEFAULT_KOUCHIN_BASE_DIR : base.getText().trim();
        setField(AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR, b + "\\東レ送付CSV");
        setField(AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_NAGAOKA_DIR, AppPaths.DEFAULT_KOUCHIN_KOKUBU_NAGAOKA_DIR);
        setField(AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_ALADDIN_DIR, b + "\\国分工場\\アラジン");
        refreshStatus();
    }

    private void ensureGrid() {
        if (built || folderGrid == null) {
            return;
        }
        built = true;
        int r = 0;
        for (FolderSpec spec : specs()) {
            Label lab = new Label(spec.label());
            TextField tf = new TextField();
            tf.setPromptText(spec.defaultPath());
            HBox.setHgrow(tf, Priority.ALWAYS);
            tf.textProperty().addListener((o, a, b) -> refreshStatus());
            Button browse = new Button("参照…");
            browse.setOnAction(ev -> pickDir(spec, tf));
            Button def = new Button("既定に戻す");
            def.setOnAction(ev -> {
                tf.setText(spec.defaultPath());
                refreshStatus();
            });
            folderGrid.add(lab, 0, r);
            folderGrid.add(tf, 1, r);
            folderGrid.add(browse, 2, r);
            folderGrid.add(def, 3, r);
            fields.put(spec.key(), tf);
            if (spec.key().equals(AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR)) {
                tf.setOnDragOver(ev -> {
                    if (ev.getDragboard().hasFiles()) {
                        ev.acceptTransferModes(javafx.scene.input.TransferMode.COPY);
                    }
                    ev.consume();
                });
                tf.setOnDragDropped(ev -> {
                    if (host != null && host.verifyTab() != null && ev.getDragboard().hasFiles()) {
                        host.verifyTab().importCsvPaths(
                                ev.getDragboard().getFiles().stream().map(File::toPath).toList());
                    }
                    ev.setDropCompleted(true);
                    ev.consume();
                });
            }
            r++;
        }
        loadFromEnv();
        applied = currentValues();
    }

    private void pickDir(FolderSpec spec, TextField tf) {
        DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle("フォルダを選択: " + spec.label());
        String cur = tf.getText();
        if (cur != null && !cur.isBlank()) {
            try {
                Path p = Path.of(cur.trim());
                if (Files.isDirectory(p)) {
                    dc.setInitialDirectory(p.toFile());
                }
            } catch (Exception ignored) {
            }
        }
        Window w = folderGrid == null || folderGrid.getScene() == null ? null : folderGrid.getScene().getWindow();
        File f = dc.showDialog(w);
        if (f != null) {
            tf.setText(f.getAbsolutePath());
        }
    }

    private void loadFromEnv() {
        Map<String, String> ui = shell == null ? Map.of() : shell.snapshotUiEnv();
        for (FolderSpec spec : specs()) {
            String v = ui.getOrDefault(spec.key(), "");
            if (v == null || v.isBlank()) {
                v = spec.defaultPath();
            }
            setField(spec.key(), v);
        }
        applied = currentValues();
        refreshJudgmentLabel();
        refreshStatus();
    }

    private void refreshJudgmentLabel() {
        if (judgmentPathLabel == null) {
            return;
        }
        KouchinPaths paths = KouchinPaths.fromEnv(shell == null ? Map.of() : shell.snapshotUiEnv());
        judgmentPathLabel.setText("手動判定.csv: " + JudgmentCsv.manualFile(paths)
                + "  /  前月過不足.csv: " + JudgmentCsv.priorFile(paths)
                + "（既定は国分固定 ●自動検証。BASE に自動追従しません）");
    }

    private void refreshStatus() {
        if (statusLabel == null) {
            return;
        }
        statusLabel.setText(hasUnappliedEdits() ? "未適用の変更があります。適用するまで検証・月次トレンドは実行できません。" : "");
        if (host != null && host.verifyTab() != null) {
            host.verifyTab().refreshRunEnabled();
        }
        if (host != null && host.trendTab() != null) {
            host.trendTab().refreshRunEnabled();
        }
    }

    private Map<String, String> currentValues() {
        Map<String, String> m = new LinkedHashMap<>();
        for (Map.Entry<String, TextField> e : fields.entrySet()) {
            String v = e.getValue().getText() == null ? "" : e.getValue().getText().trim();
            m.put(e.getKey(), v);
        }
        return Map.copyOf(m);
    }

    private void setField(String key, String value) {
        TextField tf = fields.get(key);
        if (tf != null) {
            tf.setText(value == null ? "" : value);
        }
    }

    private void setStatus(String text) {
        if (statusLabel != null) {
            statusLabel.setText(text);
        }
    }

    private static List<FolderSpec> specs() {
        List<FolderSpec> list = new ArrayList<>();
        list.add(new FolderSpec(AppPaths.KEY_PM_AI_KOUCHIN_BASE_DIR, "自動検証ルート",
                AppPaths.DEFAULT_KOUCHIN_BASE_DIR, false));
        list.add(new FolderSpec(AppPaths.KEY_PM_AI_KOUCHIN_TORAY_CSV_DIR, "①東レ送付CSV",
                AppPaths.DEFAULT_KOUCHIN_TORAY_CSV_DIR, true));
        list.add(new FolderSpec(AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_NAGAOKA_DIR, "国分② 長岡後加工賃明細",
                AppPaths.DEFAULT_KOUCHIN_KOKUBU_NAGAOKA_DIR, true));
        list.add(new FolderSpec(AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_ALADDIN_DIR, "国分③ 月次実績",
                AppPaths.DEFAULT_KOUCHIN_KOKUBU_ALADDIN_DIR, true));
        list.add(new FolderSpec(AppPaths.KEY_PM_AI_KOUCHIN_KONAN_SHISAN_DIR, "湖南② 後加工試算（検証と月次トレンドで同一）",
                AppPaths.DEFAULT_KOUCHIN_KONAN_SHISAN_DIR, false));
        list.add(new FolderSpec(AppPaths.KEY_PM_AI_KOUCHIN_KONAN_ALADDIN_DIR, "湖南③ 月次実績",
                AppPaths.DEFAULT_KOUCHIN_KONAN_ALADDIN_DIR, false));
        list.add(new FolderSpec(AppPaths.KEY_PM_AI_KOUCHIN_KONAN_MONTHLY_DIR, "湖南 月次処理ファイル（検証C）",
                AppPaths.DEFAULT_KOUCHIN_KONAN_MONTHLY_DIR, false));
        list.add(new FolderSpec(AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_YEAR_DIRS, "国分 年度フォルダ（トレンド・;区切り）",
                AppPaths.DEFAULT_KOUCHIN_KOKUBU_YEAR_DIRS, false));
        list.add(new FolderSpec(AppPaths.KEY_PM_AI_KOUCHIN_OUTPUT_DIR, "第3コピー先（空なら使わない）", "", false));
        return list;
    }
}
