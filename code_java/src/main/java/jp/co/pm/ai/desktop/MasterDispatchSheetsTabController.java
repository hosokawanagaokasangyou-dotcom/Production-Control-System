package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Label;
import javafx.scene.control.TabPane;
import javafx.scene.layout.StackPane;
import javafx.stage.Window;

import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsDocument;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsSeeder;
import jp.co.pm.ai.desktop.io.MasterDispatchSheetsJsonStore;
import jp.co.pm.ai.desktop.ui.MasterDispatchSheetEditRules;
import jp.co.pm.ai.desktop.ui.MasterDispatchSheetGridSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;
import jp.co.pm.ai.desktop.ui.SpreadsheetThemeBridge;

/**
 * メインタブ MASTER。現在工場の skills / need / speed / 組み合わせ表を JSON で表示・編集する。
 */
public final class MasterDispatchSheetsTabController {

    private static final ZoneId TOKYO = ZoneId.of("Asia/Tokyo");

    @FXML private Label statusLabel;
    @FXML private TabPane innerTabPane;
    @FXML private StackPane skillsHost;
    @FXML private StackPane needHost;
    @FXML private StackPane speedHost;
    @FXML private StackPane comboHost;

    private final SpreadsheetView skillsView = new SpreadsheetView();
    private final SpreadsheetView needView = new SpreadsheetView();
    private final SpreadsheetView speedView = new SpreadsheetView();
    private final SpreadsheetView comboView = new SpreadsheetView();

    private MainShellController shell;
    private MasterDispatchSheetsDocument document = MasterDispatchSheetsDocument.empty("");
    private Path loadedJsonPath;

    @FXML
    private void initialize() {
        install(skillsHost, skillsView, 1);
        install(needHost, needView, 3);
        install(speedHost, speedView, 3);
        install(comboHost, comboView, 4);
        applyDocument(document);
    }

    private static void install(StackPane host, SpreadsheetView view, int leadingCols) {
        SpreadsheetThemeBridge.install(view);
        SpreadsheetTabularSupport.installPmAiReadableSpreadsheetChrome(view);
        view.setEditable(true);
        StackPane.setMargin(view, new Insets(0));
        host.getChildren().setAll(view);
        SpreadsheetTabularSupport.installSpreadsheetChromeRelayoutDebouncerForHost(host, () -> leadingCols);
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    void onMainShellTabSelected() {
        reloadFromCurrentFactory(false);
    }

    void reloadFromCurrentFactory(boolean reimport) {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        FactorySite site = AppPaths.currentDispatchFactorySite(ui);
        Path json = AppPaths.masterDispatchSheetsJsonPath(ui);
        Path source = AppPaths.masterDispatchSheetsSourceWorkbookPath(ui);
        try {
            MasterDispatchSheetsSeeder.Result result =
                    MasterDispatchSheetsSeeder.loadOrImport(json, source, site.name(), reimport);
            document = result.document();
            loadedJsonPath = json;
            applyDocument(document);
            statusLabel.setText(statusText(site, json, source, result));
        } catch (Exception e) {
            statusLabel.setText("読込に失敗しました: " + e.getMessage());
        }
    }

    @FXML
    private void onSaveAction() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        FactorySite site = AppPaths.currentDispatchFactorySite(ui);
        Path json = loadedJsonPath != null ? loadedJsonPath : AppPaths.masterDispatchSheetsJsonPath(ui);
        Path source = AppPaths.masterDispatchSheetsSourceWorkbookPath(ui);
        List<List<String>> skillsRows =
                MasterDispatchSheetGridSupport.extract(
                        skillsView, MasterDispatchSheetEditRules.SheetKind.SKILLS);
        List<List<String>> needRows =
                MasterDispatchSheetGridSupport.extract(
                        needView, MasterDispatchSheetEditRules.SheetKind.NEED);
        List<List<String>> speedRows =
                MasterDispatchSheetGridSupport.extract(
                        speedView, MasterDispatchSheetEditRules.SheetKind.SPEED);
        List<List<String>> comboRows =
                MasterDispatchSheetGridSupport.extract(
                        comboView, MasterDispatchSheetEditRules.SheetKind.COMBINATIONS);
        List<String> errors = new java.util.ArrayList<>();
        errors.addAll(
                MasterDispatchSheetEditRules.validateForSave(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, skillsRows));
        errors.addAll(
                MasterDispatchSheetEditRules.validateForSave(
                        MasterDispatchSheetEditRules.SheetKind.NEED, needRows));
        errors.addAll(
                MasterDispatchSheetEditRules.validateForSave(
                        MasterDispatchSheetEditRules.SheetKind.SPEED, speedRows));
        errors.addAll(
                MasterDispatchSheetEditRules.validateForSave(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, comboRows));
        if (!errors.isEmpty()) {
            Alert alert = new Alert(AlertType.ERROR);
            alert.setTitle("MASTER 保存");
            alert.setHeaderText("入力内容を直してから保存してください。");
            alert.setContentText(String.join("\n", errors.subList(0, Math.min(8, errors.size()))));
            Window w = statusLabel != null && statusLabel.getScene() != null ? statusLabel.getScene().getWindow() : null;
            if (w != null) {
                alert.initOwner(w);
            }
            alert.showAndWait();
            statusLabel.setText("保存を中止しました。検証エラーがあります。");
            return;
        }
        LinkedHashMap<String, MasterDispatchSheetsDocument.SheetGrid> sheets = new LinkedHashMap<>();
        sheets.put(
                MasterDispatchSheetsDocument.KEY_SKILLS,
                new MasterDispatchSheetsDocument.SheetGrid("skills", skillsRows));
        sheets.put(
                MasterDispatchSheetsDocument.KEY_NEED,
                new MasterDispatchSheetsDocument.SheetGrid("need", needRows));
        sheets.put(
                MasterDispatchSheetsDocument.KEY_SPEED,
                new MasterDispatchSheetsDocument.SheetGrid("speed", speedRows));
        sheets.put(
                MasterDispatchSheetsDocument.KEY_TEAM_COMBINATIONS,
                new MasterDispatchSheetsDocument.SheetGrid("組み合わせ表", comboRows));
        document =
                new MasterDispatchSheetsDocument(
                        MasterDispatchSheetsDocument.SCHEMA_VERSION,
                        site.name(),
                        source.toString(),
                        OffsetDateTime.now(TOKYO).truncatedTo(ChronoUnit.SECONDS).toString(),
                        sheets);
        try {
            MasterDispatchSheetsJsonStore.write(json, document);
            loadedJsonPath = json;
            statusLabel.setText("保存しました: " + json);
        } catch (Exception e) {
            statusLabel.setText("保存に失敗しました: " + e.getMessage());
        }
    }

    @FXML
    private void onReimportAction() {
        Alert confirm = new Alert(AlertType.CONFIRMATION);
        confirm.setTitle("共有フォルダから再取込");
        confirm.setHeaderText("現在工場の JSON を Excel で上書きします。");
        confirm.setContentText("手編集は失われます。よろしいですか。");
        Window w = statusLabel != null && statusLabel.getScene() != null ? statusLabel.getScene().getWindow() : null;
        if (w != null) {
            confirm.initOwner(w);
        }
        Optional<ButtonType> ans = confirm.showAndWait();
        if (ans.isEmpty() || ans.get() != ButtonType.OK) {
            return;
        }
        reloadFromCurrentFactory(true);
    }

    private void applyDocument(MasterDispatchSheetsDocument doc) {
        MasterDispatchSheetsDocument d = doc != null ? doc : MasterDispatchSheetsDocument.empty("");
        attachGrid(
                skillsView,
                MasterDispatchSheetEditRules.SheetKind.SKILLS,
                d.sheet("skills").rows(),
                0,
                1);
        attachGrid(
                needView,
                MasterDispatchSheetEditRules.SheetKind.NEED,
                d.sheet("need").rows(),
                0,
                3);
        attachGrid(
                speedView,
                MasterDispatchSheetEditRules.SheetKind.SPEED,
                d.sheet("speed").rows(),
                0,
                3);
        attachGrid(
                comboView,
                MasterDispatchSheetEditRules.SheetKind.COMBINATIONS,
                d.sheet("teamCombinations").rows(),
                0,
                4);
    }

    private static void attachGrid(
            SpreadsheetView view,
            MasterDispatchSheetEditRules.SheetKind kind,
            List<List<String>> rows,
            int frozenDataHeaderRows,
            int leadingCols) {
        view.setGrid(MasterDispatchSheetGridSupport.buildEditable(kind, rows));
        int colCount = view.getGrid() != null ? view.getGrid().getColumnCount() : 1;
        List<String> titles = MasterDispatchSheetEditRules.columnTitles(kind, rows, colCount);
        List<Double> widths =
                MasterDispatchSheetEditRules.preferredColumnWidths(rows, colCount, titles);
        Platform.runLater(
                () -> applyMasterSheetChrome(view, widths, frozenDataHeaderRows, leadingCols));
    }

    private static void applyMasterSheetChrome(
            SpreadsheetView view,
            List<Double> widths,
            int frozenDataHeaderRows,
            int leadingCols) {
        SpreadsheetTabularSupport.applyColumnWidths(view, widths, 120);
        SpreadsheetTabularSupport.applyFixedLeadingColumns(view, leadingCols);
        SpreadsheetTabularSupport.applyColumnFiltersWithDialog(view);
        SpreadsheetTabularSupport.pinSpreadsheetFilterRow(view);
        if (frozenDataHeaderRows > 0) {
            SpreadsheetTabularSupport.pinSpreadsheetRows(view, 1, frozenDataHeaderRows);
        }
        SpreadsheetTabularSupport.applyUnconstrainedColumnResizePolicyAfterSkinSettles(view);
        Platform.runLater(() -> SpreadsheetTabularSupport.applyColumnWidths(view, widths, 120));
    }

    private static String statusText(
            FactorySite site, Path json, Path source, MasterDispatchSheetsSeeder.Result result) {
        String factory = site != null ? site.displayLabelJa() : "";
        String outcome =
                switch (result.outcome()) {
                    case LOADED_EXISTING -> "既存 JSON を読み込みました。";
                    case IMPORTED -> "共有フォルダの master から吸い出して JSON を作成しました。";
                    case EMPTY_MISSING_SOURCE -> "吸い出し元の master が見つかりません。空の表を表示しています。";
                };
        return factory
                + "  JSON: "
                + json
                + "\n吸い出し元: "
                + source
                + "\n"
                + outcome;
    }
}
