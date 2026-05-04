package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.control.Accordion;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.StackPane;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane;
import jp.co.pm.ai.desktop.ui.GanttScheduleStyle;
import jp.co.pm.ai.desktop.ui.GanttSheetKind;

/**
 * 「結果_設備ガント」等の時刻軸シートを plan JSON から読み、グラフィック表示する独立タブ。
 */
public final class EquipmentGanttGraphicTabController {

    private static final String DEFAULT_SHEET = "結果_設備ガント";

    private static final String HINT =
            "計画結果ビューアと同じ production_plan_multi_day*.json を指定します。"
                    + " 設備タイムライン（時刻列 HH:MM）と判定されるシートだけコンボに出ます。"
                    + " グラフィック表示はメインのこのタブから利用してください。";

    @FXML
    private Button reloadButton;

    @FXML
    private Button syncLatestButton;

    @FXML
    private TextField planJsonField;

    @FXML
    private Button browsePlanButton;

    @FXML
    private Label statusLabel;

    @FXML
    private Label hintLabel;

    @FXML
    private Accordion sourceAccordion;

    @FXML
    private TitledPane sourceTitledPane;

    @FXML
    private ComboBox<String> sheetCombo;

    @FXML
    private BorderPane contentPane;

    private MainShellController shell;

    private Stage ownerStage;

    private String lastLoadedPlanPath = "";

    @FXML
    private void initialize() {
        if (hintLabel != null) {
            hintLabel.setText(HINT);
        }
        if (sheetCombo != null) {
            sheetCombo.setOnAction(e -> applySelectedSheet());
        }
        if (sourceAccordion != null && sourceTitledPane != null) {
            sourceAccordion.setExpandedPane(sourceTitledPane);
            sourceTitledPane.setExpanded(false);
        }
        if (contentPane != null) {
            contentPane.setCenter(emptyPlaceholder("JSON を指定して再読みしてください。"));
        }
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.ownerStage = shell.getPrimaryStage();
        Platform.runLater(this::reloadFromFields);
    }

    /**
     * 実行タブの計画ブックパスと同じステムの .json があればフィールドに反映し再読み。
     */
    void tryAutofillJsonFromStage2Xlsx(String productionPlanPath, String memberSchedulePath) {
        if (planJsonField == null) {
            return;
        }
        String p = productionPlanPath != null ? productionPlanPath.strip() : "";
        if (p.isEmpty()) {
            return;
        }
        Path json = siblingJson(Path.of(p));
        if (json != null && Files.isRegularFile(json)) {
            planJsonField.setText(json.toString());
            Platform.runLater(this::reloadFromFields);
        }
    }

    @FXML
    private void onReloadButtonAction() {
        reloadFromFields();
    }

    @FXML
    private void onSyncLatestButtonAction() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path dir = AppPaths.defaultPlanningOutputDir(ui);
        try {
            Path plan = newestMatching(dir, "production_plan_multi_day_*.json");
            if (plan != null) {
                planJsonField.setText(plan.toString());
            }
            if (plan == null) {
                statusLabel.setText("このフォルダに production_plan_multi_day*.json がありません: " + dir);
                return;
            }
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return;
        }
        reloadFromFields();
    }

    @FXML
    private void onBrowsePlanJsonAction() {
        FileChooser ch = new FileChooser();
        ch.setTitle("JSON");
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("JSON", "*.json"));
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("All", "*.*"));
        if (shell != null) {
            try {
                Map<String, String> ui = shell.snapshotUiEnv();
                Path dir = AppPaths.defaultPlanningOutputDir(ui);
                if (Files.isDirectory(dir)) {
                    ch.setInitialDirectory(dir.toFile());
                }
            } catch (Exception ignored) {
                // ignore
            }
        }
        java.io.File picked = ch.showOpenDialog(ownerStage);
        if (picked != null) {
            planJsonField.setText(picked.getAbsolutePath());
            reloadFromFields();
        }
    }

    private void reloadFromFields() {
        if (contentPane == null || sheetCombo == null) {
            return;
        }
        reloadButton.setDisable(true);
        syncLatestButton.setDisable(true);
        try {
            String ps = planJsonField != null ? planJsonField.getText().strip() : "";
            Path planPath = ps.isEmpty() ? null : Path.of(ps);
            if (planPath == null || !Files.isRegularFile(planPath)) {
                contentPane.setCenter(emptyPlaceholder("ファイルが指定されていないか、見つかりません。"));
                statusLabel.setText("読み込み対象なし");
                sheetCombo.getItems().clear();
                return;
            }

            Path loadPath = resolvePlanJsonForGraphic(planPath);
            boolean usingLogicalView = !loadPath.equals(planPath);

            Map<String, JsonTableIo.SheetTable> sheets =
                    JsonTableIo.loadSheetsWorkbook(loadPath);
            lastLoadedPlanPath = planPath.toString();

            Map<String, JsonTableIo.SheetTable> eligible = filterEquipmentTimelineSheets(sheets);
            if (eligible.isEmpty()) {
                sheetCombo.getItems().clear();
                contentPane.setCenter(
                        emptyPlaceholder(
                                "設備タイムライン形式のシートが見つかりません（時刻列 HH:MM のシート）。"));
                statusLabel.setText("対象シートなし: " + planPath.getFileName());
                return;
            }

            List<String> names =
                    eligible.keySet().stream().sorted().collect(Collectors.toList());
            String previous =
                    sheetCombo.getSelectionModel().getSelectedItem() != null
                            ? sheetCombo.getSelectionModel().getSelectedItem()
                            : "";
            sheetCombo.getItems().setAll(names);

            String pick = previous;
            if (pick.isEmpty() || !eligible.containsKey(pick)) {
                pick = eligible.containsKey(DEFAULT_SHEET) ? DEFAULT_SHEET : names.get(0);
            }
            sheetCombo.getSelectionModel().select(pick);

            applySelectedSheetFromMap(eligible);
            if (usingLogicalView) {
                statusLabel.setText(
                        "読み込み: "
                                + planPath.getFileName()
                                + " → "
                                + loadPath.getFileName()
                                + " (論理ビュー) / シート数(対象)="
                                + names.size());
            } else {
                statusLabel.setText(
                        "読み込み: " + planPath.getFileName() + " / シート数(対象)=" + names.size());
            }
            if (sourceTitledPane != null) {
                sourceTitledPane.setExpanded(false);
            }
        } catch (Exception ex) {
            contentPane.setCenter(emptyPlaceholder("エラー"));
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            if (shell != null) {
                shell.appendLog("[equipment-gantt-graphic] " + ex.getMessage());
            }
        } finally {
            reloadButton.setDisable(false);
            syncLatestButton.setDisable(false);
        }
    }

    private void applySelectedSheet() {
        String ps = planJsonField != null ? planJsonField.getText().strip() : "";
        Path planPath = ps.isEmpty() ? null : Path.of(ps);
        if (planPath == null || !Files.isRegularFile(planPath)) {
            return;
        }
        if (!planPath.toString().equals(lastLoadedPlanPath)) {
            reloadFromFields();
            return;
        }
        try {
            Path loadPath = resolvePlanJsonForGraphic(planPath);
            Map<String, JsonTableIo.SheetTable> sheets =
                    JsonTableIo.loadSheetsWorkbook(loadPath);
            Map<String, JsonTableIo.SheetTable> eligible = filterEquipmentTimelineSheets(sheets);
            applySelectedSheetFromMap(eligible);
        } catch (IOException ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
        }
    }

    private void applySelectedSheetFromMap(Map<String, JsonTableIo.SheetTable> eligible) {
        String name = sheetCombo.getSelectionModel().getSelectedItem();
        if (name == null || name.isBlank()) {
            return;
        }
        JsonTableIo.SheetTable st = eligible.get(name);
        if (st == null) {
            return;
        }
        EquipmentGraphicGanttPane.agentLogSheetLoad(
                name, st.columns() != null ? st.columns().size() : 0);
        ObservableList<ObservableList<String>> rows = toObservableRows(st);
        contentPane.setCenter(
                EquipmentGraphicGanttPane.build(st.columns(), rows));
    }

    private static Map<String, JsonTableIo.SheetTable> filterEquipmentTimelineSheets(
            Map<String, JsonTableIo.SheetTable> sheets) {
        Map<String, JsonTableIo.SheetTable> out = new LinkedHashMap<>();
        for (Map.Entry<String, JsonTableIo.SheetTable> e : sheets.entrySet()) {
            GanttSheetKind k =
                    GanttScheduleStyle.resolveKind(e.getKey(), e.getValue().columns());
            if (k == GanttSheetKind.EQUIPMENT_TIMELINE) {
                out.put(e.getKey(), e.getValue());
            }
        }
        return out;
    }

    private static ObservableList<ObservableList<String>> toObservableRows(JsonTableIo.SheetTable t) {
        ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
        List<String> cols = t.columns();
        for (Map<String, String> map : t.rows()) {
            ObservableList<String> line = FXCollections.observableArrayList();
            for (String h : cols) {
                line.add(map != null ? map.getOrDefault(h, "") : "");
            }
            rows.add(line);
        }
        return rows;
    }

    private StackPane emptyPlaceholder(String msg) {
        StackPane p = new StackPane(new Label(msg));
        StackPane.setAlignment(p.getChildren().get(0), Pos.CENTER);
        return p;
    }

    /**
     * 同フォルダに {@code <stem>_logical_view.json} があり、引数がミラー用 .json
     * のときは結合展開済みの論理ビューを優先する（設備ガント グラフィック用）。
     */
    private static Path resolvePlanJsonForGraphic(Path planJsonFromField) {
        if (planJsonFromField == null) {
            return null;
        }
        Path fn = planJsonFromField.getFileName();
        if (fn == null) {
            return planJsonFromField;
        }
        String name = fn.toString();
        if (!name.endsWith(".json")) {
            return planJsonFromField;
        }
        String stem = name.substring(0, name.length() - 5);
        if (stem.endsWith("_logical_view")) {
            return planJsonFromField;
        }
        Path sibling = planJsonFromField.resolveSibling(stem + "_logical_view.json");
        if (Files.isRegularFile(sibling)) {
            return sibling;
        }
        return planJsonFromField;
    }

    private static Path siblingJson(Path workbookPath) {
        if (workbookPath == null) {
            return null;
        }
        Path fn = workbookPath.getFileName();
        if (fn == null) {
            return null;
        }
        String name = fn.toString();
        String stem;
        if (name.endsWith(".xlsx")) {
            stem = name.substring(0, name.length() - 5);
        } else if (name.endsWith(".xlsm")) {
            stem = name.substring(0, name.length() - 5);
        } else {
            return null;
        }
        return workbookPath.resolveSibling(stem + ".json");
    }

    private static Path newestMatching(Path dir, String glob) throws IOException {
        if (!Files.isDirectory(dir)) {
            return null;
        }
        Path best = null;
        long bestTime = Long.MIN_VALUE;
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir, glob)) {
            for (Path p : stream) {
                if (!Files.isRegularFile(p)) {
                    continue;
                }
                long t = Files.getLastModifiedTime(p).toMillis();
                if (t >= bestTime) {
                    bestTime = t;
                    best = p;
                }
            }
        }
        return best;
    }
}
