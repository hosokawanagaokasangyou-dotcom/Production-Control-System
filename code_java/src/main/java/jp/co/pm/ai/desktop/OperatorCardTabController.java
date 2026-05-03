package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ComboBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.layout.StackPane;
import javafx.scene.text.Font;
import javafx.print.PageLayout;
import javafx.print.PageOrientation;
import javafx.print.Paper;
import javafx.print.Printer;
import javafx.print.PrinterJob;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.io.JsonTableIo.SheetTable;
import jp.co.pm.ai.desktop.print.OperatorCardDocumentBuilder;
import jp.co.pm.ai.desktop.print.OperatorCardDocumentBuilder.OperatorCardBuildException;
import jp.co.pm.ai.desktop.print.OperatorCardPage;
import jp.co.pm.ai.desktop.print.OperatorCardPreviewFactory;

/**
 * Operator dispatch card (A4 preview / print) tab ? layout {@code OperatorCardTab.fxml}.
 */
public final class OperatorCardTabController {

    @FXML
    private Button syncLatestButton;

    @FXML
    private Button refreshPreviewButton;

    @FXML
    private Button printButton;

    @FXML
    private TextField memberJsonField;

    @FXML
    private TextField dispatchJsonField;

    @FXML
    private Button browseMemberButton;

    @FXML
    private Button browseDispatchButton;

    @FXML
    private DatePicker startDatePicker;

    @FXML
    private ComboBox<String> operatorCombo;

    @FXML
    private ComboBox<String> fontCombo;

    @FXML
    private CheckBox printAllOperatorsCheckBox;

    @FXML
    private Label statusLabel;

    @FXML
    private StackPane previewHost;

    private MainShellController shell;

    private Stage ownerStage;

    private Map<String, SheetTable> cachedMemberSheets = Map.of();

    @FXML
    private void initialize() {
        if (startDatePicker != null) {
            startDatePicker.setValue(LocalDate.now());
        }
        if (previewHost != null) {
            previewHost.setAlignment(Pos.TOP_CENTER);
            Label placeholder =
                    new Label(
                            "member_schedule*.json \u3068\u7d50\u679c_\u914d\u53f0\u8868.json"
                                    + " \u3092\u6307\u5b9a\u3057\u3001\u30d7\u30ec\u30d3\u30e5\u30fc\u66f4\u65b0\u3092\u62bc\u3057\u3066\u304f\u3060\u3055\u3044\u3002");
            previewHost.getChildren().setAll(placeholder);
        }
    }

    void bindShell(MainShellController shell) {
        this.shell = shell;
        this.ownerStage = shell.getPrimaryStage();
        populateFontCombo();
        Platform.runLater(this::applyDefaultPathsFromEnv);
    }

    private void applyDefaultPathsFromEnv() {
        if (shell == null || dispatchJsonField == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path disp = AppPaths.resolveResultDispatchTableJsonPath(ui);
        dispatchJsonField.setPromptText(disp.toString());
        if (dispatchJsonField.getText() == null || dispatchJsonField.getText().isBlank()) {
            if (Files.isRegularFile(disp)) {
                dispatchJsonField.setText(disp.toString());
            }
        }
    }

    private void populateFontCombo() {
        if (fontCombo == null) {
            return;
        }
        ObservableList<String> fam = FXCollections.observableArrayList(Font.getFamilies());
        FXCollections.sort(fam);
        fontCombo.setItems(fam);
        fontCombo.setValue(pickDefaultFont(fam));
    }

    static String pickDefaultFont(ObservableList<String> families) {
        if (families == null || families.isEmpty()) {
            return "SansSerif";
        }
        List<String> prefer =
                List.of(
                        "BIZ UD\u30b4\u30b7\u30c3\u30af",
                        "BIZ UD Gothic",
                        "BIZ UDP\u30b4\u30b7\u30c3\u30af",
                        "BIZ UDPGothic");
        for (String p : prefer) {
            if (families.contains(p)) {
                return p;
            }
        }
        if (families.contains("Meiryo UI")) {
            return "Meiryo UI";
        }
        return families.get(0);
    }

    @FXML
    private void onBrowseMemberJsonAction() {
        browseJson(memberJsonField);
    }

    @FXML
    private void onBrowseDispatchJsonAction() {
        browseJson(dispatchJsonField);
    }

    private void browseJson(TextField target) {
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
            target.setText(picked.getAbsolutePath());
            reloadMemberCachesAndOperators();
        }
    }

    @FXML
    private void onSyncLatestButtonAction() {
        if (shell == null) {
            return;
        }
        Map<String, String> ui = shell.snapshotUiEnv();
        Path dir = AppPaths.defaultPlanningOutputDir(ui);
        try {
            Path mem = newestMatching(dir, "member_schedule_*.json");
            Path dispDirFile = AppPaths.resolveResultDispatchTableJsonPath(ui);
            if (mem != null) {
                memberJsonField.setText(mem.toString());
            }
            if (Files.isRegularFile(dispDirFile)) {
                dispatchJsonField.setText(dispDirFile.toString());
            }
            if (mem == null && !Files.isRegularFile(dispDirFile)) {
                statusLabel.setText(
                        "\u6700\u65b0 JSON \u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093: " + dir);
                return;
            }
            statusLabel.setText(
                    "sync: member="
                            + (mem != null ? mem.getFileName() : "-")
                            + ", dispatch="
                            + (Files.isRegularFile(dispDirFile) ? dispDirFile.getFileName() : "-"));
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            return;
        }
        reloadMemberCachesAndOperators();
    }

    @FXML
    private void onRefreshPreviewButtonAction() {
        rebuildPreview();
    }

    @FXML
    private void onPrintButtonAction() {
        printCards();
    }

    /** Mirrors stage-2 artifact refresh: fill {@code member_schedule*.json} sibling path when possible. */
    void tryAutofillMemberJsonFromStage2(String memberSchedulePath) {
        if (memberJsonField == null) {
            return;
        }
        String m = memberSchedulePath != null ? memberSchedulePath.strip() : "";
        if (m.isEmpty()) {
            return;
        }
        Path json = siblingJson(Path.of(m));
        if (json != null && Files.isRegularFile(json)) {
            memberJsonField.setText(json.toString());
            reloadMemberCachesAndOperators();
        }
    }

    private static Path siblingJson(Path workbookPath) {
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
        } else if (name.endsWith(".json")) {
            return workbookPath;
        } else {
            return null;
        }
        return workbookPath.resolveSibling(stem + ".json");
    }

    private void reloadMemberCachesAndOperators() {
        cachedMemberSheets = Map.of();
        String mp = memberJsonField != null ? memberJsonField.getText().strip() : "";
        if (mp.isEmpty()) {
            if (operatorCombo != null) {
                operatorCombo.getItems().clear();
            }
            return;
        }
        Path p = Path.of(mp);
        if (!Files.isRegularFile(p)) {
            statusLabel.setText("\u30d5\u30a1\u30a4\u30eb\u306a\u3057: " + p);
            return;
        }
        try {
            cachedMemberSheets = JsonTableIo.loadSheetsWorkbook(p);
            List<String> ops = JsonTableIo.memberOperatorNames(cachedMemberSheets);
            if (operatorCombo != null) {
                String prev = operatorCombo.getValue();
                operatorCombo.getItems().setAll(ops);
                if (prev != null && ops.contains(prev)) {
                    operatorCombo.setValue(prev);
                } else if (!ops.isEmpty()) {
                    operatorCombo.setValue(ops.get(0));
                }
            }
            statusLabel.setText(
                    "\u8aad\u307f\u8fbc\u307f: \u30aa\u30da\u30ec\u30fc\u30bf\u30fc " + ops.size() + " \u540d");
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            cachedMemberSheets = Map.of();
        }
    }

    private void rebuildPreview() {
        if (previewHost == null) {
            return;
        }
        try {
            OperatorCardPage page = buildSelectedPage();
            String font = fontCombo != null ? fontCombo.getValue() : "SansSerif";
            Parent root = OperatorCardPreviewFactory.buildRoot(page, font);
            ScrollPane sp = new ScrollPane(root);
            sp.setFitToWidth(true);
            sp.setPannable(true);
            previewHost.getChildren().setAll(sp);
            statusLabel.setText(
                    "\u30d7\u30ec\u30d3\u30e5\u30fc: " + page.operatorName() + " / " + page.days().size() + " \u65e5\u5206");
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            Label err = new Label(String.valueOf(ex.getMessage()));
            previewHost.getChildren().setAll(err);
        }
    }

    private OperatorCardPage buildSelectedPage() throws IOException, OperatorCardBuildException {
        List<Map<String, String>> dispatchRows = loadDispatchRows();
        LocalDate start = startDatePicker != null ? startDatePicker.getValue() : LocalDate.now();
        if (start == null) {
            throw new OperatorCardBuildException("start date is null");
        }
        String op = operatorCombo != null ? operatorCombo.getValue() : null;
        if (op == null || op.isBlank()) {
            throw new OperatorCardBuildException("select operator");
        }
        return OperatorCardDocumentBuilder.buildPage(
                op, cachedMemberSheets, dispatchRows, start);
    }

    private List<Map<String, String>> loadDispatchRows() throws IOException {
        String dp = dispatchJsonField != null ? dispatchJsonField.getText().strip() : "";
        if (dp.isEmpty()) {
            throw new IOException("results dispatch JSON path is empty");
        }
        Path p = Path.of(dp);
        if (!Files.isRegularFile(p)) {
            throw new IOException("dispatch file not found: " + p);
        }
        SheetTable t = JsonTableIo.loadFlatTable(p);
        List<Map<String, String>> rows = new ArrayList<>();
        for (Map<String, String> r : t.rows()) {
            rows.add(r);
        }
        return rows;
    }

    private void printCards() {
        if (shell == null) {
            return;
        }
        reloadMemberCachesAndOperators();
        if (cachedMemberSheets.isEmpty()) {
            statusLabel.setText("member_schedule JSON \u3092\u8aad\u307f\u8fbc\u3081\u307e\u305b\u3093");
            return;
        }
        List<Map<String, String>> dispatchRows;
        try {
            dispatchRows = loadDispatchRows();
        } catch (IOException ex) {
            statusLabel.setText(ex.getMessage());
            return;
        }
        LocalDate start = startDatePicker != null ? startDatePicker.getValue() : LocalDate.now();
        if (start == null) {
            statusLabel.setText("\u958b\u59cb\u65e5\u3092\u8a2d\u5b9a\u3057\u3066\u304f\u3060\u3055\u3044");
            return;
        }
        List<String> operators = new ArrayList<>();
        boolean all =
                printAllOperatorsCheckBox != null && printAllOperatorsCheckBox.isSelected();
        if (all) {
            operators.addAll(JsonTableIo.memberOperatorNames(cachedMemberSheets));
        } else {
            String op = operatorCombo != null ? operatorCombo.getValue() : null;
            if (op == null || op.isBlank()) {
                statusLabel.setText("\u30aa\u30da\u30ec\u30fc\u30bf\u30fc\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044");
                return;
            }
            operators.add(op);
        }
        if (operators.isEmpty()) {
            statusLabel.setText("\u5370\u5237\u5bfe\u8c61\u304c\u3042\u308a\u307e\u305b\u3093");
            return;
        }

        String font = fontCombo != null ? fontCombo.getValue() : "SansSerif";

        PrinterJob job = PrinterJob.createPrinterJob();
        if (!job.showPrintDialog(ownerStage)) {
            return;
        }
        Printer printer = job.getPrinter();
        PageLayout layout =
                printer.createPageLayout(
                        Paper.A4, PageOrientation.PORTRAIT, Printer.MarginType.DEFAULT);

        try {
            for (String opName : operators) {
                OperatorCardPage page =
                        OperatorCardDocumentBuilder.buildPage(
                                opName, cachedMemberSheets, dispatchRows, start);
                Parent root = OperatorCardPreviewFactory.buildRoot(page, font);
                boolean ok = job.printPage(layout, root);
                if (!ok) {
                    shell.appendLog("[operator-card] printPage returned false for " + opName);
                    break;
                }
            }
        } catch (Exception ex) {
            statusLabel.setText(ex.getMessage() != null ? ex.getMessage() : ex.toString());
            shell.appendLog("[operator-card] " + ex.getMessage());
            return;
        } finally {
            job.endJob();
        }
        statusLabel.setText(
                "\u5370\u5237\u5b8c\u4e86: " + operators.size() + " \u679a\u8a18\u4e8b");
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
