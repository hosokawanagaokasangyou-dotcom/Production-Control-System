package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.ExcelSheetTitlesProbe;
import jp.co.pm.ai.desktop.ipc.IpcStdoutTap;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.embed.swing.SwingNode;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import javax.swing.SwingUtilities;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

/**
 * MVP: ProcessBuilder \u3067\u6bb5\u968e1/2 \u8d77\u52d5\u3001\u30ed\u30b0\u8868\u793a\u3001JFreeChart \u57cb\u3081\u8fbc\u307f\u30b5\u30f3\u30d7\u30eb\u3002
 * bootstrap-javafx / mvp-python-bridge
 */
public class PmAiFxApp extends Application {

    private static final String STAGE1 = "task_extract_stage1.py";
    private static final String STAGE2 = "plan_simulation_stage2.py";
    private static final String PREFIX_CHILD = "[child] ";
    private static final String NDJSON_START = PREFIX_CHILD + "{";

    private TextArea logArea;
    private TextField workbookField;
    private TextField pythonExeField;
    private TextField scriptDirField;
    private Label statusLabel;
    private final AtomicBoolean runLock = new AtomicBoolean(false);

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("\u5de5\u7a0b\u7ba1\u7406 AI \u914d\u53f0 \u2014 JavaFX MVP");

        workbookField = new TextField();
        workbookField.setPromptText("TASK_INPUT_WORKBOOK (.xlsm full path)");

        Button browseWb = new Button("\u53c2\u7167\u2026");
        browseWb.setOnAction(e -> pickWorkbook(primaryStage));

        pythonExeField = new TextField(defaultPythonExecutable());
        pythonExeField.setPromptText("Python executable");

        scriptDirField = new TextField(AppPaths.resolvePythonScriptDir().toString());
        scriptDirField.setPromptText("code/python (PM_AI_CODE_PYTHON_DIR)");

        Button refreshDir = new Button("\u81ea\u52d5\u691c\u51fa");
        refreshDir.setOnAction(e -> scriptDirField.setText(AppPaths.resolvePythonScriptDir().toString()));

        Button peekSheets = new Button("\u30b7\u30fc\u30c8\u4e00\u89a7 (POI)");
        peekSheets.setOnAction(e -> peekSheetsAction());

        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.setPadding(new Insets(12));
        int r = 0;
        grid.add(new Label("\u30bf\u30b9\u30af\u5165\u529b\u30d6\u30c3\u30af"), 0, r);
        grid.add(workbookField, 1, r);
        grid.add(browseWb, 2, r);
        r++;
        grid.add(new Label("Python"), 0, r);
        grid.add(pythonExeField, 1, r);
        r++;
        grid.add(new Label("\u30b9\u30af\u30ea\u30d7\u30c8\u30c7\u30a3\u30ec\u30af\u30c8\u30ea"), 0, r);
        grid.add(scriptDirField, 1, r);
        grid.add(refreshDir, 2, r);
        r++;
        HBox actions = new HBox(8,
                buttonStage(STAGE1, "\u6bb5\u968e1 \u5b9f\u884c"),
                buttonStage(STAGE2, "\u6bb5\u968e2 \u5b9f\u884c"),
                peekSheets);
        grid.add(actions, 1, r);

        logArea = new TextArea();
        logArea.setEditable(false);
        logArea.setWrapText(true);
        VBox.setVgrow(logArea, Priority.ALWAYS);

        statusLabel = new Label(
                "exit: n/a \u2014 0=OK / 1=error / 2=fatal / 3=PlanningValidationError / 9=cancel");

        VBox mainVBox = new VBox(8, grid,
                new Label("log (stdout+stderr merged)"), logArea, statusLabel);
        VBox.setVgrow(logArea, Priority.ALWAYS);
        mainVBox.setPadding(new Insets(0, 12, 12, 12));

        Tab tabMain = new Tab("\u5b9f\u884c\u30fb\u30ed\u30b0", mainVBox);
        tabMain.setClosable(false);

        Tab tabChart = new Tab("\u30b0\u30e9\u30d5 (JFreeChart)", buildChartPane());
        tabChart.setClosable(false);

        Label gridNote = new Label(
                "ControlsFX SpreadsheetView: sheet-ui-parity TODO.");
        Tab tabGridNote = new Tab("\u30c7\u30fc\u30bf\u30b0\u30ea\u30c3\u30c9 (\u4e88\u5b9a)", gridNote);
        tabGridNote.setClosable(false);

        TabPane tabs = new TabPane(tabMain, tabChart, tabGridNote);
        BorderPane root = new BorderPane(tabs);
        root.setStyle("-fx-background-color: #ececec;");
        Scene scene = new Scene(root, 960, 640);
        primaryStage.setMinWidth(640);
        primaryStage.setMinHeight(480);
        primaryStage.setScene(scene);
        primaryStage.show();
        javafx.application.Platform.runLater(() -> {
            primaryStage.toFront();
            primaryStage.requestFocus();
        });

        appendLog("[boot] PYTHONUTF8=1 PYTHONIOENCODING=utf-8 for child process.");
    }

    private Parent buildChartPane() {
        SwingNode swingNode = new SwingNode();
        DefaultCategoryDataset ds = new DefaultCategoryDataset();
        ds.addValue(12, "actual", "M-A");
        ds.addValue(8, "actual", "M-B");
        ds.addValue(15, "plan", "M-A");
        SwingUtilities.invokeLater(() -> {
            JFreeChart chart = ChartFactory.createBarChart(
                    "sample by equipment (JFreeChart)",
                    "equipment",
                    "qty",
                    ds,
                    PlotOrientation.VERTICAL,
                    true,
                    true,
                    false);
            ChartPanel panel = new ChartPanel(chart);
            panel.setFillZoomRectangle(true);
            javafx.application.Platform.runLater(() -> swingNode.setContent(panel));
        });
        BorderPane bp = new BorderPane(swingNode);
        bp.setPadding(new Insets(12));
        bp.setMinHeight(400);
        BorderPane.setMargin(swingNode, new Insets(8));
        return bp;
    }

    private Button buttonStage(String script, String label) {
        Button b = new Button(label);
        b.setOnAction(e -> runStage(script));
        return b;
    }

    private void pickWorkbook(Stage stage) {
        FileChooser ch = new FileChooser();
        ch.setTitle("\u30de\u30af\u30ed\u30d6\u30c3\u30af\u9078\u629e");
        ch.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel", "*.xlsm", "*.xlsx"));
        var f = ch.showOpenDialog(stage);
        if (f != null) {
            workbookField.setText(f.getAbsolutePath());
        }
    }

    private void peekSheetsAction() {
        String p = workbookField.getText() != null ? workbookField.getText().trim() : "";
        if (p.isEmpty()) {
            appendLog("[POI] set TASK_INPUT_WORKBOOK path first.");
            return;
        }
        try {
            var names = ExcelSheetTitlesProbe.sheetNames(Path.of(p));
            appendLog("[POI] sheets=" + names.size() + " " + String.join(", ", names));
        } catch (Exception ex) {
            appendLog("[POI] error: " + ex.getMessage());
        }
    }

    private void runStage(String script) {
        if (!runLock.compareAndSet(false, true)) {
            appendLog("[busy] already running (single flight).");
            return;
        }
        Path py = Path.of(pythonExeField.getText().trim());
        Path dir = Path.of(scriptDirField.getText().trim());
        String wb = workbookField.getText() != null ? workbookField.getText().trim() : "";
        appendLog("--- start: " + script + " ---");
        Map<String, String> extra = new HashMap<>();
        RunRequest req = new RunRequest(py, dir, script, wb, extra);
        statusLabel.setText("running\u2026");

        PythonProcessRunner.runAsync(
                        req,
                        line -> {
                            if (line.startsWith(NDJSON_START)) {
                                String payload = line.substring(PREFIX_CHILD.length());
                                IpcStdoutTap.handleLine(payload, this::appendLog);
                            } else {
                                appendLog(line);
                            }
                        },
                        ex -> appendLog("[error] " + ex.getMessage()))
                .whenComplete((code, err) -> {
                    runLock.set(false);
                    javafx.application.Platform.runLater(() -> {
                        if (err != null) {
                            statusLabel.setText("failed: " + err.getMessage());
                            appendLog("[end] exceptional exit");
                        } else {
                            int c = code != null ? code : -1;
                            statusLabel.setText(exitCodeLegend(c));
                            appendLog("[end] exitCode=" + c + " " + exitHint(c));
                        }
                    });
                });
    }

    private static String exitCodeLegend(int code) {
        return "exit="
                + code
                + " \u2014 0=OK / 1=error / 2=fatal / 3=PlanningValidationError / 9=cancel";
    }

    private static String exitHint(int code) {
        return switch (code) {
            case 0 -> "(success)";
            case 1 -> "(general failure)";
            case 2 -> "(fatal / missing TASK_INPUT / file)";
            case 3 -> "(PlanningValidationError)";
            case 9 -> "(user cancel)";
            default -> "";
        };
    }

    private void appendLog(String line) {
        logArea.appendText(line + "\n");
    }

    private static String defaultPythonExecutable() {
        String env = System.getenv("PM_AI_PYTHON");
        if (env != null && !env.isBlank()) {
            return env;
        }
        return System.getProperty("os.name", "").toLowerCase().contains("win") ? "python" : "python3";
    }

    public static void main(String[] args) {
        System.setProperty("file.encoding", "UTF-8");
        warnIfGuiLikelyUnavailable();
        launch(args);
    }

    /**
     * Linux/WSL without DISPLAY: JavaFX window will not show; warn once on stderr.
     */
    private static void warnIfGuiLikelyUnavailable() {
        String os = System.getProperty("os.name", "").toLowerCase();
        if (!os.contains("linux")) {
            return;
        }
        if (System.getenv("DISPLAY") != null || System.getenv("WAYLAND_DISPLAY") != null) {
            return;
        }
        if (System.getenv("WSL_DISTRO_NAME") != null || System.getenv("WSL_INTEROP") != null) {
            System.err.println(
                    "[PmAiFxApp] DISPLAY/WAYLAND_DISPLAY not set (typical on WSL without WSLg)."
                            + " JavaFX needs a display. Use WSLg, or set DISPLAY for an X server,"
                            + " or run on Windows: code_java\\mvnw.cmd javafx:run");
        }
    }
}
