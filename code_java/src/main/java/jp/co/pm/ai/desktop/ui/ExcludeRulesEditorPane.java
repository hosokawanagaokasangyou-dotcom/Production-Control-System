package jp.co.pm.ai.desktop.ui;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Supplier;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.geometry.Insets;
import javafx.scene.Parent;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * Edits {@code PM_AI_EXCLUDE_RULES_JSON} (UTF-8): list of rules or {@code {"rules":[...]}}.
 * Replaces maintenance of the macro sheet ???_?z??s?v?H?? when the JSON path is set and valid.
 */
public final class ExcludeRulesEditorPane {

    private static final ObjectMapper JSON = new ObjectMapper();

    private ExcludeRulesEditorPane() {}

    public static Parent create(Stage owner, Supplier<Map<String, String>> env, Consumer<String> log) {
        TextField pathField = new TextField();
        pathField.setPromptText("PM_AI_EXCLUDE_RULES_JSON \u2014 .json \u30d5\u30eb\u30d1\u30b9");

        TextArea body = new TextArea();
        body.setPromptText(
                "[\n  { \"\u5de5\u7a0b\u540d\": \"...\", \"\u6a5f\u68b0\u540d\": \"...\", ... }\n]\n"
                        + "\u307e\u305f\u306f {\"rules\":[...]}");
        VBox.setVgrow(body, Priority.ALWAYS);

        Button fromEnv = new Button("\u74b0\u5883\u5909\u6570\u304b\u3089\u30d1\u30b9");
        fromEnv.setOnAction(
                e -> {
                    String p = env.get().get(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON);
                    pathField.setText(p != null ? p : "");
                });

        Button pick = new Button("\u53c2\u7167...");
        pick.setOnAction(
                e -> {
                    FileChooser fc = new FileChooser();
                    fc.setTitle("PM_AI_EXCLUDE_RULES_JSON");
                    FileChooserForEnvKey.apply(fc, AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON);
                    String cur = pathField.getText();
                    if (cur != null && !cur.isBlank()) {
                        try {
                            Path p = Path.of(cur.trim());
                            if (Files.isRegularFile(p) && p.getParent() != null) {
                                fc.setInitialDirectory(p.getParent().toFile());
                            }
                        } catch (Exception ignored) {
                        }
                    }
                    var f = fc.showOpenDialog(owner);
                    if (f != null) {
                        pathField.setText(f.getAbsolutePath());
                    }
                });

        Button load = new Button("\u8aad\u8fbc");
        load.setOnAction(
                e -> {
                    String p = pathField.getText() != null ? pathField.getText().trim() : "";
                    if (p.isEmpty()) {
                        log.accept("[exclude-json] path empty");
                        return;
                    }
                    try {
                        String s = Files.readString(Path.of(p), StandardCharsets.UTF_8);
                        body.setText(s);
                        log.accept("[exclude-json] load ok: " + p);
                    } catch (IOException ex) {
                        log.accept("[exclude-json] load error: " + ex.getMessage());
                    }
                });

        Button save = new Button("\u4fdd\u5b58");
        save.setOnAction(
                e -> {
                    String p = pathField.getText() != null ? pathField.getText().trim() : "";
                    if (p.isEmpty()) {
                        log.accept("[exclude-json] path empty (set PM_AI_EXCLUDE_RULES_JSON or type path)");
                        return;
                    }
                    try {
                        Files.writeString(Path.of(p), body.getText(), StandardCharsets.UTF_8);
                        log.accept("[exclude-json] save ok: " + p);
                    } catch (IOException ex) {
                        log.accept("[exclude-json] save error: " + ex.getMessage());
                    }
                });

        Button validate = new Button("JSON \u691c\u8a3c");
        validate.setOnAction(
                e -> {
                    String t = body.getText() != null ? body.getText().trim() : "";
                    if (t.isEmpty()) {
                        log.accept("[exclude-json] body empty");
                        return;
                    }
                    try {
                        JsonNode n = JSON.readTree(t);
                        if (n.isObject() && n.has("rules") && n.get("rules").isArray()) {
                            log.accept(
                                    "[exclude-json] ok: object with rules["
                                            + n.get("rules").size()
                                            + "]");
                        } else if (n.isArray()) {
                            log.accept("[exclude-json] ok: array len=" + n.size());
                        } else {
                            log.accept(
                                    "[exclude-json] parse ok but expected array or {\"rules\":[]}, got: "
                                            + n.getNodeType());
                        }
                    } catch (Exception ex) {
                        log.accept("[exclude-json] invalid: " + ex.getMessage());
                    }
                });

        Label hint =
                new Label(
                        "PM_AI_EXCLUDE_RULES_JSON \u304c\u5b9f\u5728\u304b\u3064\u6709\u52b9\u306a\u3089"
                                + " \u8a2d\u5b9a_\u914d\u53f0\u4e0d\u8981\u5de5\u7a0b \u306e Excel \u4fdd\u5b88\u3092\u7701\u7565\u53ef\u3002"
                                + " \u74b0\u5883\u5909\u6570\u30bf\u30d6\u306b\u540c\u540d\u3092\u8ffd\u52a0\u3057\u3066\u30d1\u30b9\u3092\u5171\u6709\u3057\u3066\u304f\u3060\u3055\u3044\u3002");
        hint.setWrapText(true);

        HBox pathRow = new HBox(8, new Label("\u30d1\u30b9"), pathField, fromEnv, pick);
        pathRow.setSpacing(8);
        HBox actions = new HBox(8, load, save, validate);
        VBox box = new VBox(8, hint, pathRow, actions, body);
        box.setPadding(new Insets(12));
        VBox.setVgrow(body, Priority.ALWAYS);
        BorderPane root = new BorderPane(box);
        return root;
    }
}
