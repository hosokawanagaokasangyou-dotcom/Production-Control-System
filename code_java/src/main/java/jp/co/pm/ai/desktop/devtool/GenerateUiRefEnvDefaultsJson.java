package jp.co.pm.ai.desktop.devtool;

import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.io.WorkbookEnvSheetReader;

/**
 * One-shot: read {@code plan/UI\u53c2\u7167\u7528_*.xlsx} and write {@code src/main/resources/.../ui_ref_env_defaults.json}.
 *
 * <p>Run from {@code code_java}:
 *
 * <pre>
 * mvn -q compile exec:java -Dexec.mainClass=jp.co.pm.ai.desktop.devtool.GenerateUiRefEnvDefaultsJson
 * </pre>
 */
public final class GenerateUiRefEnvDefaultsJson {

    private static final String UI_REF_FILE =
            "UI\u53c2\u7167\u7528_\u751f\u7523\u7ba1\u7406_AI\u914d\u53f0(RC1).xlsx";

    private GenerateUiRefEnvDefaultsJson() {}

    public static void main(String[] args) throws Exception {
        Path xlsx =
                args.length > 0
                        ? Path.of(args[0])
                        : Path.of("..").resolve("plan").resolve(UI_REF_FILE).normalize();
        if (!Files.isRegularFile(xlsx)) {
            System.err.println("Workbook not found: " + xlsx.toAbsolutePath());
            System.exit(1);
        }
        var rows = WorkbookEnvSheetReader.read(xlsx);
        ObjectMapper om = new ObjectMapper();
        ObjectNode root = om.createObjectNode();
        root.put(
                "source_workbook",
                "Production-Control-System/plan/"
                        + UI_REF_FILE
                        + " (snapshot; regenerate with GenerateUiRefEnvDefaultsJson)");
        root.put("sheet", WorkbookEnvSheetReader.SHEET_NAME);
        ArrayNode arr = om.createArrayNode();
        for (var r : rows) {
            ObjectNode o = om.createObjectNode();
            o.put("key", r.key());
            o.put("value", r.value());
            if (r.description() != null && !r.description().isEmpty()) {
                o.put("description", r.description());
            }
            arr.add(o);
        }
        root.set("entries", arr);
        Path out =
                Path.of("src/main/resources/jp/co/pm/ai/desktop/ui_ref_env_defaults.json")
                        .toAbsolutePath()
                        .normalize();
        Files.createDirectories(out.getParent());
        om.writerWithDefaultPrettyPrinter().writeValue(out.toFile(), root);
        System.out.println("Wrote " + out + " (" + arr.size() + " entries)");
    }
}
