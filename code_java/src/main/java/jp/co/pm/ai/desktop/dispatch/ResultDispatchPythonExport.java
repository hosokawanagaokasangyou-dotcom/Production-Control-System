package jp.co.pm.ai.desktop.dispatch;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.TimeUnit;

/**
 * Runs {@code export_result_dispatch_from_json.py} after JSON write-back (same xlsx layout as stage2).
 */
public final class ResultDispatchPythonExport {

    private ResultDispatchPythonExport() {}

    /**
     * @return stdout line (xlsx path) or empty if failed
     */
    public static String exportXlsxNearJson(Path jsonPath, Path pythonExe, Path pythonScriptDir) throws Exception {
        Path script = pythonScriptDir.resolve("export_result_dispatch_from_json.py");
        if (!Files.isRegularFile(script)) {
            return "";
        }
        ProcessBuilder pb =
                new ProcessBuilder(
                        pythonExe.toString(),
                        script.toString(),
                        jsonPath.toAbsolutePath().toString());
        pb.directory(pythonScriptDir.toFile());
        pb.redirectErrorStream(true);
        Process p = pb.start();
        String out;
        try (BufferedReader br =
                new BufferedReader(new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = br.readLine()) != null) {
                if (!sb.isEmpty()) {
                    sb.append('\n');
                }
                sb.append(line);
            }
            out = sb.toString().trim();
        }
        boolean finished = p.waitFor(180, TimeUnit.SECONDS);
        if (!finished) {
            p.destroyForcibly();
            return "";
        }
        if (p.exitValue() != 0) {
            return "";
        }
        return out;
    }
}
