package jp.co.pm.ai.desktop.dispatch;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.TimeUnit;

/**
 * Runs {@code dispatch_interactive_trial.py} (writes shortages JSON, updates {@code 結果_配台表.json}
 * and exports xlsx).
 */
public final class ResultDispatchTrialPython {

    private ResultDispatchTrialPython() {}

    /**
     * @return path to shortages JSON (stdout line) or empty
     */
    public static String runTrial(Path jsonPath, Path pythonExe, Path pythonScriptDir) throws Exception {
        Path script = pythonScriptDir.resolve("dispatch_interactive_trial.py");
        if (!Files.isRegularFile(script)) {
            throw new IllegalStateException("missing dispatch_interactive_trial.py in " + pythonScriptDir);
        }
        ProcessBuilder pb =
                new ProcessBuilder(
                        pythonExe.toString(),
                        script.toAbsolutePath().toString(),
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
        boolean finished = p.waitFor(600, TimeUnit.SECONDS);
        if (!finished) {
            p.destroyForcibly();
            throw new IllegalStateException("dispatch_interactive_trial.py timeout");
        }
        if (p.exitValue() != 0) {
            throw new IllegalStateException(
                    "dispatch_interactive_trial.py exit " + p.exitValue() + ": " + out);
        }
        return out;
    }
}
