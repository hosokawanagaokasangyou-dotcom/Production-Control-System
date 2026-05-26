package jp.co.pm.ai.desktop.dispatch.rules.stage;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.security.MessageDigest;
import java.time.Instant;
import java.util.HexFormat;
import java.util.Map;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.dispatch.rules.paths.DispatchRulePaths;

/** Capture frozen rule JSON before stage 1-3.5 runs. */
public final class DispatchRuleStageRunOverlay {

    public record CaptureResult(String runId, Path snapshotPath, String capturedAt) {}

    private static final ObjectMapper JSON = new ObjectMapper();

    private DispatchRuleStageRunOverlay() {}

    public static CaptureResult captureForStage(String stage, Map<String, String> ui) throws IOException {
        DispatchRulePaths.ensureWorkJsonFromRepoIfMissing(ui);
        Path work = DispatchRulePaths.workJsonPath(ui);
        if (!Files.isRegularFile(work)) {
            return new CaptureResult("", work, "");
        }
        Path dir = DispatchRulePaths.runSnapshotsDirectory(ui);
        Files.createDirectories(dir);
        byte[] bytes = Files.readAllBytes(work);
        String digest = sha256Prefix(bytes);
        String runId = stage + "_" + Instant.now().toString().replace(":", "").substring(0, 15) + "_" + digest;
        Path target = dir.resolve(runId + ".json");
        Files.copy(work, target, StandardCopyOption.REPLACE_EXISTING);
        appendIndex(dir, runId, stage, digest, target);
        DispatchRuleBuilderRunContext.get().beginRun(stage, runId, target);
        return new CaptureResult(runId, target, Instant.now().toString());
    }

    public static void applySnapshotEnv(Map<String, String> env, Path snapshotPath) {
        if (snapshotPath != null && Files.isRegularFile(snapshotPath)) {
            env.put(DispatchRulePaths.KEY_PM_AI_DISPATCH_SPECIAL_RULES_JSON, snapshotPath.toString());
        }
    }

    private static void appendIndex(Path dir, String runId, String stage, String digest, Path target)
            throws IOException {
        Path indexPath = dir.resolve("index.json");
        ObjectNode index;
        if (Files.isRegularFile(indexPath)) {
            index = (ObjectNode) JSON.readTree(Files.readString(indexPath, StandardCharsets.UTF_8));
        } else {
            index = JSON.createObjectNode();
            index.put("version", 1);
            index.set("entries", JSON.createArrayNode());
        }
        ArrayNode entries = (ArrayNode) index.withArray("entries");
        ObjectNode entry = JSON.createObjectNode();
        entry.put("run_id", runId);
        entry.put("stage", stage);
        entry.put("capturedAt", Instant.now().toString());
        entry.put("sourceHash", digest);
        entry.put("path", target.toString());
        entries.insert(0, entry);
        while (entries.size() > 20) {
            entries.remove(entries.size() - 1);
        }
        Files.writeString(
                indexPath,
                JSON.writerWithDefaultPrettyPrinter().writeValueAsString(index),
                StandardCharsets.UTF_8);
    }

    private static String sha256Prefix(byte[] bytes) {
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] hash = md.digest(bytes);
            return HexFormat.of().formatHex(hash).substring(0, 8);
        } catch (Exception ex) {
            return "00000000";
        }
    }
}
