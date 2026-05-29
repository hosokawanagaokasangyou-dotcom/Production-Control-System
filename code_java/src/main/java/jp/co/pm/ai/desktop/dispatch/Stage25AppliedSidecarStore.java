package jp.co.pm.ai.desktop.dispatch;

import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/** 段階2.5 適用メタ（結果_配台表.json 隣の sidecar）。 */
public final class Stage25AppliedSidecarStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String SIDECAR_SUFFIX = ".stage2_5_applied.json";

    private Stage25AppliedSidecarStore() {}

    public record Stage25Meta(
            boolean stage25Applied,
            String jobId,
            String appliedAt,
            int changedProfileCount,
            String learningArchiveStatus) {

        public static Stage25Meta empty() {
            return new Stage25Meta(false, "", "", 0, "");
        }
    }

    public static Path sidecarPathFor(Path resultDispatchJson) {
        if (resultDispatchJson == null) {
            return null;
        }
        return Path.of(resultDispatchJson.toString() + SIDECAR_SUFFIX);
    }

    public static Stage25Meta tryLoadMeta(Path resultDispatchJson) {
        Path p = sidecarPathFor(resultDispatchJson);
        if (p == null || !Files.isRegularFile(p)) {
            return Stage25Meta.empty();
        }
        try {
            JsonNode root = JSON.readTree(p.toFile());
            return new Stage25Meta(
                    root.path("stage2_5_applied").asBoolean(false),
                    root.path("job_id").asText(""),
                    root.path("applied_at").asText(""),
                    root.path("changed_profile_count").asInt(0),
                    root.path("learning_archive_status").asText(""));
        } catch (Exception ex) {
            return Stage25Meta.empty();
        }
    }
}
