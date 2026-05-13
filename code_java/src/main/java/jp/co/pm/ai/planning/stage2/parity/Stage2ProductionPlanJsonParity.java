package jp.co.pm.ai.planning.stage2.parity;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 段階2の JSON（Python / Java いずれのエンジン）をツリー比較する。{@link JsonNode#equals} による構造一致判定。
 */
public final class Stage2ProductionPlanJsonParity {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private Stage2ProductionPlanJsonParity() {}

    /**
     * 両ファイルを UTF-8 JSON として読み、ツリー全体が等しいか判定する。大きいファイルでもストリーム読み。
     */
    public static Stage2ParityCheckResult compareFiles(Path a, Path b) throws IOException {
        JsonNode na;
        JsonNode nb;
        try (InputStream ia = Files.newInputStream(a);
                InputStream ib = Files.newInputStream(b)) {
            na = MAPPER.readTree(ia);
            nb = MAPPER.readTree(ib);
        }
        if (na.equals(nb)) {
            return new Stage2ParityCheckResult(
                    true,
                    "JSON（ツリー全体）が一致しました。\n\nA: "
                            + a
                            + "\nB: "
                            + b
                            + "\n\nformat_version: "
                            + textOrDash(na.get("format_version"))
                            + " / sheets: "
                            + textOrDash(na.get("sheets")));
        }
        String va = textOrDash(na.get("format_version"));
        String vb = textOrDash(nb.get("format_version"));
        return new Stage2ParityCheckResult(
                false,
                "JSON が一致しません（ツリー比較）。\n\nA: "
                        + a
                        + "\n  format_version="
                        + va
                        + "\nB: "
                        + b
                        + "\n  format_version="
                        + vb
                        + "\n\n※ 数値誤差・キー順・省略フィールドの差は別途要確認です。");
    }

    private static String textOrDash(JsonNode n) {
        if (n == null || n.isNull() || n.isMissingNode()) {
            return "—";
        }
        return n.asText();
    }
}
