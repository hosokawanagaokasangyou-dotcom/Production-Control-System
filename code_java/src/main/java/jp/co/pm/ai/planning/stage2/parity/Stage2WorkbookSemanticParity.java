package jp.co.pm.ai.planning.stage2.parity;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.planning.stage2.output.Stage2WorkbookJsonWriter;

/**
 * 段階2の計画／人員 xlsx を {@link Stage2WorkbookJsonWriter} 経由の正規表現に揃え、ファイル名以外をツリー比較する。
 */
public final class Stage2WorkbookSemanticParity {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private Stage2WorkbookSemanticParity() {}

    public static Stage2ParityCheckResult compareXlsx(Path a, Path b) throws IOException {
        Map<String, Object> pa = Stage2WorkbookJsonWriter.buildPayloadFromXlsx(a, Map.of());
        Map<String, Object> pb = Stage2WorkbookJsonWriter.buildPayloadFromXlsx(b, Map.of());
        pa.remove("source_xlsx");
        pb.remove("source_xlsx");
        JsonNode ja = MAPPER.valueToTree(pa);
        JsonNode jb = MAPPER.valueToTree(pb);
        if (ja.equals(jb)) {
            return new Stage2ParityCheckResult(
                    true,
                    "ブック xlsx のセル内容（シート名・列・行）が一致しました。\n\nA: "
                            + a
                            + "\nB: "
                            + b);
        }
        return new Stage2ParityCheckResult(
                false,
                "ブック xlsx のセル内容が一致しません（DataFormatter 正規化後のツリー比較）。\n\nA: "
                        + a
                        + "\nB: "
                        + b
                        + "\n\n※ 書式のみの差は無視できません。Python/Java でシート構成や値が異なる可能性があります。");
    }
}
