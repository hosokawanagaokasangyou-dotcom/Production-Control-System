package jp.co.pm.ai.planning.stage2.parity;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.MissingNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * 計画ブック JSON のゴールデン比較用に、実行ごとに変わるトップレベル項目を落とす。
 *
 * <p>Python 正本との厳密一致は段階的に拡張する。現状は {@code source_xlsx}（ファイル名にタイムスタンプが含まれる）や
 * Java 側のメタ {@code engine_note} を除去してシート内容の比較に集中する。
 */
public final class Stage2PlanJsonGoldenNormalizer {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private Stage2PlanJsonGoldenNormalizer() {}

    /** トップレベルから比較に不要なキーを削除したコピーを返す（{@code sheets} はそのまま）。 */
    public static JsonNode stripTopLevelVolatile(JsonNode root) {
        if (root == null || !root.isObject()) {
            return root;
        }
        ObjectNode o = MAPPER.createObjectNode();
        root.fields().forEachRemaining(e -> {
            String k = e.getKey();
            if ("source_xlsx".equals(k) || "engine_note".equals(k)) {
                return;
            }
            o.set(k, e.getValue().deepCopy());
        });
        return o;
    }

    /** {@code sheets} 配下の 1 シートだけを取り出す。無ければ missing node。 */
    public static JsonNode sheetSubtree(JsonNode planRoot, String sheetName) {
        if (planRoot == null || !planRoot.isObject()) {
            return MissingNode.getInstance();
        }
        JsonNode sheets = planRoot.get("sheets");
        if (sheets == null || !sheets.isObject()) {
            return MissingNode.getInstance();
        }
        return sheets.path(sheetName);
    }
}
