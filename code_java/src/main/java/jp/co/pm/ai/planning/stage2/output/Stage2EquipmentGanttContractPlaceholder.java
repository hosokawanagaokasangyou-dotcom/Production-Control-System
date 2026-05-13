package jp.co.pm.ai.planning.stage2.output;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * Java 段階2のプレースホルダ用に、Python が出力する設備ガント契約 JSON（計画 stem に続く「設」の兄弟ファイル）と同型の最小オブジェクトを書く。
 *
 * <p>{@code kwargs_packed.timeline_events} は空配列。実配台イベントは Python 経路と一致しない。UI 側は表の組み立てまで可能だが
 * タイムライン帯は描画されない（{@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane} の既知挙動）。
 */
public final class Stage2EquipmentGanttContractPlaceholder {

    private static final ObjectMapper JSON = new ObjectMapper();

    private Stage2EquipmentGanttContractPlaceholder() {}

    public static void write(Path contractPath, LocalDate sortedDatesAnchor) throws IOException {
        Files.createDirectories(contractPath.getParent());
        ObjectNode root = JSON.createObjectNode();
        root.put("schema_version", 1);
        root.put("kind", "equipment_gantt");
        root.put("fn", "_write_results_equipment_gantt_sheet");
        ObjectNode packed = root.putObject("kwargs_packed");
        packed.set("timeline_events", JSON.createArrayNode());
        ArrayNode equip = JSON.createArrayNode();
        equip.add("配台未実施+Java ステージ2");
        packed.set("equipment_list", equip);
        ArrayNode dates = JSON.createArrayNode();
        ObjectNode d = JSON.createObjectNode();
        d.put("__t", "date");
        d.put("v", sortedDatesAnchor.toString());
        dates.add(d);
        packed.set("sorted_dates", dates);
        Files.writeString(
                contractPath,
                JSON.writerWithDefaultPrettyPrinter().writeValueAsString(root),
                StandardCharsets.UTF_8);
    }
}
