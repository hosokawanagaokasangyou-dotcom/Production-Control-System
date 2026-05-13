package jp.co.pm.ai.planning.stage2;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.planning.stage2.core.Stage2TaskQueueBuilder;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/** golden/case_minimal の task_queue と Java 構築結果を照合する。 */
class Stage2TaskQueueGoldenTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Test
    void case_minimal_taskQueueMatchesExpectedResource() throws Exception {
        try (var in =
                Stage2TaskQueueGoldenTest.class
                        .getResourceAsStream("/stage2-parity/golden/case_minimal/task_queue_expected.json")) {
            assertTrue(in != null, "classpath golden json");
            JsonNode root = MAPPER.readTree(in.readAllBytes());
            List<String> expectedIds = new ArrayList<>();
            for (JsonNode n : root.get("request_ids")) {
                expectedIds.add(n.asText());
            }
            var tab =
                    new PlanInputTabularIo.TabularSheet(
                            List.of("依頼NO", "工程名"), List.of(List.of("T1", "加工A")));
            Stage2InputSnapshot snap =
                    new Stage2InputSnapshot(
                            java.nio.file.Path.of("master"),
                            List.of("op"),
                            java.util.Optional.empty(),
                            java.util.Optional.empty(),
                            0,
                            java.nio.file.Path.of("plan"),
                            "S",
                            tab);
            List<String> actual = Stage2TaskQueueBuilder.requestIdsInOrder(Stage2TaskQueueBuilder.build(snap));
            assertEquals(expectedIds, actual);
        }
    }
}
