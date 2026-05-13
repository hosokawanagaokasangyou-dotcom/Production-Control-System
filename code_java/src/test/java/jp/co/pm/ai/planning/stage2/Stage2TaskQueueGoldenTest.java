package jp.co.pm.ai.planning.stage2;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.planning.stage2.core.Stage2TaskQueueBuilder;
import jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot;

/** {@code golden/<case_id>} の task_queue と Java 構築結果を照合する。 */
class Stage2TaskQueueGoldenTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    static Stream<Arguments> goldenCases() {
        return Stream.of(
                Arguments.of(
                        "case_minimal",
                        List.of("依頼NO", "工程名"),
                        List.of(List.of("T1", "加工A"))),
                Arguments.of(
                        "case_two_tasks",
                        List.of("依頼NO", "工程名"),
                        List.of(List.of("T1", "加工A"), List.of("T2", "加工B"))),
                Arguments.of(
                        "case_skip_empty_process",
                        List.of("依頼NO", "工程名"),
                        List.of(
                                List.of("T1", "加工A"),
                                List.of("T2", ""),
                                List.of("T3", "加工C"))));
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("goldenCases")
    void taskQueueMatchesGoldenResource(String caseId, List<String> headers, List<List<String>> rows)
            throws Exception {
        String resource = "/stage2-parity/golden/" + caseId + "/task_queue_expected.json";
        try (InputStream in = Stage2TaskQueueGoldenTest.class.getResourceAsStream(resource)) {
            assertTrue(in != null, "classpath golden: " + resource);
            JsonNode root = MAPPER.readTree(in.readAllBytes());
            List<String> expectedIds = new ArrayList<>();
            for (JsonNode n : root.get("request_ids")) {
                expectedIds.add(n.asText());
            }
            var tab = new PlanInputTabularIo.TabularSheet(headers, rows);
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
            assertEquals(expectedIds, actual, "case=" + caseId);
        }
    }
}
