package jp.co.pm.ai.planning.stage2.core;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class Stage2TaskQueueBuilderTest {

    @Test
    void findRequestIdColumn_prefersExactHeader() {
        List<String> headers = List.of("工程名", "依頼NO", "備考");
        assertEquals(1, Stage2TaskQueueBuilder.findRequestIdColumn(headers));
    }

    @Test
    void findProcessNameColumn_detectsKoumei() {
        List<String> headers = List.of("依頼NO", "工程名");
        assertEquals(1, Stage2TaskQueueBuilder.findProcessNameColumn(headers));
    }

    @Test
    void build_skipsRowWhenProcessNameEmpty() {
        var tab =
                new jp.co.pm.ai.desktop.io.PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO", "工程名"),
                        List.of(
                                List.of("T1", "加工A"),
                                List.of("T2", ""),
                                List.of("T3", "加工C")));
        var snap =
                new jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot(
                        java.nio.file.Path.of("m"),
                        List.of("m1"),
                        java.util.Optional.empty(),
                        java.util.Optional.empty(),
                        0,
                        java.nio.file.Path.of("p"),
                        "S",
                        tab);
        List<Stage2QueuedTask> q = Stage2TaskQueueBuilder.build(snap);
        assertEquals(2, q.size());
        assertEquals("T1", q.get(0).requestId());
        assertEquals("T3", q.get(1).requestId());
    }

    @Test
    void build_deduplicatesRequestIds() {
        var tab =
                new jp.co.pm.ai.desktop.io.PlanInputTabularIo.TabularSheet(
                        List.of("依頼NO", "工程名"),
                        List.of(
                                List.of("A1", "x"),
                                List.of("A1", "y"),
                                List.of("B2", "z")));
        var snap =
                new jp.co.pm.ai.planning.stage2.input.Stage2InputSnapshot(
                        java.nio.file.Path.of("m"),
                        List.of("m1"),
                        java.util.Optional.empty(),
                        java.util.Optional.empty(),
                        0,
                        java.nio.file.Path.of("p"),
                        "S",
                        tab);
        List<Stage2QueuedTask> q = Stage2TaskQueueBuilder.build(snap);
        assertEquals(2, q.size());
        assertEquals("A1", q.get(0).requestId());
        assertEquals("B2", q.get(1).requestId());
    }
}
