package jp.co.pm.ai.planning.stage2.core;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.planning.stage2.Stage2RunContext;

class Stage2DispatchTrialOrderApplierTest {

    @Test
    void apply_allFromSheet_sortsByColumnThenRow() {
        List<Stage2QueuedTask> in =
                List.of(
                        new Stage2QueuedTask(3, "C", Optional.of(10), 0),
                        new Stage2QueuedTask(2, "A", Optional.of(5), 0),
                        new Stage2QueuedTask(4, "B", Optional.of(5), 0));
        List<String> log = new ArrayList<>();
        Stage2RunContext ctx = new Stage2RunContext(java.util.Map.of(), "", log::add);
        List<Stage2QueuedTask> out = Stage2DispatchTrialOrderApplier.apply(in, ctx);
        assertEquals(3, out.size());
        assertEquals("A", out.get(0).requestId());
        assertEquals(5, out.get(0).dispatchTrialOrderEffective());
        assertEquals("B", out.get(1).requestId());
        assertEquals(5, out.get(1).dispatchTrialOrderEffective());
        assertEquals("C", out.get(2).requestId());
        assertTrue(log.stream().anyMatch(s -> s.contains("配台試行順番")));
    }

    @Test
    void apply_partialSheet_assignsSequentialByExcelRow() {
        List<Stage2QueuedTask> in =
                List.of(
                        new Stage2QueuedTask(5, "X", Optional.of(1), 0),
                        new Stage2QueuedTask(3, "Y", Optional.empty(), 0),
                        new Stage2QueuedTask(4, "Z", Optional.empty(), 0));
        List<String> log = new ArrayList<>();
        Stage2RunContext ctx = new Stage2RunContext(java.util.Map.of(), "", log::add);
        List<Stage2QueuedTask> out = Stage2DispatchTrialOrderApplier.apply(in, ctx);
        assertEquals("Y", out.get(0).requestId());
        assertEquals(1, out.get(0).dispatchTrialOrderEffective());
        assertEquals("Z", out.get(1).requestId());
        assertEquals(2, out.get(1).dispatchTrialOrderEffective());
        assertEquals("X", out.get(2).requestId());
        assertEquals(3, out.get(2).dispatchTrialOrderEffective());
        assertTrue(log.stream().anyMatch(s -> s.contains("暫定")));
    }

    @Test
    void parseDispatchTrialOrder_acceptsDecimalString() {
        assertEquals(Optional.of(2), Stage2DispatchTrialOrderApplier.parseDispatchTrialOrderFromSheet("1.8"));
        assertEquals(Optional.of(7), Stage2DispatchTrialOrderApplier.parseDispatchTrialOrderFromSheet("7"));
        assertEquals(Optional.empty(), Stage2DispatchTrialOrderApplier.parseDispatchTrialOrderFromSheet(""));
    }
}
