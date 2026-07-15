package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;

import org.junit.jupiter.api.Test;

class LimitedOperatorEditContextTest {

    @Test
    void fromRowReadsProcessAndMachineForBothPlanInputTables() {
        LimitedOperatorEditContext context =
                LimitedOperatorEditContext.fromRow(
                        List.of("依頼NO", "工程名", "担当OP_限定", "機械名"),
                        List.of("T1", "ラミネート", "", "1号機"));

        assertEquals("ラミネート", context.processName());
        assertEquals("1号機", context.machineName());
        context.validateComplete();
    }

    @Test
    void validateCompleteRejectsMissingProcessOrMachine() {
        LimitedOperatorEditContext context =
                LimitedOperatorEditContext.fromRow(
                        List.of("工程名", "機械名"), List.of("ラミネート", ""));

        assertThrows(IllegalArgumentException.class, context::validateComplete);
    }
}
