package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class StageRunLogProgressParserTest {

    @Test
    void extractDetail_stripsChildPrefixAndTimestamp() {
        String line =
                "[child] 2026-08-08 07:45:26,785 INFO 段階2: PowerQuery で配合表を生成しています";
        assertEquals(
                "段階2: PowerQuery で配合表を生成しています",
                StageRunLogProgressParser.extractDetail(line).orElseThrow());
    }

    @Test
    void extractDetail_parsesPmAiProgress() {
        assertEquals(
                "進捗 50%",
                StageRunLogProgressParser.extractDetail("[child] PM_AI_PROGRESS 50").orElseThrow());
    }

    @Test
    void extractDetail_skipsNoiseLines() {
        assertTrue(StageRunLogProgressParser.extractDetail("--- start: plan_simulation_stage2.py ---").isEmpty());
        assertTrue(StageRunLogProgressParser.extractDetail("[child] [stage1] キャッシュをクリアしました。").isEmpty());
    }
}
