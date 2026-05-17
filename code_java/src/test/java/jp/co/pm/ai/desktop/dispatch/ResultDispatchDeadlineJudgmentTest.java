package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.LinkedHashMap;
import java.util.Map;
import org.junit.jupiter.api.Test;

class ResultDispatchDeadlineJudgmentTest {

    @Test
    void usesAnswerDueBeforeSpecified() {
        Map<String, String> row = row(
                "依頼NO", "A001",
                "回答納期", "2026/05/20",
                "指定納期", "2026/05/25",
                "加工終了日時", "2026/05/19 15:00");
        assertEquals(ResultDispatchDeadlineJudgment.OK, ResultDispatchDeadlineJudgment.judgmentOkNg(row));
    }

    @Test
    void fallsBackToSpecifiedWhenAnswerEmpty() {
        Map<String, String> row = row(
                "依頼NO", "A001",
                "回答納期", "",
                "指定納期", "2026/05/20",
                "加工終了日時", "2026/05/20 10:00");
        assertEquals(ResultDispatchDeadlineJudgment.NG, ResultDispatchDeadlineJudgment.judgmentOkNg(row));
    }

    @Test
    void nonVRequiresCompletionBeforeDueDayStart() {
        Map<String, String> row = row(
                "依頼NO", "A001",
                "指定納期", "2026/05/20",
                "加工終了日時", "2026/05/19 23:59");
        assertEquals(ResultDispatchDeadlineJudgment.OK, ResultDispatchDeadlineJudgment.judgmentOkNg(row));

        row.put("加工終了日時", "2026/05/20 00:00");
        assertEquals(ResultDispatchDeadlineJudgment.NG, ResultDispatchDeadlineJudgment.judgmentOkNg(row));
    }

    @Test
    void vPrefixAllowsUntilSixteenOnDueDay() {
        Map<String, String> row = row(
                "依頼NO", "V123",
                "指定納期", "2026/05/20",
                "加工終了日時", "2026/05/20 16:00");
        assertEquals(ResultDispatchDeadlineJudgment.OK, ResultDispatchDeadlineJudgment.judgmentOkNg(row));

        row.put("加工終了日時", "2026/05/20 16:01");
        assertEquals(ResultDispatchDeadlineJudgment.NG, ResultDispatchDeadlineJudgment.judgmentOkNg(row));
    }

    @Test
    void emptyWhenNoPlanEndOrDue() {
        Map<String, String> row = row("依頼NO", "A001", "指定納期", "2026/05/20");
        assertEquals("", ResultDispatchDeadlineJudgment.judgmentOkNg(row));

        row.put("加工終了日時", "2026/05/19 12:00");
        row.put("指定納期", "");
        row.put("回答納期", "");
        assertEquals("", ResultDispatchDeadlineJudgment.judgmentOkNg(row));
    }

    private static Map<String, String> row(String... kv) {
        Map<String, String> m = new LinkedHashMap<>();
        for (int i = 0; i + 1 < kv.length; i += 2) {
            m.put(kv[i], kv[i + 1]);
        }
        return m;
    }
}
