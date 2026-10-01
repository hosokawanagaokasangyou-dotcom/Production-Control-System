package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.reconciliation.RequestFormOriginalFee;

class NagaokaExcelFixAdviceTest {

    @Test
    @DisplayName("②が原本と違うときは長岡側の元シート単価を直す")
    void tellsKokubuToFixSourcePrices() {
        List<NagaokaExcelFixAdvice.Line> lines = NagaokaExcelFixAdvice.lines(result(200.0, 100.0), plan(200.0));
        assertEquals(1, lines.size());
        String action = lines.get(0).action();
        assertTrue(action.contains("原本加工賃"), action);
        assertTrue(action.contains("東レまとめのAAは直接直さない"), action);
        assertTrue(action.contains("H〜X"), action);
        assertTrue(action.contains("直すのは長岡側だけ"), action);
    }

    @Test
    @DisplayName("②が原本と一致し①だけ違うときはExcelを直さない")
    void leavesExcelWhenNagaokaMatchesOriginal() {
        String action = NagaokaExcelFixAdvice.lines(result(100.0, 200.0), plan(200.0)).get(0).action();
        assertTrue(action.contains("Excelの金額は直さない"), action);
        assertTrue(action.contains("東レへ報告"), action);
    }

    @Test
    @DisplayName("湖南は加工賃試算の単価を直す")
    void konanEditsRatesNotTotal() {
        VerifyResult konan = new VerifyResult(
                FactoryProfile.of(FactoryId.KONAN),
                new YearMonthKey(2026, 9),
                List.of(new RecordA("192265M", "V9-9", 200.0, 100.0, 100.0, 100.0, Judge.MISMATCH, "")),
                List.of(),
                info(),
                List.of(),
                null,
                null,
                null);
        String action = NagaokaExcelFixAdvice.lines(konan, plan(200.0)).get(0).action();
        assertTrue(action.contains("種類別の単価を直す"), action);
        assertTrue(NagaokaExcelFixAdvice.principle(FactoryProfile.of(FactoryId.KOKUBU)).contains("加工1・加工2の並びは工程順ではない"));
    }

    private static VerifyResult result(double amount1, double amount2) {
        return new VerifyResult(
                FactoryProfile.of(FactoryId.KOKUBU),
                new YearMonthKey(2026, 9),
                List.of(new RecordA(
                        "192265M", "V9-9", amount1, amount2, amount1 - amount2, amount1 - amount2, Judge.MISMATCH, "")),
                List.of(),
                info(),
                List.of(),
                null,
                null,
                null);
    }

    private static Map<String, Object> info() {
        Map<String, Object> info = new HashMap<>();
        info.put("許容差", 0.5);
        return info;
    }

    private static RequestFormOriginalAttacher.Plan plan(double yen) {
        return new RequestFormOriginalAttacher.Plan(
                List.of(new RequestFormOriginalAttacher.Item(
                        "V9-9",
                        "V9-9",
                        "検証A",
                        null,
                        "",
                        "",
                        "",
                        new RequestFormOriginalFee.Result(100, 2, yen, "配台順でスライス後"))),
                List.of());
    }
}
