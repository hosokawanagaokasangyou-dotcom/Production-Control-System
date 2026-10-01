package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
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

    @Test
    @DisplayName("同じ依頼NOの契約は合算せず、契約ごとの原本加工賃で直す")
    void splitsSameIraiByKeiyaku() {
        VerifyResult result = new VerifyResult(
                FactoryProfile.of(FactoryId.KOKUBU),
                new YearMonthKey(2026, 9),
                List.of(
                        new RecordA("192591B", "T9-7", 7000.0, 8800.0, -1800.0, -1800.0, Judge.MISMATCH, ""),
                        new RecordA("192592", "T9-7", 1500.0, 1500.0, 0.0, null, Judge.MATCH, "")),
                List.of(),
                info(),
                List.of(),
                null,
                null,
                null);
        RequestFormOriginalAttacher.Plan plan = new RequestFormOriginalAttacher.Plan(
                List.of(new RequestFormOriginalAttacher.Item(
                        "T9-7",
                        "T9-7",
                        "検証A",
                        null,
                        "",
                        "",
                        "",
                        new RequestFormOriginalFee.Result(
                                300,
                                0,
                                10300,
                                "合計",
                                List.of(
                                        new RequestFormOriginalFee.ContractFee("192591B", 200, 44, 8800, "44×200"),
                                        new RequestFormOriginalFee.ContractFee("192592", 100, 15, 1500, "15×100"))))),
                List.of());
        List<NagaokaExcelFixAdvice.Line> lines = NagaokaExcelFixAdvice.lines(result, plan);
        assertEquals(1, lines.size(), lines.toString());
        assertEquals("192591B", lines.get(0).keiyaku());
        assertEquals(7000.0, lines.get(0).amount1(), 0.001);
        assertEquals(8800.0, lines.get(0).amount2(), 0.001);
        assertEquals(8800.0, lines.get(0).originalYen(), 0.001);
        assertTrue(lines.get(0).action().contains("Excelの金額は直さない"), lines.get(0).action());
        assertTrue(lines.get(0).action().contains("-1,800"), lines.get(0).action());
        assertTrue(!lines.get(0).action().contains("10,300"), lines.get(0).action());
    }

    @Test
    @DisplayName("内訳が無い複数契約には依頼全体の原本加工賃を当てない")
    void doesNotCopySheetTotalOntoEachContract() {
        VerifyResult result = new VerifyResult(
                FactoryProfile.of(FactoryId.KOKUBU),
                new YearMonthKey(2026, 9),
                List.of(
                        new RecordA("192265M", "Y9-9", 400000.0, 400000.0, 0.0, null, Judge.MATCH, ""),
                        new RecordA("192266", "Y9-9", 383511.0, 392120.0, -8609.0, -8609.0, Judge.MISMATCH, "")),
                List.of(),
                info(),
                List.of(),
                null,
                null,
                null);
        RequestFormOriginalAttacher.Plan sheetTotal = new RequestFormOriginalAttacher.Plan(
                List.of(new RequestFormOriginalAttacher.Item(
                        "Y9-9", "Y9-9", "検証A", null, "", "", "",
                        new RequestFormOriginalFee.Result(0, 0, 1037300.0, "依頼全体"))),
                List.of());
        List<NagaokaExcelFixAdvice.Line> lines = NagaokaExcelFixAdvice.lines(result, sheetTotal);
        assertEquals(1, lines.size(), lines.toString());
        assertEquals("192266", lines.get(0).keiyaku());
        assertEquals(383511.0, lines.get(0).amount1(), 0.001);
        assertNull(lines.get(0).originalYen());
        assertTrue(lines.get(0).action().contains("合算した金額は使わない"), lines.get(0).action());
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
