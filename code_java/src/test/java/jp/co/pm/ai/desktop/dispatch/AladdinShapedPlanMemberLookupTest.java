package jp.co.pm.ai.desktop.dispatch;

import java.util.List;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

class AladdinShapedPlanMemberLookupTest {

    @Test
    void lookup_staticColumnWhenProcessMatches() {
        List<String> headers =
                List.of("機械名", "依頼NO", "工程名", "担当OP_指定", "2026/05/23");
        List<List<String>> rows =
                List.of(List.of("M1", "R1", "スリット", "田中一郎", "50"));
        Assertions.assertEquals(
                "田中一郎",
                AladdinShapedPlanMemberLookup.lookup(
                        headers, rows, "M1", "R1", "スリット", "2026/05/23"));
    }

    @Test
    void lookup_fallsBackToMachineAndRequestWhenProcessDiffers() {
        List<String> headers =
                List.of("機械名", "依頼NO", "工程名", "担当OP指定", "2026/05/23");
        List<List<String>> rows =
                List.of(List.of("M1", "R1", "カット", "鈴木次郎", "50"));
        Assertions.assertEquals(
                "鈴木次郎",
                AladdinShapedPlanMemberLookup.lookup(
                        headers, rows, "M1", "R1", "スリット", "2026/05/23"));
    }

    @Test
    void lookup_dateColumnTextWhenStaticEmpty() {
        List<String> headers = List.of("機械名", "依頼NO", "工程名", "2026/05/23");
        List<List<String>> rows = List.of(List.of("M1", "R1", "P1", "山田太郎"));
        Assertions.assertEquals(
                "山田太郎",
                AladdinShapedPlanMemberLookup.lookup(
                        headers, rows, "M1", "R1", "P1", "2026/05/23"));
    }

    @Test
    void lookup_rejectsMachineCodeInDateColumn() {
        List<String> headers = List.of("機械名", "依頼NO", "工程名", "2026/05/25");
        List<List<String>> rows = List.of(List.of("M1", "R1", "P1", "[Y5-135]"));
        Assertions.assertEquals(
                "",
                AladdinShapedPlanMemberLookup.lookup(
                        headers, rows, "M1", "R1", "P1", "2026/05/25"));
    }
}
