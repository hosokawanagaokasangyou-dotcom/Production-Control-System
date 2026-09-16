package jp.co.pm.ai.desktop.io;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

/**
 * 国分工場の工程別生産計画問合せ（先頭メタ＋ {@code MM/dd(曜)} 日付行＋ Excel シリアル受注日）。
 */
class TaskInputSourceRawGridIoAladdinPlanDisplayStepsTest {

    @Test
    void displaySteps_normalizesKokubuMonthDayHeadersUsingExcelSerialJuchuDate() {
        List<List<String>> data = new ArrayList<>();
        data.add(List.of("倉庫       : 511101 国分工場01本倉庫"));
        data.add(List.of("表示基準日 : 2026年09月01日"));
        data.add(List.of("完了区分   : 0:未完"));
        data.add(List.of("並び替え   : 0:倉庫-工程-注文"));
        data.add(List.of("", "", "", "", "08/02(日)", "08/02(日)", "08/02(日)"));
        data.add(List.of("機械名", "工程名", "依頼NO", "受注日", "加工数量", "加工速度", "加工時間"));
        data.add(List.of("スライス機1", "スライス", "R1", "46168", "120", "50", "2"));

        PlanInputTabularIo.TabularSheet raw =
                new PlanInputTabularIo.TabularSheet(List.of("列1"), data);
        PlanInputTabularIo.TabularSheet out =
                TaskInputSourceRawGridIo.applyAladdinProcessingPlanDisplaySteps(raw);

        Assertions.assertTrue(out.headers().contains("機械名"));
        Assertions.assertTrue(out.headers().contains("2026/08/02"));
        Assertions.assertFalse(out.headers().contains("加工速度"));
        Assertions.assertFalse(out.headers().contains("加工時間"));
        Assertions.assertEquals(1, out.rows().size());
        Assertions.assertEquals("スライス機1", out.rows().get(0).get(out.headers().indexOf("機械名")));
        Assertions.assertEquals("120", out.rows().get(0).get(out.headers().indexOf("2026/08/02")));
    }
}
