package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class KonanDailyReportLookupTest {

    private static final String HEADER =
            "倉庫コード,倉庫,機械コード,機械名,工程コード,工程名,加工日付,依頼NO,受注NO,加工実績NO,"
                    + "加工担当者1コード,加工担当者1名,加工担当者2コード,加工担当者2名,加工担当者3コード,加工担当者3名,"
                    + "開始時間,終了時間,残業時間_分,残業当者1コード,残業当者1名,残業当者2コード,残業当者2名,"
                    + "残業当者3コード,残業当者3名,停機時間_分,休憩時間_分,稼働時間_分,所要時間,製品,製品梱包,製品色,製品区分,"
                    + "投入原反,原反梱包,原反色,原反区分,受注数量,換算数量,加工日加工予定数,実加工量,実製品出来高,原反着予定,"
                    + "指定納期,加工開始予定日,加工完了予定日,加工内容,商品特記事項,製造条件特記事項,納入先1,納入先2,コア着日,"
                    + "回答納期,出荷予定日,完了区分,加工完了日,注文単位加工完了区分,注文単位加工完了日,加工実績累計,実製品出来高累計";

    @Test
    void formatCompletionDisplay_mapsDailyReportValues() {
        assertEquals("未了", KonanDailyReportLookup.formatCompletionDisplay("0:未完"));
        assertEquals("完了", KonanDailyReportLookup.formatCompletionDisplay("1:完了"));
        assertEquals("", KonanDailyReportLookup.formatCompletionDisplay(""));
    }

    @Test
    void loadFromPath_picksLatestDayPerKey(@TempDir Path temp) throws Exception {
        Path csv = temp.resolve("sample.csv");
        String body =
                "meta1\n"
                        + "meta2\n"
                        + "meta3\n"
                        + HEADER
                        + "\n"
                        + row("スライス機1　湖南", "スライス", "2026/06/22", "Y6-19", "0:未完")
                        + "\n"
                        + row("スライス機1　湖南", "スライス", "2026/06/23", "Y6-19", "0:未完")
                        + "\n"
                        + row("スライス機1　湖南", "スライス", "2026/06/23", "Y6-18", "1:完了");
        Files.writeString(csv, body, StandardCharsets.UTF_8);

        KonanDailyReportLookup lookup = KonanDailyReportLookup.loadFromPath(csv);
        assertTrue(lookup.isLoaded());
        assertEquals(csv.toAbsolutePath().normalize().toString(), lookup.sourcePath());
        assertEquals("未了", lookup.completionDisplay("Y6-19", "スライス", "スライス機1　湖南"));
        assertEquals("完了", lookup.completionDisplay("Y6-18", "スライス", "スライス機1　湖南"));
        assertEquals("", lookup.completionDisplay("UNKNOWN", "スライス", "スライス機1　湖南"));
    }

    @Test
    void readTableFromPath_returnsAllDataRows(@TempDir Path temp) throws Exception {
        Path csv = temp.resolve("table.csv");
        String body =
                "meta1\n"
                        + "meta2\n"
                        + "meta3\n"
                        + HEADER
                        + "\n"
                        + row("スライス機1　湖南", "スライス", "2026/06/22", "Y6-19", "0:未完")
                        + "\n"
                        + row("スライス機1　湖南", "スライス", "2026/06/23", "Y6-18", "1:完了");
        Files.writeString(csv, body, StandardCharsets.UTF_8);

        KonanDailyReportLookup.DailyReportCsvTable table =
                KonanDailyReportLookup.readTableFromPath(csv);
        assertEquals(csv.toAbsolutePath().normalize().toString(), table.sourcePath());
        assertEquals(List.of("meta1", "meta2", "meta3"), table.metaLines());
        assertEquals(HEADER.split(",", -1).length, table.headers().size());
        assertEquals(2, table.rows().size());
        assertEquals("Y6-19", table.rows().getFirst().get("依頼NO"));
        assertEquals("0:未完", table.rows().getFirst().get("完了区分"));
        assertEquals("1:完了", table.rows().get(1).get("完了区分"));
    }

    @Test
    void orderCompletionStatus_aggregatesPerIrai(@TempDir Path temp) throws Exception {
        Path csv = temp.resolve("order.csv");
        String body =
                "meta1\nmeta2\nmeta3\n"
                        + HEADER
                        + "\n"
                        + row("SEC機湖南", "SEC", "2026/07/05", "C6-11", "1:完了")
                        + "\n"
                        + row("スライス機1　湖南", "スライス", "2026/06/23", "Y6-19", "0:未完");
        Files.writeString(csv, body, StandardCharsets.UTF_8);
        KonanDailyReportLookup lookup = KonanDailyReportLookup.loadFromPath(csv);
        assertEquals("完了", lookup.orderCompletionStatus("C6-11"));
        assertEquals("未了", lookup.orderCompletionStatus("Y6-19"));
        assertEquals("―", lookup.orderCompletionStatus("UNKNOWN"));
        assertEquals(1, lookup.entriesForOrder("C6-11").size());
    }

    private static String row(
            String machine, String process, String day, String iraiNo, String completion) {
        int cols = HEADER.split(",", -1).length;
        String[] cells = new String[cols];
        put(cells, "機械名", machine);
        put(cells, "工程名", process);
        put(cells, "加工日付", day);
        put(cells, "依頼NO", iraiNo);
        put(cells, "完了区分", completion);
        return String.join(",", cells);
    }

    private static void put(String[] cells, String name, String value) {
        String[] headers = HEADER.split(",", -1);
        for (int i = 0; i < headers.length; i++) {
            if (headers[i].equals(name)) {
                cells[i] = value;
                return;
            }
        }
        throw new IllegalArgumentException(name);
    }
}
