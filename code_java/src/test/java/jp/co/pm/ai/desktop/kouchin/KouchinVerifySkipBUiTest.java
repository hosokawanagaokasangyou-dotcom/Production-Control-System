package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.kouchin.verify.BothResult;
import jp.co.pm.ai.kouchin.verify.FactoryId;
import jp.co.pm.ai.kouchin.verify.FactoryProfile;
import jp.co.pm.ai.kouchin.verify.VerifyResult;
import jp.co.pm.ai.kouchin.verify.YearMonthKey;

class KouchinVerifySkipBUiTest {

    @Test
    @DisplayName("検証Bスキップ時のバナーは工場名と警告を含む")
    void skipBBannerIncludesFactoryAndWarning() {
        VerifyResult kokubu = skipped("【検証Bスキップ】③月次実績に対象月（2026年7月度）のデータがありません。");
        String text = KouchinVerifyTabController.skipBBannerText(new BothResult(
                kokubu, null, null, null, ""));
        assertTrue(text.contains("国分"));
        assertTrue(text.contains("【検証Bスキップ】"));
        assertTrue(text.contains("2026年7月度"));
    }

    @Test
    @DisplayName("両工場スキップならバナーに両方出す")
    void skipBBannerJoinsBothFactories() {
        VerifyResult kokubu = skipped("【検証Bスキップ】国分は対象月なし");
        VerifyResult konan = skippedKonan("【検証Bスキップ】湖南は対象月なし");
        String text = KouchinVerifyTabController.skipBBannerText(new BothResult(
                kokubu, konan, null, null, ""));
        assertTrue(text.contains("国分"));
        assertTrue(text.contains("湖南"));
        assertEquals(2, text.split("【検証Bスキップ】", -1).length - 1);
    }

    @Test
    @DisplayName("スキップしていなければバナーは空")
    void skipBBannerEmptyWhenNotSkipped() {
        Map<String, Object> info = new HashMap<>();
        VerifyResult r = new VerifyResult(
                FactoryProfile.of(FactoryId.KOKUBU),
                new YearMonthKey(2026, 7),
                List.of(),
                List.of(),
                info,
                List.of("他の警告"),
                null,
                null,
                null);
        assertEquals("", KouchinVerifyTabController.skipBBannerText(new BothResult(r, null, null, null, "")));
        assertEquals("", KouchinVerifyTabController.skipBBannerText(null));
    }

    private static VerifyResult skipped(String warning) {
        Map<String, Object> info = new HashMap<>();
        info.put("検証Bスキップ", true);
        info.put("検証Bスキップ理由", warning);
        return new VerifyResult(
                FactoryProfile.of(FactoryId.KOKUBU),
                new YearMonthKey(2026, 7),
                List.of(),
                List.of(),
                info,
                List.of(warning),
                null,
                null,
                null);
    }

    private static VerifyResult skippedKonan(String warning) {
        Map<String, Object> info = new HashMap<>();
        info.put("検証Bスキップ", true);
        info.put("検証Bスキップ理由", warning);
        return new VerifyResult(
                FactoryProfile.of(FactoryId.KONAN),
                new YearMonthKey(2026, 7),
                List.of(),
                List.of(),
                info,
                List.of(warning),
                null,
                null,
                null);
    }
}
