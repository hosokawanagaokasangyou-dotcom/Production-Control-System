package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertNotEquals;

import org.junit.jupiter.api.Test;

class MemberAttendanceGridCacheKeyTest {

    @Test
    void cacheKey_differsByFactoryForSameYearMonth() {
        assertNotEquals(
                MemberAttendanceTabController.memberGridCacheKey("KONAN", 2026, 8),
                MemberAttendanceTabController.memberGridCacheKey("KOKUBU", 2026, 8),
                "工場切替後に同一年月の旧工場キャッシュを使ってはならない");
    }
}
