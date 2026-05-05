package jp.co.pm.ai.desktop.io.gantt;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

class PersonNameBadgeTextTest {

    @Test
    void surname_twoChars_yamada_taro() {
        assertEquals("山田", PersonNameBadgeText.badgeTwoFromRawName("山田 太郎"));
    }

    @Test
    void singleToken_fullUsedAsSei() {
        assertEquals("佐藤", PersonNameBadgeText.badgeTwoFromRawName("佐藤"));
    }

    @Test
    void short_surname_oneChar() {
        assertEquals("林", PersonNameBadgeText.badgeTwoFromRawName("林"));
    }

    @Test
    void normalize_tomita_variants() {
        assertEquals("冨田", PersonNameBadgeText.badgeTwoFromRawName("富田 花子"));
    }

    @Test
    void badgeList_op_then_sub_taskSplit() {
        List<String> b =
                PersonNameBadgeText.badgeListFromOpSub(
                        "山田 太郎", "佐藤,鈴木 花", false);
        assertEquals(List.of("山田", "佐藤", "鈴木"), b);
    }

    @Test
    void badgeList_startupCommaSplit() {
        List<String> b =
                PersonNameBadgeText.badgeListFromOpSub(
                        "田中 一郎", "佐藤、鈴木", true);
        assertEquals(List.of("田中", "佐藤", "鈴木"), b);
    }

    @Test
    void firstN_codePoints_surrogate_safe() {
        assertEquals("𠮷野", PersonNameBadgeText.firstNCodePoints("𠮷野口", 2));
    }
}
