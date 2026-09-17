package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class KouchinVerifyOperatorActionTest {

    @Test
    void operatorResultForDrop_mapsEmptyCopiedAndNone() {
        assertEquals("empty", KouchinVerifyTabController.operatorResultForDrop(true, false));
        assertEquals("ok", KouchinVerifyTabController.operatorResultForDrop(false, true));
        assertEquals("none", KouchinVerifyTabController.operatorResultForDrop(false, false));
    }

    @Test
    void operatorResultForVerify_mapsErrorWarnOk() {
        assertEquals(
                "error",
                KouchinVerifyTabController.operatorResultForVerify(true, "失敗: 中断"));
        assertEquals(
                "warn",
                KouchinVerifyTabController.operatorResultForVerify(
                        false, "検証完了。 Excel成功1 失敗 読取不可"));
        assertEquals(
                "ok",
                KouchinVerifyTabController.operatorResultForVerify(false, "検証完了。"));
    }
}
