package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.MainShellTabId;

class MainShellInnerTabCatalogRequestFormInputTest {

    @Test
    void labelsFor_requestFormInput_includesJuchuSearchAfterIndexSheet() {
        assertEquals(
                List.of(
                        "一括照合データベース・受注管理",
                        "目次シート",
                        "受注検索",
                        "【設定】",
                        "後加工商品マスタ",
                        "マスター一覧"),
                MainShellInnerTabCatalog.labelsFor(MainShellTabId.REQUEST_FORM_INPUT));
    }

    @Test
    void nestedInnerTabs_masterList_isIndex5() {
        assertEquals(
                List.of("機械コード", "工程マスタ", "加工内容マスタ"),
                MainShellInnerTabCatalog.nestedInnerTabLabelsUnderInnerTab(
                        MainShellTabId.REQUEST_FORM_INPUT, 5));
        assertEquals(
                List.of(),
                MainShellInnerTabCatalog.nestedInnerTabLabelsUnderInnerTab(
                        MainShellTabId.REQUEST_FORM_INPUT, 4));
    }
}
