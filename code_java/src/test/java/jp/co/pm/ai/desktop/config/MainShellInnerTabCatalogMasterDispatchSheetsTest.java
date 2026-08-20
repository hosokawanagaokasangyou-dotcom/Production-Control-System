package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.MainShellTabId;

class MainShellInnerTabCatalogMasterDispatchSheetsTest {

    @Test
    void labelsFor_masterDispatchSheets_areFourSheetTabs() {
        assertEquals(
                List.of("資格（skills）", "必要人数（need）", "加工速度（speed）", "組み合わせ表"),
                MainShellInnerTabCatalog.labelsFor(MainShellTabId.MASTER_DISPATCH_SHEETS));
    }
}
