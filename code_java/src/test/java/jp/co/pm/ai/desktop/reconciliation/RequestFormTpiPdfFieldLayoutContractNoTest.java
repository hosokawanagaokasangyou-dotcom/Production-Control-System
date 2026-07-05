package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class RequestFormTpiPdfFieldLayoutContractNoTest {

    @Test
    void parseContractNo_prefersNyukoPOverTableX() {
        String text =
                """
                出荷指図・契約No X000081337
                投入原反 X000081196
                ・入庫お願いします。『P000075558』
                """;
        assertEquals("P000075558", RequestFormTpiPdfFieldLayout.parseContractNo(text));
    }

    @Test
    void parseContractNo_ignoresShippingInstructionX() {
        String text = "出荷指図・契約No X000081337";
        assertEquals("", RequestFormTpiPdfFieldLayout.parseContractNo(text));
        assertEquals("X000081337", RequestFormTpiPdfFieldLayout.parseXContractNo(text));
    }

    @Test
    void parseContractNo_usesPInHachuContractNoColumn() {
        String text =
                """
                依頼NO. PN06-01 希望納期 2026年 6月 3日 湖南
                発注・契約No P000075287
                加工製品 ① 7C8 FEL3002BY05WDLG-EC 2,000
                """;
        assertEquals("P000075287", RequestFormTpiPdfFieldLayout.parseContractNo(text));
    }

    @Test
    void parseContractNo_usesPInBodyWhenNoNyukoNote() {
        String text = "投入原反　① 7A1 FEL3002BY05WDLG 2,300 4/28 X000079828\n2,300 P000074932";
        assertEquals("P000074932", RequestFormTpiPdfFieldLayout.parseContractNo(text));
    }
}
