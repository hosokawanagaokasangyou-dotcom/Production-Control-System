package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.FactorySite;

class FactorySiteComboPresentationTest {

    @Test
    void selectedDisplayLabel_isFactoryNameOnlyWithoutAccessCheck() {
        assertEquals("国分工場", FactorySiteComboPresentation.selectedDisplayLabel(FactorySite.KOKUBU));
        assertEquals("湖南工場", FactorySiteComboPresentation.selectedDisplayLabel(FactorySite.KONAN));
        assertEquals("", FactorySiteComboPresentation.selectedDisplayLabel(null));
        assertFalse(
                FactorySiteComboPresentation.selectedDisplayLabel(FactorySite.KOKUBU)
                        .contains("ユーザー登録無し"));
    }
}
