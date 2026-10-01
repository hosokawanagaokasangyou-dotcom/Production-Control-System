package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class RequestFormOriginalFeeRulesTest {

    @Test
    @DisplayName("スライスとスリットの加工後数量を判定指示に含める")
    void quantityRulesCoverSliceAndSlit() {
        String rules = RequestFormOriginalFee.quantityRules();
        assertTrue(rules.contains("スライス"), rules);
        assertTrue(rules.contains("厚みを1/2"), rules);
        assertTrue(rules.contains("2倍"), rules);
        assertTrue(rules.contains("端部をトリミング"), rules);
        assertTrue(rules.contains("変わらない"), rules);
        assertTrue(rules.contains("二つにスリット"), rules);
        assertTrue(rules.contains("三つにスリット"), rules);
        assertTrue(rules.contains("整数倍"), rules);
        assertTrue(rules.contains("依頼書原本の加工1・加工2の並びではない"), rules);
        assertTrue(rules.contains("配台システム"), rules);
        assertTrue(rules.contains("加工内容"), rules);
        assertTrue(rules.contains("原本の並びで代用せず"), rules);
    }
}
