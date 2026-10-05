package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class UnifiedMailBuilderTest {

    @Test
    @DisplayName("メール本文は難波の挨拶と拠点表にする")
    void usesNanbaGreetingAndTable() {
        MailSnapshot kokubu = new MailSnapshot(
                "kokubu", "国分工場", "2026年9月度", "2026年10月度",
                13248142, 13251848, -3690, -16,
                4, 4, -16, 0, 0, 0, 0, 0, 0, "", "");
        MailSnapshot konan = new MailSnapshot(
                "konan", "湖南工場", "2026年9月度", "2026年10月度",
                5447638, 5465238, 0, -17600,
                1, 1, -17600, 0, 0, 0, 0, 0, 0, "", "");
        String text = UnifiedMailBuilder.buildText(kokubu, konan);
        assertTrue(text.contains("長岡産業の難波です。"), text);
        assertTrue(text.contains("国分（A010）\t13,248,142\t13,251,848\t-3,690\t-16\t13,248,142\t0"), text);
        assertTrue(text.contains("湖南（A010P）\t5,447,638\t5,465,238\t0\t-17,600\t5,447,638\t0"), text);
        assertTrue(text.contains("合計　5件の差異（-17,616円）不足がありました。"), text);
        assertTrue(text.contains("次月（2026年10月度）"), text);
    }
}
