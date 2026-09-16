package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class VerifyEngineResidualTest {

    @Test
    @DisplayName("報告残差は手動①正を内訳に含める")
    void residualIncludesManual1() {
        double residual = VerifyEngine.reportResidual(
                10_000, 8_000, 0, 0, 1_000, 1_000, 0, 0);
        assertEquals(0.0, residual, 0.001);
    }
}
