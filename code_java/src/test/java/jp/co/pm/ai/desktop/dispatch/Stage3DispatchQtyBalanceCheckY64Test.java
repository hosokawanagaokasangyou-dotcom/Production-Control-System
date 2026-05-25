package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/** Y6-4 SEC: 計画1600m・実配台1000m のとき照合が NG になること。 */
class Stage3DispatchQtyBalanceCheckY64Test {

  @Test
  void ngWhenPlan1600ButActualTimeline1000() {
    String check =
            Stage3DispatchQtyBalanceCheck.formatCheck(
                    1600, 0, 1000, true, 200);
    assertTrue(check.startsWith("NG"), check);
    assertTrue(check.contains("1600"));
    assertTrue(check.contains("1000"));
  }

  @Test
  void okWhenActualMatchesRemaining() {
    assertEquals(
            "OK",
            Stage3DispatchQtyBalanceCheck.formatCheck(1600, 0, 1600, true, 200));
  }
}
