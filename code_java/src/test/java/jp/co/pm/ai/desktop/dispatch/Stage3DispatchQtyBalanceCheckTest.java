package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class Stage3DispatchQtyBalanceCheckTest {

  @Test
  void okWhenStage3MatchesConvertedMinusActual() {
    assertEquals("OK", Stage3DispatchQtyBalanceCheck.formatCheck(9000, 7500, 1500, true));
    assertEquals("OK", Stage3DispatchQtyBalanceCheck.formatCheck(3600, 2600, 1000, true));
    assertEquals("OK", Stage3DispatchQtyBalanceCheck.formatCheck(9000, 0, 9000, true));
  }

  @Test
  void ngWhenMismatch() {
    String ng = Stage3DispatchQtyBalanceCheck.formatCheck(9000, 7500, 1600, true);
    assertTrue(ng.startsWith("NG"));
    assertTrue(Stage3DispatchQtyBalanceCheck.isNgResult(ng));
  }

  @Test
  void emptyWhenNoStage3ColumnOrZeroDispatch() {
    assertEquals("", Stage3DispatchQtyBalanceCheck.formatCheck(9000, 7500, 1500, false));
    assertEquals("", Stage3DispatchQtyBalanceCheck.formatCheck(9000, 7500, 0, true));
  }

  @Test
  void isNgResultOnlyForNgPrefix() {
    assertFalse(Stage3DispatchQtyBalanceCheck.isNgResult("OK"));
    assertFalse(Stage3DispatchQtyBalanceCheck.isNgResult(""));
  }
}
