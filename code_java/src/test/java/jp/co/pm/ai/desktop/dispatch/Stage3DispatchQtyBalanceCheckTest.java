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
  void rollAlignedLabelWhenRemainingBelowRollUnit() {
    assertEquals(
            "20 (91m)",
            Stage3DispatchQtyBalanceCheck.formatCheck(20, 0, 91, true, 91));
    assertFalse(Stage3DispatchQtyBalanceCheck.isNgResult("20 (91m)"));
  }

  @Test
  void okWhenRollUnitZeroUsesRawRemaining() {
    assertEquals("OK", Stage3DispatchQtyBalanceCheck.formatCheck(20, 0, 20, true, 0));
  }

  @Test
  void ngWhenMismatchEvenWithRollUnit() {
    String ng = Stage3DispatchQtyBalanceCheck.formatCheck(20, 0, 180, true, 91);
    assertTrue(ng.startsWith("NG"));
    assertTrue(ng.contains("91"));
    assertTrue(ng.contains("180"));
    assertTrue(Stage3DispatchQtyBalanceCheck.isNgResult(ng));
  }

  @Test
  void ngWhenMismatchWithoutRollUnit() {
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
    assertFalse(Stage3DispatchQtyBalanceCheck.isNgResult("20 (91m)"));
    assertFalse(Stage3DispatchQtyBalanceCheck.isNgResult(""));
  }

  @Test
  void rollAlignedDispatchMMatchesStage2CeilFormula() {
    assertEquals(91.0, Stage3DispatchQtyBalanceCheck.rollAlignedDispatchM(20, 91), 1e-9);
    assertEquals(182.0, Stage3DispatchQtyBalanceCheck.rollAlignedDispatchM(100, 91), 1e-9);
    assertEquals(20.0, Stage3DispatchQtyBalanceCheck.rollAlignedDispatchM(20, 0), 1e-9);
  }
}
