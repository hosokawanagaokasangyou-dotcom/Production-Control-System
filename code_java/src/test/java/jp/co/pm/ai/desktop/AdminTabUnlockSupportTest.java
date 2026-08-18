package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.HashSet;
import java.util.Set;

import org.junit.jupiter.api.Test;

class AdminTabUnlockSupportTest {

    @Test
    void generateLockNo_isFourDigitZeroPaddedNumber() {
        for (int i = 0; i < 40; i++) {
            String lockNo = AdminTabUnlockSupport.generateLockNo();
            assertEquals(AdminTabUnlockSupport.LOCK_NO_DIGITS, lockNo.length());
            assertTrue(lockNo.chars().allMatch(Character::isDigit), lockNo);
            assertTrue(AdminTabUnlockSupport.verifyLockNo(lockNo, lockNo));
        }
    }

    @Test
    void generateLockNo_variesAcrossCalls() {
        Set<String> distinct = new HashSet<>();
        for (int i = 0; i < 30; i++) {
            distinct.add(AdminTabUnlockSupport.generateLockNo());
        }
        assertTrue(distinct.size() > 1, "連続生成が同一値のみ");
    }

    @Test
    void verifyLockNo_acceptsExactMatchAndStrippedInput() {
        assertTrue(AdminTabUnlockSupport.verifyLockNo("0482", "0482"));
        assertTrue(AdminTabUnlockSupport.verifyLockNo("0482", " 0482 "));
        assertTrue(AdminTabUnlockSupport.verifyLockNo("0482", "\t0482\n"));
    }

    @Test
    void verifyLockNo_rejectsMismatchEmptyAndNull() {
        assertFalse(AdminTabUnlockSupport.verifyLockNo("0482", "482"));
        assertFalse(AdminTabUnlockSupport.verifyLockNo("0482", "0483"));
        assertFalse(AdminTabUnlockSupport.verifyLockNo("0482", ""));
        assertFalse(AdminTabUnlockSupport.verifyLockNo("0482", null));
        assertFalse(AdminTabUnlockSupport.verifyLockNo(null, "0482"));
        assertFalse(AdminTabUnlockSupport.verifyLockNo("", ""));
    }
}
