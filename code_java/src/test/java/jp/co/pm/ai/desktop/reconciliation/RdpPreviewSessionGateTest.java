package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/**
 * 連続実行で1本目プレビューの interrupt 通知が2本目を壊さないための世代ゲート。
 */
class RdpPreviewSessionGateTest {

    @Test
    void delayedStopFromPreviousSessionIsIgnoredAfterInvalidate() {
        RdpPreviewSessionGate gate = new RdpPreviewSessionGate();
        long first = gate.begin();
        gate.invalidate();
        long second = gate.begin();

        assertFalse(gate.isCurrent(first), "1本目終了後の遅延 notify は無効");
        assertTrue(gate.isCurrent(second), "2本目セッションは有効");
    }

    @Test
    void beginSupersedesPreviousSessionWithoutExplicitInvalidate() {
        RdpPreviewSessionGate gate = new RdpPreviewSessionGate();
        long first = gate.begin();
        long second = gate.begin();

        assertFalse(gate.isCurrent(first));
        assertTrue(gate.isCurrent(second));
    }

    @Test
    void invalidateAloneDropsInFlightSession() {
        RdpPreviewSessionGate gate = new RdpPreviewSessionGate();
        long first = gate.begin();
        gate.invalidate();

        assertFalse(gate.isCurrent(first));
    }
}
