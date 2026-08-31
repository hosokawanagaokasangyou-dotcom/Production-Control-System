package jp.co.pm.ai.desktop.reconciliation;

import java.util.concurrent.atomic.AtomicLong;

/**
 * RDP 右ペインプレビューのセッション世代。
 *
 * <p>1本目終了時の worker interrupt が {@code Platform.runLater} 経由で遅延通知されるため、
 * 連続実行の2本目プレビューを後から {@code removePreviewPane} してしまう競合を防ぐ。
 */
final class RdpPreviewSessionGate {

    private final AtomicLong generation = new AtomicLong();

    /** 新しいプレビューセッションを開始し、そのトークンを返す。 */
    long begin() {
        return generation.incrementAndGet();
    }

    /** 進行中セッションを無効化する（意図的な停止・ペイン削除時）。 */
    void invalidate() {
        generation.incrementAndGet();
    }

    /** {@code token} が現在のセッションと一致するか。 */
    boolean isCurrent(long token) {
        return token == generation.get();
    }
}
