package jp.co.pm.ai.desktop.dispatch;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

import javafx.animation.Animation;
import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.util.Duration;

/**
 * 配台リプレイ（プラン「配台結果リプレイ（B）」）。
 *
 * <p>{@code 結果_配台表.json} 由来の行（{@code ResultDispatchDocument#rows()}）を加工開始日時順に並べ、
 * JavaFX {@link Timeline} で 1 ステップずつ {@link StepVisitor} へ通知する。セルへのハイライト・
 * スクロールは呼び出し側（ワイド表を持つコントローラ）が {@link StepVisitor} 内で行う。
 *
 * <p>Python 変更不要・グリッドモデル非破壊（選択/フォーカスのみ）。FX アプリケーションスレッドで使う。
 */
public final class DispatchReplayController {

    /** 1 ステップ ＝ ある (依頼NO, 機械) を、ある配台日に割り付けた事象。 */
    public record Step(
            String requestNo,
            String machine,
            LocalDate dispatchDate,
            LocalDateTime sortKey,
            int qty,
            String label) {}

    /** 各ステップ再生時に呼ばれる。{@code index} は 0 始まり、{@code total} は総ステップ数。 */
    @FunctionalInterface
    public interface StepVisitor {
        void visit(Step step, int index, int total);
    }

    private static final List<DateTimeFormatter> START_FORMATS =
            List.of(
                    DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm"),
                    DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"),
                    DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm"),
                    DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"),
                    DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm"),
                    DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss"));

    private static final List<DateTimeFormatter> DATE_FORMATS =
            List.of(
                    DateTimeFormatter.ofPattern("yyyy-MM-dd"),
                    DateTimeFormatter.ofPattern("yyyy/MM/dd"),
                    DateTimeFormatter.ofPattern("yyyy/M/d"),
                    DateTimeFormatter.ofPattern("yyyy-M-d"));

    private Timeline timeline;
    private List<Step> steps = List.of();
    private double intervalMillis = 300.0;

    /**
     * 表示行（{@code ResultDispatchDocument#rows()}）から再生キューを作る。
     * 当日配台数量 &gt; 0 かつ配台日が解釈可能な行のみ対象。加工開始日時→配台日の順で安定ソート。
     */
    public static List<Step> buildStepsFromRows(List<Map<String, String>> rows) {
        List<Step> out = new ArrayList<>();
        if (rows == null) {
            return out;
        }
        for (Map<String, String> r : rows) {
            if (r == null) {
                continue;
            }
            int qty = parseQty(r.get(ResultDispatchSchema.COL_DISPATCH_QTY));
            if (qty <= 0) {
                continue;
            }
            LocalDate date = parseDate(r.get(ResultDispatchSchema.COL_DISPATCH_DATE));
            if (date == null) {
                continue;
            }
            String requestNo = trimmed(r.get("依頼NO")); // Literal Japanese key
            String machine = trimmed(r.get(ResultDispatchSchema.COL_MACHINE));
            LocalDateTime sortKey = parseStart(r.get("加工開始日時")); // Literal Japanese key
            if (sortKey == null) {
                sortKey = date.atStartOfDay();
            }
            String label =
                    requestNo
                            + (machine.isEmpty() ? "" : " / " + machine)
                            + " → "
                            + date
                            + "（"
                            + qty
                            + "）";
            out.add(new Step(requestNo, machine, date, sortKey, qty, label));
        }
        out.sort(
                Comparator.comparing(Step::sortKey)
                        .thenComparing(Step::dispatchDate)
                        .thenComparing(Step::requestNo, Comparator.nullsLast(Comparator.naturalOrder())));
        return out;
    }

    /** 再生キューを差し替える（再生中なら停止）。 */
    public void load(List<Step> next) {
        stop();
        this.steps = (next == null) ? List.of() : List.copyOf(next);
    }

    public List<Step> steps() {
        return steps;
    }

    /** 1 ステップの間隔(ms)。最小 30ms にクランプ。 */
    public void setIntervalMillis(double ms) {
        this.intervalMillis = Math.max(30.0, ms);
    }

    public boolean isRunning() {
        return timeline != null && timeline.getStatus() == Animation.Status.RUNNING;
    }

    /**
     * キューを先頭から再生する。各ステップで {@code visitor} を呼び、完了時に {@code onFinished} を呼ぶ。
     * 既に再生中なら何もしない。FX スレッドで呼ぶこと。
     */
    public void play(StepVisitor visitor, Runnable onFinished) {
        if (visitor == null || steps.isEmpty() || isRunning()) {
            if (onFinished != null && steps.isEmpty()) {
                onFinished.run();
            }
            return;
        }
        stop();
        final int total = steps.size();
        Timeline tl = new Timeline();
        for (int i = 0; i < total; i++) {
            final int idx = i;
            tl.getKeyFrames()
                    .add(
                            new KeyFrame(
                                    Duration.millis(intervalMillis * (i + 1)),
                                    ev -> visitor.visit(steps.get(idx), idx, total)));
        }
        tl.setOnFinished(
                ev -> {
                    this.timeline = null;
                    if (onFinished != null) {
                        onFinished.run();
                    }
                });
        this.timeline = tl;
        tl.playFromStart();
    }

    /** 一時停止（{@link #resume()} で続行可能）。 */
    public void pause() {
        if (isRunning()) {
            timeline.pause();
        }
    }

    public void resume() {
        if (timeline != null && timeline.getStatus() == Animation.Status.PAUSED) {
            timeline.play();
        }
    }

    /** 停止して破棄。 */
    public void stop() {
        if (timeline != null) {
            timeline.stop();
            timeline = null;
        }
    }

    private static int parseQty(String raw) {
        if (raw == null) {
            return 0;
        }
        String s = raw.trim();
        if (s.isEmpty()) {
            return 0;
        }
        try {
            return (int) Math.round(Double.parseDouble(s.replace(",", "")));
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private static LocalDate parseDate(String raw) {
        if (raw == null) {
            return null;
        }
        String s = raw.trim();
        if (s.isEmpty()) {
            return null;
        }
        for (DateTimeFormatter f : DATE_FORMATS) {
            try {
                return LocalDate.parse(s, f);
            } catch (Exception ignore) {
                // try next
            }
        }
        return null;
    }

    private static LocalDateTime parseStart(String raw) {
        if (raw == null) {
            return null;
        }
        String s = raw.trim();
        if (s.isEmpty()) {
            return null;
        }
        for (DateTimeFormatter f : START_FORMATS) {
            try {
                return LocalDateTime.parse(s, f);
            } catch (Exception ignore) {
                // try next
            }
        }
        return null;
    }

    private static String trimmed(String raw) {
        return raw == null ? "" : raw.trim();
    }
}
