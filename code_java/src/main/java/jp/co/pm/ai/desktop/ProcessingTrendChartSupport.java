package jp.co.pm.ai.desktop;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * 加工トレンドのチャート描画で使う純粋関数（FX 非依存）。
 *
 * <ul>
 *   <li>Y 軸の「きれいな」上限・目盛間隔（1 / 2 / 2.5 / 5 × 10ⁿ）</li>
 *   <li>期間長に応じた X 軸ラベルの間引き対象日</li>
 *   <li>期間長に応じた棒の間隔</li>
 * </ul>
 */
public final class ProcessingTrendChartSupport {

    /** 日次表示で全日にラベルを出す上限。 */
    static final int LABEL_ALL_DAYS_MAX = 31;
    /** 月曜・月初にラベルを出す上限。これを超えると 1 日・15 日のみ。 */
    static final int LABEL_WEEKLY_DAYS_MAX = 120;
    /** Y 軸の目標目盛本数。 */
    static final int TARGET_TICKS = 6;

    private ProcessingTrendChartSupport() {}

    /** Y 軸の上限と目盛間隔。 */
    public record NiceRange(double upperBound, double tickUnit) {}

    /**
     * 最大値から上限・目盛間隔を決める。上限は最大値より少なくとも 0.15 目盛ぶん上に置き、
     * 最上段の棒・点が枠に張り付かないようにする。最大値が 0 以下なら {@code (10, 2)}。
     */
    public static NiceRange niceRange(double max) {
        if (Double.isNaN(max) || Double.isInfinite(max) || max <= 0) {
            return new NiceRange(10, 2);
        }
        double raw = max / TARGET_TICKS;
        double magnitude = Math.pow(10, Math.floor(Math.log10(raw)));
        double unit = 10 * magnitude;
        for (double m : new double[] {1, 2, 2.5, 5, 10}) {
            double candidate = m * magnitude;
            if (candidate >= raw) {
                unit = candidate;
                break;
            }
        }
        double upper = Math.ceil(max / unit) * unit;
        if (upper - max < unit * 0.15) {
            upper += unit;
        }
        return new NiceRange(upper, unit);
    }

    /** 週次間引き（〜120 日、約 9px/日）で隣り合うラベル同士に最低限あける日数（"M/d" 1 個ぶんの幅）。 */
    static final int LABEL_MIN_GAP_DAYS = 3;
    /** 半月間引き（120 日超、約 3px/日。年またぎで "yy/M/d" になる）でのラベル間の最低日数。 */
    static final int LABEL_MIN_GAP_DAYS_SEMIMONTHLY = 10;

    /**
     * X 軸にラベルを表示する日の集合。
     *
     * <ul>
     *   <li>{@value #LABEL_ALL_DAYS_MAX} 日以下: 全日</li>
     *   <li>{@value #LABEL_WEEKLY_DAYS_MAX} 日以下: 月曜と月初（先頭日も）</li>
     *   <li>それ以上: 1 日と 15 日（先頭日も）</li>
     * </ul>
     *
     * 間引きモードでは、候補同士が最低日数（週次 {@value #LABEL_MIN_GAP_DAYS} 日／半月
     * {@value #LABEL_MIN_GAP_DAYS_SEMIMONTHLY} 日）未満に近接したら
     * 優先度（月初 &gt; 月曜／15 日 &gt; 先頭日）の高い方だけ残し、文字の重なりを防ぐ。
     */
    public static Set<LocalDate> labelledDates(List<LocalDate> dates) {
        Set<LocalDate> out = new LinkedHashSet<>();
        int n = dates.size();
        if (n == 0) {
            return out;
        }
        if (n <= LABEL_ALL_DAYS_MAX) {
            out.addAll(dates);
            return out;
        }
        boolean weekly = n <= LABEL_WEEKLY_DAYS_MAX;
        int minGap = weekly ? LABEL_MIN_GAP_DAYS : LABEL_MIN_GAP_DAYS_SEMIMONTHLY;
        List<LocalDate> kept = new ArrayList<>();
        List<Integer> keptPriority = new ArrayList<>();
        for (int i = 0; i < n; i++) {
            LocalDate d = dates.get(i);
            int priority = labelPriority(d, i == 0, weekly);
            if (priority == 0) {
                continue;
            }
            int lastIdx = kept.size() - 1;
            if (lastIdx >= 0 && ChronoUnit.DAYS.between(kept.get(lastIdx), d) < minGap) {
                if (priority > keptPriority.get(lastIdx)) {
                    kept.set(lastIdx, d);
                    keptPriority.set(lastIdx, priority);
                }
                continue;
            }
            kept.add(d);
            keptPriority.add(priority);
        }
        out.addAll(kept);
        return out;
    }

    /** 0 = ラベル無し、大きいほど優先。 */
    private static int labelPriority(LocalDate d, boolean isFirst, boolean weekly) {
        int dom = d.getDayOfMonth();
        if (dom == 1) {
            return 3;
        }
        if (weekly ? d.getDayOfWeek() == DayOfWeek.MONDAY : dom == 15) {
            return 2;
        }
        return isFirst ? 1 : 0;
    }

    /** 棒グラフのカテゴリ間隔（px）。日数が増えるほど詰める。 */
    public static double categoryGapFor(int days) {
        if (days <= 31) {
            return 8;
        }
        if (days <= 62) {
            return 4;
        }
        if (days <= 120) {
            return 2;
        }
        return 1;
    }

    /** 同一カテゴリ内の棒同士の間隔（px）。 */
    public static double barGapFor(int days) {
        return days <= 62 ? 1 : 0;
    }
}
