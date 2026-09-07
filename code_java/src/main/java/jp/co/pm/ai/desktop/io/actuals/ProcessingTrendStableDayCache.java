package jp.co.pm.ai.desktop.io.actuals;

import java.time.LocalDate;
import java.util.Map;
import java.util.Objects;
import java.util.TreeMap;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 加工トレンドの「現在日より 30 日以上前」の日次バケットを再利用するメモリキャッシュ。
 *
 * <p>最新ソースファイルの mtime が変わっても、当該日付の数値は変化しない前提（業務要件）。
 * 累計・見込・移動平均はキャッシュせず、呼び出し側で毎回再合成する。
 *
 * <p>キーに mtime / size は含めない（パス同一性のみ）。工場切替・段階1キャッシュクリアで {@link #clearAll()}。
 */
public final class ProcessingTrendStableDayCache {

    /** {@code today.minusDays(STABLE_LOOKBACK_DAYS)} 以前（含む）を安定区間とする。 */
    public static final int STABLE_LOOKBACK_DAYS = 30;

    /** 1 日分の主実績・予定・比較実績（m）。 */
    public record DayValues(double primaryActualM, double planM, double compareActualM) {}

    /**
     * 系列キー。機械・工程は正規化済み（空文字＝すべて）。
     *
     * @param pathIdentity {@link EquipmentStatusDashboardSourceLoader#pathIdentity} 相当（mtime なし）
     */
    public record SeriesKey(
            String pathIdentity,
            ProcessingTrendAggregator.ActualSource actualSource,
            ProcessingTrendAggregator.PlanSource planSource,
            String machineNorm,
            String processNorm) {

        public SeriesKey {
            pathIdentity = pathIdentity != null ? pathIdentity : "";
            machineNorm = machineNorm != null ? machineNorm : "";
            processNorm = processNorm != null ? processNorm : "";
            Objects.requireNonNull(actualSource, "actualSource");
            Objects.requireNonNull(planSource, "planSource");
        }
    }

    private static final ProcessingTrendStableDayCache SHARED = new ProcessingTrendStableDayCache();

    private final ConcurrentHashMap<SeriesKey, ConcurrentHashMap<LocalDate, DayValues>> store =
            new ConcurrentHashMap<>();

    public static ProcessingTrendStableDayCache shared() {
        return SHARED;
    }

    public static LocalDate stableEndInclusive(LocalDate today) {
        LocalDate t = today != null ? today : LocalDate.now();
        return t.minusDays(STABLE_LOOKBACK_DAYS);
    }

    public void clearAll() {
        store.clear();
    }

    /**
     * 安定区間の日次値を {@code byDay} に流し込む。
     *
     * @return 期間内の安定日がすべてキャッシュにあれば {@code true}（ゼロ日もエントリ必須）
     */
    public boolean tryFillStableDays(
            SeriesKey key,
            LocalDate from,
            LocalDate to,
            LocalDate today,
            TreeMap<LocalDate, double[]> byDay) {
        if (key == null || byDay == null || from == null || to == null) {
            return false;
        }
        LocalDate stableEnd = stableEndInclusive(today);
        ConcurrentHashMap<LocalDate, DayValues> series = store.get(key);
        if (series == null || series.isEmpty()) {
            return false;
        }
        LocalDate end = to.isBefore(stableEnd) ? to : stableEnd;
        if (from.isAfter(end)) {
            return true; // 期間がすべて recent → 安定埋め不要
        }
        for (LocalDate d = from; !d.isAfter(end); d = d.plusDays(1)) {
            DayValues v = series.get(d);
            if (v == null) {
                return false;
            }
            double[] slot = byDay.get(d);
            if (slot == null) {
                continue;
            }
            slot[0] = v.primaryActualM();
            slot[1] = v.planM();
            slot[2] = v.compareActualM();
        }
        return true;
    }

    /** 移動平均用の期間前バッファを安定キャッシュから埋める。 */
    public void fillPriorActuals(
            SeriesKey key,
            Map<LocalDate, double[]> priorActuals,
            LocalDate today) {
        if (key == null || priorActuals == null || priorActuals.isEmpty()) {
            return;
        }
        ConcurrentHashMap<LocalDate, DayValues> series = store.get(key);
        if (series == null) {
            return;
        }
        LocalDate stableEnd = stableEndInclusive(today);
        for (Map.Entry<LocalDate, double[]> e : priorActuals.entrySet()) {
            LocalDate d = e.getKey();
            if (d == null || d.isAfter(stableEnd)) {
                continue;
            }
            DayValues v = series.get(d);
            if (v != null) {
                e.getValue()[0] = v.primaryActualM();
            }
        }
    }

    /** 集計成功後に安定日を保存する（ゼロ日も含む）。 */
    public void putStableDays(
            SeriesKey key,
            TreeMap<LocalDate, double[]> byDay,
            LocalDate today) {
        if (key == null || byDay == null) {
            return;
        }
        LocalDate stableEnd = stableEndInclusive(today);
        ConcurrentHashMap<LocalDate, DayValues> series =
                store.computeIfAbsent(key, k -> new ConcurrentHashMap<>());
        for (Map.Entry<LocalDate, double[]> e : byDay.entrySet()) {
            LocalDate d = e.getKey();
            if (d == null || d.isAfter(stableEnd)) {
                continue;
            }
            double[] slot = e.getValue();
            if (slot == null || slot.length < 3) {
                continue;
            }
            series.put(d, new DayValues(slot[0], slot[1], slot[2]));
        }
    }
}
