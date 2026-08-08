package jp.co.pm.ai.desktop.ui;

import java.text.Collator;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus;
import jp.co.pm.ai.desktop.io.actuals.EquipmentMachineStatus.Status;

/** ダッシュボードのカード一覧に対する絞込・並べ替え・件数集計（JavaFX 非依存）。 */
public final class EquipmentStatusDashboardFilter {

    /** カードの並び順。 */
    public enum SortOrder {
        MACHINE_NAME("機械名"),
        STOPPED_FIRST("停機を先頭"),
        COMPLETION_ASC("達成率が低い順");

        private final String label;

        SortOrder(String label) {
            this.label = label;
        }

        public String label() {
            return label;
        }

        public static SortOrder fromLabel(String label) {
            if (label != null) {
                for (SortOrder o : values()) {
                    if (o.label.equals(label.strip())) {
                        return o;
                    }
                }
            }
            return MACHINE_NAME;
        }
    }

    /** 状態別の機械台数。 */
    public record StatusCounts(int running, int stopped, int completed) {

        public int total() {
            return running + stopped + completed;
        }

        public int of(Status status) {
            if (status == null) {
                return 0;
            }
            return switch (status) {
                case RUNNING -> running;
                case STOPPED -> stopped;
                case COMPLETED -> completed;
            };
        }
    }

    /** 停機（実績なし）を並べ替え上「最下位の達成率」として扱うための番兵値。 */
    static final double NO_ACTUAL_COMPLETION_PCT = -1.0;

    private EquipmentStatusDashboardFilter() {}

    public static StatusCounts countByStatus(List<EquipmentMachineStatus> statuses) {
        int running = 0;
        int stopped = 0;
        int completed = 0;
        if (statuses != null) {
            for (EquipmentMachineStatus s : statuses) {
                if (s == null || s.status() == null) {
                    continue;
                }
                switch (s.status()) {
                    case RUNNING -> running++;
                    case STOPPED -> stopped++;
                    case COMPLETED -> completed++;
                }
            }
        }
        return new StatusCounts(running, stopped, completed);
    }

    /**
     * 状態フィルタ・機械名キーワードで絞り込み、指定順に並べ替える。
     *
     * @param allowedStatuses 空または {@code null} なら全状態を通す
     * @param machineKeyword 空または {@code null} なら絞り込まない。NFKC 正規化・大小無視・空白無視で部分一致
     */
    public static List<EquipmentMachineStatus> apply(
            List<EquipmentMachineStatus> statuses,
            Set<Status> allowedStatuses,
            String machineKeyword,
            SortOrder order) {
        if (statuses == null || statuses.isEmpty()) {
            return List.of();
        }
        boolean filterStatus = allowedStatuses != null && !allowedStatuses.isEmpty();
        String keyword = normalizeKeyword(machineKeyword);
        List<EquipmentMachineStatus> result = new ArrayList<>(statuses.size());
        for (EquipmentMachineStatus s : statuses) {
            if (s == null) {
                continue;
            }
            if (filterStatus && !allowedStatuses.contains(s.status())) {
                continue;
            }
            if (!keyword.isEmpty() && !normalizeKeyword(s.machineName()).contains(keyword)) {
                continue;
            }
            result.add(s);
        }
        result.sort(comparatorFor(order));
        return result;
    }

    static Comparator<EquipmentMachineStatus> comparatorFor(SortOrder order) {
        Comparator<EquipmentMachineStatus> byName =
                Comparator.comparing(
                        s -> s.machineName() != null ? s.machineName() : "",
                        Collator.getInstance(Locale.JAPAN));
        return switch (order != null ? order : SortOrder.MACHINE_NAME) {
            case STOPPED_FIRST ->
                    Comparator.comparingInt(
                                    (EquipmentMachineStatus s) -> statusPriority(s.status()))
                            .thenComparing(byName);
            case COMPLETION_ASC ->
                    Comparator.comparingDouble(EquipmentStatusDashboardFilter::sortCompletionPct)
                            .thenComparing(byName);
            case MACHINE_NAME -> byName;
        };
    }

    /** 停機を最優先で見せるための並べ替え優先度。 */
    static int statusPriority(Status status) {
        if (status == null) {
            return 9;
        }
        return switch (status) {
            case STOPPED -> 0;
            case RUNNING -> 1;
            case COMPLETED -> 2;
        };
    }

    /** 並べ替え用の達成率。実績が無い機械は {@link #NO_ACTUAL_COMPLETION_PCT}。 */
    static double sortCompletionPct(EquipmentMachineStatus status) {
        if (status == null || status.actualTask() == null) {
            return NO_ACTUAL_COMPLETION_PCT;
        }
        return status.actualTask()
                .map(EquipmentMachineStatus.ActualTaskRow::completionPct)
                .orElse(NO_ACTUAL_COMPLETION_PCT);
    }

    static String normalizeKeyword(String raw) {
        if (raw == null) {
            return "";
        }
        String normalized = Normalizer.normalize(raw.strip(), Normalizer.Form.NFKC);
        return normalized.toLowerCase(Locale.ROOT).replace(" ", "").replace("\u3000", "");
    }
}
