package jp.co.pm.ai.desktop.dispatch;

import java.time.LocalDate;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** 段階3.5 ウィザード内の編集状態（master 基準値と現在値）。 */
public final class OvertimeSimulationEditState {

    /** 残業時間入力の刻み（分）。 */
    public static final int OVERTIME_MINUTES_STEP = 15;

    /** 残業時間の上限（分）。 */
    public static final int OVERTIME_MINUTES_MAX = 720;

    public record CellState(
            boolean baselineWorking,
            boolean currentWorking,
            int baselineOvertimeMinutes,
            int currentOvertimeMinutes,
            boolean overtimeEdited) {}

    private final List<String> members;
    private final List<LocalDate> dates;
    private final Map<LocalDate, Map<String, CellState>> cells = new LinkedHashMap<>();

    public OvertimeSimulationEditState(AttendanceOvertimePreview.Preview preview) {
        this.members = List.copyOf(preview.members());
        this.dates = List.copyOf(preview.dates());
        for (LocalDate d : dates) {
            Map<String, CellState> row = new LinkedHashMap<>();
            for (String m : members) {
                AttendanceOvertimePreview.CellInfo info =
                        preview.cells().getOrDefault(d, Map.of()).get(m);
                boolean working = info != null && info.working();
                int ot = info != null ? info.overtimeMinutes() : 0;
                row.put(
                        m,
                        new CellState(working, working, ot, ot, false));
            }
            cells.put(d, row);
        }
    }

    public List<String> members() {
        return members;
    }

    public List<LocalDate> dates() {
        return dates;
    }

    public CellState cell(LocalDate date, String member) {
        return cells.getOrDefault(date, Map.of()).get(member);
    }

    public void toggleWorking(LocalDate date, String member) {
        Map<String, CellState> row = cells.get(date);
        if (row == null) {
            return;
        }
        CellState cs = row.get(member);
        if (cs == null) {
            return;
        }
        boolean next = !cs.currentWorking();
        row.put(
                member,
                new CellState(
                        cs.baselineWorking(),
                        next,
                        cs.baselineOvertimeMinutes(),
                        next ? cs.currentOvertimeMinutes() : 0,
                        cs.overtimeEdited()));
    }

    public void setOvertimeMinutes(LocalDate date, String member, int minutes) {
        Map<String, CellState> row = cells.get(date);
        if (row == null) {
            return;
        }
        CellState cs = row.get(member);
        if (cs == null || !cs.currentWorking()) {
            return;
        }
        int clamped = snapOvertimeMinutes(minutes);
        row.put(
                member,
                new CellState(
                        cs.baselineWorking(),
                        cs.currentWorking(),
                        cs.baselineOvertimeMinutes(),
                        clamped,
                        true));
    }

    public boolean hasChanges() {
        for (LocalDate d : dates) {
            for (String m : members) {
                CellState cs = cell(d, m);
                if (cs == null) {
                    continue;
                }
                if (cs.currentWorking() != cs.baselineWorking()) {
                    return true;
                }
                if (cs.overtimeEdited()) {
                    return true;
                }
            }
        }
        return false;
    }

    public String buildSummaryText(String noChangeSuffix) {
        StringBuilder sb = new StringBuilder();
        int workOn = 0;
        int workOff = 0;
        int otCount = 0;
        for (LocalDate d : dates) {
            for (String m : members) {
                CellState cs = cell(d, m);
                if (cs == null) {
                    continue;
                }
                if (!cs.baselineWorking() && cs.currentWorking()) {
                    workOn++;
                } else if (cs.baselineWorking() && !cs.currentWorking()) {
                    workOff++;
                }
                if (cs.overtimeEdited() && cs.currentWorking()) {
                    otCount++;
                }
            }
        }
        sb.append("休日出勤（○化）: ").append(workOn).append(" セル\n");
        sb.append("休日扱い（グレー化）: ").append(workOff).append(" セル\n");
        sb.append("残業時間の変更: ").append(otCount).append(" セル\n");
        if (!hasChanges() && noChangeSuffix != null && !noChangeSuffix.isBlank()) {
            sb.append(noChangeSuffix);
        }
        return sb.toString();
    }

    /** ウィザード入力: 0 または {@link #OVERTIME_MINUTES_STEP} 分刻みで {@link #OVERTIME_MINUTES_MAX} まで。 */
    public static int snapOvertimeMinutes(int minutes) {
        if (minutes <= 0) {
            return 0;
        }
        int snapped = (int) Math.round(minutes / (double) OVERTIME_MINUTES_STEP) * OVERTIME_MINUTES_STEP;
        return Math.min(OVERTIME_MINUTES_MAX, Math.max(0, snapped));
    }
}
