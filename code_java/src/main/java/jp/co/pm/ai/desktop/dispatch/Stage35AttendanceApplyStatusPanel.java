package jp.co.pm.ai.desktop.dispatch;

import java.nio.file.Path;

import javafx.scene.control.Label;
import javafx.scene.layout.VBox;

/**
 * 段階3.5 勤怠適用状態を配台計画手動修正タブに表示するパネル更新。
 */
public final class Stage35AttendanceApplyStatusPanel {

    private Stage35AttendanceApplyStatusPanel() {}

    public record ViewModel(
            boolean visible,
            String headline,
            String summaryLine,
            String overridesJsonPath,
            String pythonApplyLine,
            String appliedAt) {}

    public static ViewModel build(
            Stage35BaselineActualSnapshotStore.Stage35TrialMeta meta,
            Path dispatchJsonPath,
            String shortagesNote) {
        if (meta == null || !meta.hasTrialApplied()) {
            return hidden();
        }
        Stage35BaselineActualSnapshotStore.OverrideSummary summary =
                meta.overrideSummary() != null
                        ? meta.overrideSummary()
                        : Stage35BaselineActualSnapshotStore.OverrideSummary.empty();
        String overridesPath =
                meta.overtimeOverridesJson() != null && !meta.overtimeOverridesJson().isBlank()
                        ? meta.overtimeOverridesJson()
                        : defaultOverridesPath(dispatchJsonPath);
        boolean pythonApplied =
                shortagesNote != null
                        && shortagesNote.contains("残業シミュレーション適用:");
        String pythonLine =
                pythonApplied
                        ? "Python 配台試行: 残業シミュレーション JSON を適用済み"
                        : "Python 配台試行: 適用記録なし（不足 JSON 未読または段階3試行）";
        String appliedAt =
                meta.appliedAt() != null && !meta.appliedAt().isBlank()
                        ? "試行日時: " + meta.appliedAt()
                        : "";
        return new ViewModel(
                true,
                "段階3.5 勤怠適用済",
                summary.formatSummaryLine(),
                overridesPath,
                pythonLine,
                appliedAt);
    }

    public static void apply(VBox panel, Label headline, Label summary, Label overrides, Label python, Label appliedAt, ViewModel vm) {
        if (panel == null) {
            return;
        }
        boolean vis = vm != null && vm.visible();
        panel.setVisible(vis);
        panel.setManaged(vis);
        if (!vis) {
            return;
        }
        if (headline != null) {
            headline.setText(vm.headline());
        }
        if (summary != null) {
            summary.setText(vm.summaryLine());
        }
        if (overrides != null) {
            overrides.setText(
                    vm.overridesJsonPath() != null && !vm.overridesJsonPath().isBlank()
                            ? "勤怠上書き JSON: " + vm.overridesJsonPath()
                            : "");
        }
        if (python != null) {
            python.setText(vm.pythonApplyLine());
        }
        if (appliedAt != null) {
            appliedAt.setText(vm.appliedAt() != null ? vm.appliedAt() : "");
        }
    }

    private static ViewModel hidden() {
        return new ViewModel(false, "", "", "", "", "");
    }

    private static String defaultOverridesPath(Path dispatchJsonPath) {
        if (dispatchJsonPath == null) {
            return "";
        }
        Path parent = dispatchJsonPath.getParent();
        if (parent == null) {
            return "overtime_simulation_overrides.json";
        }
        return parent.resolve("overtime_simulation_overrides.json").toAbsolutePath().normalize().toString();
    }
}
