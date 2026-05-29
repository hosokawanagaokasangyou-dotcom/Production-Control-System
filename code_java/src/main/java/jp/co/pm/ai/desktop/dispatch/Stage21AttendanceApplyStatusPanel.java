package jp.co.pm.ai.desktop.dispatch;

import java.nio.file.Path;

import javafx.scene.control.Label;
import javafx.scene.layout.VBox;

/** 段階2.1 勤怠適用状態を配台計画手動修正タブに表示するパネル更新。 */
public final class Stage21AttendanceApplyStatusPanel {

    private Stage21AttendanceApplyStatusPanel() {}

    public record ViewModel(
            boolean visible,
            String headline,
            String summaryLine,
            String overridesJsonPath,
            String pythonApplyLine,
            String appliedAt) {}

    public static ViewModel build(
            Stage21TrialSnapshotStore.Stage21TrialMeta meta,
            Path dispatchJsonPath,
            String shortagesNote) {
        if (meta == null || !meta.hasAttendanceMeta()) {
            return hidden();
        }
        Stage21TrialSnapshotStore.OverrideSummary summary =
                meta.overrideSummary() != null
                        ? meta.overrideSummary()
                        : Stage21TrialSnapshotStore.OverrideSummary.empty();
        String overridesPath =
                meta.overtimeOverridesJson() != null && !meta.overtimeOverridesJson().isBlank()
                        ? meta.overtimeOverridesJson()
                        : defaultOverridesPath(meta, dispatchJsonPath);
        boolean pythonApplied =
                shortagesNote != null && shortagesNote.contains("残業シミュレーション適用:");
        String pythonLine =
                pythonApplied
                        ? "Python 段階2.1: 残業/休出シミュ JSON を適用済み（メイン output へ正本反映）"
                        : meta.hasPromotedToMain()
                                ? "Python 段階2.1: メイン output へ正本反映済み"
                                : "Python 段階2.1: 適用記録なし";
        String appliedAt =
                meta.appliedAt() != null && !meta.appliedAt().isBlank()
                        ? "試行日時: " + meta.appliedAt()
                        : "";
        return new ViewModel(
                true,
                "段階2.1 勤怠適用済",
                summary.formatSummaryLine(),
                overridesPath,
                pythonLine,
                appliedAt);
    }

    public static void apply(
            VBox panel,
            Label headline,
            Label summary,
            Label overrides,
            Label python,
            Label appliedAt,
            ViewModel vm) {
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

    private static String defaultOverridesPath(
            Stage21TrialSnapshotStore.Stage21TrialMeta meta, Path dispatchJsonPath) {
        Path p = meta != null ? meta.overtimeOverridesPath() : null;
        if (p != null) {
            return p.toAbsolutePath().normalize().toString();
        }
        if (dispatchJsonPath == null) {
            return "";
        }
        Path parent = dispatchJsonPath.getParent();
        if (parent == null) {
            return "stage21/overtime_simulation_overrides.json";
        }
        return parent.resolve("stage21")
                .resolve("overtime_simulation_overrides.json")
                .toAbsolutePath()
                .normalize()
                .toString();
    }
}
