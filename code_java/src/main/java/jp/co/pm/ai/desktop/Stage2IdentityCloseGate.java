package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.AladdinEntryDispatchPlanIdentityCheck;

/**
 * この起動で段階2が完了したあと、ローカル最新配台計画と加工計画が不一致なら終了時に確認が必要。
 */
public final class Stage2IdentityCloseGate {

    public record Decision(boolean required, String detail) {}

    private boolean stage2CompletedThisLaunch;

    private boolean excelExportFailedAfterStage2;

    public void markStage2Completed() {
        markStage2Completed(true);
    }

    public void markStage2Completed(boolean excelExportSucceeded) {
        stage2CompletedThisLaunch = true;
        if (!excelExportSucceeded) {
            excelExportFailedAfterStage2 = true;
        }
    }

    public boolean stage2CompletedThisLaunch() {
        return stage2CompletedThisLaunch;
    }

    public Decision decide(Map<String, String> ui) {
        if (!stage2CompletedThisLaunch) {
            return new Decision(false, "");
        }
        if (excelExportFailedAfterStage2) {
            return new Decision(true, "Excel出力失敗");
        }
        Path excel = AppPaths.aladdinEntryDispatchPlanLocalXlsxPath(ui);
        AladdinEntryDispatchPlanIdentityCheck.Result result =
                AladdinEntryDispatchPlanIdentityCheck.evaluate(ui, excel);
        if (!result.error() && result.identical()) {
            return new Decision(false, "");
        }
        String detail =
                result.error()
                        ? "比較失敗"
                        : (result.badgeText() != null ? result.badgeText() : "差異あり");
        return new Decision(true, detail);
    }
}
