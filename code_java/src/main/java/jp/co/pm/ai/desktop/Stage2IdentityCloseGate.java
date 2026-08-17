package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.AladdinEntryDispatchPlanIdentityCheck;

/**
 * この起動で段階2が完了したあと、ローカル最新配台計画と加工計画が不一致なら終了時に確認が必要。
 */
public final class Stage2IdentityCloseGate {

    public record Decision(boolean required, String detail, String dialogBody) {
        public Decision {
            detail = detail != null ? detail : "";
            dialogBody = dialogBody != null ? dialogBody : detail;
        }
    }

    private boolean stage2CompletedThisLaunch;

    private boolean excelExportFailedAfterStage2;

    public void markStage2Completed() {
        markStage2Completed(true);
    }

    public void markStage2Completed(boolean excelExportSucceeded) {
        stage2CompletedThisLaunch = true;
        excelExportFailedAfterStage2 = !excelExportSucceeded;
    }

    /** Python 段階2成功直後。Excel 自動生成が終わるまで終了ゲート必須。 */
    public void markStage2PipelineAwaitingExcel() {
        stage2CompletedThisLaunch = true;
        excelExportFailedAfterStage2 = true;
    }

    public void markExcelExportSucceeded() {
        excelExportFailedAfterStage2 = false;
    }

    public boolean stage2CompletedThisLaunch() {
        return stage2CompletedThisLaunch;
    }

    public Decision decide(Map<String, String> ui) {
        if (!stage2CompletedThisLaunch) {
            return new Decision(false, "", "");
        }
        if (excelExportFailedAfterStage2) {
            return new Decision(
                    true,
                    "Excel出力失敗",
                    "アラジン入力用Excelの自動生成がまだ成功していません。\n"
                            + "ローカル最新が以前の加工計画と一致していても、今回の配台が未入力の可能性があります。\n"
                            + "納期管理ビューの配台結果で同一化チェックを実行してください。");
        }
        try {
            Path excel = AppPaths.aladdinEntryDispatchPlanLocalXlsxPath(ui);
            AladdinEntryDispatchPlanIdentityCheck.Result result =
                    AladdinEntryDispatchPlanIdentityCheck.evaluate(ui, excel);
            if (!result.error() && result.identical()) {
                return new Decision(false, "", "");
            }
            String detail =
                    result.error()
                            ? "比較失敗"
                            : (result.badgeText() != null ? result.badgeText() : "差異あり");
            String body =
                    result.error()
                            ? (result.message() != null ? result.message() : "比較に失敗しました")
                            : result.dialogBody();
            return new Decision(true, detail, body);
        } catch (RuntimeException ex) {
            return new Decision(
                    true,
                    "比較失敗",
                    "比較中にエラーが発生しました。納期管理ビューの配台結果で同一化チェックを実行してください。");
        }
    }
}
