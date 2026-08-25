package jp.co.pm.ai.desktop;

import java.nio.file.Path;

import jp.co.pm.ai.desktop.ResultDispatchTableTabController.AladdinEntryExportOutcome;

/**
 * 「同一化チェック（ローカル最新）」は、先にアラジン入力用 Excel を生成してから比較する。
 */
public final class AladdinIdentityCheckLocalFlow {

    static final String ERROR_EXPORT_FAILED =
            "アラジン入力用Excelの生成に失敗したため、同一化チェックを実行できません。";

    static final String ERROR_NO_LATEST = "生成した配台計画 Excel が見つかりません。";

    public record NextStep(Path excelPath, String errorMessage) {

        public boolean canCheck() {
            return errorMessage == null && excelPath != null;
        }
    }

    private AladdinIdentityCheckLocalFlow() {}

    public static NextStep afterExport(AladdinEntryExportOutcome outcome) {
        if (outcome == null || !outcome.succeeded()) {
            Exception error = outcome != null ? outcome.error() : null;
            String detail =
                    error != null && error.getMessage() != null && !error.getMessage().isBlank()
                            ? error.getMessage()
                            : "原因不明";
            return new NextStep(null, ERROR_EXPORT_FAILED + "\n" + detail);
        }
        Path latest = outcome.result().latestPath();
        if (latest == null) {
            return new NextStep(null, ERROR_NO_LATEST);
        }
        return new NextStep(latest, null);
    }
}
