package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import jp.co.pm.ai.desktop.reconciliation.RequestFormPipelineStatusService.PipelineStatusRow;

/**
 * 原本転記・計画確認タブで段階1実行前に確認が必要なデータ不足・不一致を判定する。
 */
public final class RequestFormPipelineIssueCheck {

    public enum IssueKind {
        RAW_INPUT_DATE_MISMATCH("原反投入日不一致"),
        TRANSFER_MISMATCH("転記未一致"),
        ALADDIN_MISSING("アラジン計画なし"),
        CONTRACT_NO_MISSING("契約NO未入力"),
        NO_ORIGINAL("依頼書原本なし");

        private final String label;

        IssueKind(String label) {
            this.label = label;
        }

        public String label() {
            return label;
        }
    }

    private RequestFormPipelineIssueCheck() {}

    public static List<IssueKind> detect(PipelineStatusRow row, boolean aladdinJsonAvailable) {
        if (row == null) {
            return List.of();
        }
        List<IssueKind> issues = new ArrayList<>();
        if (RawInputDateCrossSourceCheck.STATUS_MISMATCH.equals(
                row.rawInputDateCrossCheck() != null
                        ? row.rawInputDateCrossCheck().status()
                        : "")) {
            issues.add(IssueKind.RAW_INPUT_DATE_MISMATCH);
        }
        if (row.mismatchCount() > 0) {
            issues.add(IssueKind.TRANSFER_MISMATCH);
        }
        if (aladdinJsonAvailable && row.juchuRegistered() && !row.aladdinPresent()) {
            issues.add(IssueKind.ALADDIN_MISSING);
        }
        if (row.juchuRegistered() && isContractNoMissing(row.contractNoStatus())) {
            issues.add(IssueKind.CONTRACT_NO_MISSING);
        }
        if (row.juchuRegistered() && !row.originalPresent()) {
            issues.add(IssueKind.NO_ORIGINAL);
        }
        return List.copyOf(issues);
    }

    public static String formatSummary(List<IssueKind> issues) {
        if (issues == null || issues.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < issues.size(); i++) {
            if (i > 0) {
                sb.append("・");
            }
            sb.append(issues.get(i).label());
        }
        return sb.toString();
    }

    public static String formatConfirmedDisplay(boolean hasIssues, boolean confirmed) {
        if (!hasIssues) {
            return "―";
        }
        return confirmed ? "済" : "未";
    }

    static boolean isContractNoMissing(String contractNoStatus) {
        return Objects.equals("未入力", contractNoStatus != null ? contractNoStatus.strip() : "");
    }
}
