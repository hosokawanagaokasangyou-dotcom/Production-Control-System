package jp.co.pm.ai.desktop.ui;

import java.util.List;
import java.util.Optional;

/** 配台計画タスク入力の列別セル編集分岐。 */
public final class PlanInputCellEditRouting {

    public static final String COL_LIMITED_OPERATOR = "担当OP_限定";

    public enum Editor {
        TEXT,
        LIMITED_OPERATOR_CHECKLIST
    }

    private PlanInputCellEditRouting() {}

    public static Editor editorFor(String columnTitle) {
        return COL_LIMITED_OPERATOR.equals(columnTitle)
                ? Editor.LIMITED_OPERATOR_CHECKLIST
                : Editor.TEXT;
    }

    public static String applyLimitedOperatorResult(
            String currentValue, Optional<List<String>> selectedNames) {
        return selectedNames
                .map(LimitedOperatorJsonCodec::encode)
                .orElse(currentValue != null ? currentValue : "");
    }
}
