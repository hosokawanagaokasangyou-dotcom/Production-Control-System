package jp.co.pm.ai.desktop.dispatch;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Column names aligned with Python {@code RESULT_DISPATCH_TABLE_STATIC_HEADERS} plus dispatch columns.
 */
public final class ResultDispatchSchema {

    /** Python {@code RESULT_TASK_COL_DISPATCH_TRIAL_ORDER} と同じ列名（配台試行順番）。 */
    public static final String COL_DISPATCH_TRIAL_ORDER = "配台試行順番";

    public static final String COL_PROCESS = "工程名";
    public static final String COL_MACHINE = "機械名";
    public static final String COL_DISPATCH_DATE = "配台日";
    public static final String COL_DISPATCH_QTY = "当日配台数量";

    /** Static columns in pipeline order (excluding {@link #COL_DISPATCH_DATE} / {@link #COL_DISPATCH_QTY}). */
    public static final List<String> STATIC_HEADERS =
            List.of(
                    COL_DISPATCH_TRIAL_ORDER,
                    COL_PROCESS,
                    COL_MACHINE,
                    "受注日",
                    "受注NO",
                    "依頼NO",
                    "品名(原反)",
                    "使用原反",
                    "原反数",
                    "品名(製品)",
                    "製品名",
                    "換算数量",
                    "実加工数",
                    "加工内容",
                    "在庫場所",
                    "原反投入日",
                    "指定納期",
                    "回答納期",
                    "加工完了日",
                    "加工完了区分",
                    "実出来高",
                    "計画合計",
                    "原反投入場所",
                    "加工開始日時",
                    "加工終了日時");

    private ResultDispatchSchema() {}

    /** Full column list for JSON {@code columns} when serializing canonical rows. */
    public static List<String> canonicalColumnOrder() {
        List<String> out = new ArrayList<>(STATIC_HEADERS.size() + 2);
        out.addAll(STATIC_HEADERS);
        out.add(COL_DISPATCH_DATE);
        out.add(COL_DISPATCH_QTY);
        return Collections.unmodifiableList(out);
    }

    /** Same set as Python {@code RESULT_DISPATCH_TABLE_DATE_HEADERS} (yyyy/MM/dd display in Excel). */
    public static boolean isDateColumn(String col) {
        if (col == null) {
            return false;
        }
        return switch (col) {
            case "受注日",
                    "原反投入日",
                    "指定納期",
                    "回答納期",
                    "加工完了日",
                    COL_DISPATCH_DATE -> true;
            default -> false;
        };
    }
}
