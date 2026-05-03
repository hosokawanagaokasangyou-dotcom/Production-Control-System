package jp.co.pm.ai.desktop.dispatch;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Column names aligned with Python {@code RESULT_DISPATCH_TABLE_STATIC_HEADERS} plus dispatch columns.
 */
public final class ResultDispatchSchema {

    /** Same as Python {@code RESULT_TASK_COL_DISPATCH_TRIAL_ORDER} / ????_?z??\ ???? */
    public static final String COL_DISPATCH_TRIAL_ORDER = "\u914d\u53f0\u8a66\u884c\u9806\u756a";

    public static final String COL_PROCESS = "\u5de5\u7a0b\u540d";
    public static final String COL_MACHINE = "\u6a5f\u68b0\u540d";
    public static final String COL_DISPATCH_DATE = "\u914d\u53f0\u65e5";
    public static final String COL_DISPATCH_QTY = "\u5f53\u65e5\u914d\u53f0\u6570\u91cf";

    /** Static columns in pipeline order (excluding {@link #COL_DISPATCH_DATE} / {@link #COL_DISPATCH_QTY}). */
    public static final List<String> STATIC_HEADERS =
            List.of(
                    COL_DISPATCH_TRIAL_ORDER,
                    COL_PROCESS,
                    COL_MACHINE,
                    "\u53d7\u6ce8\u65e5",
                    "\u53d7\u6ce8NO",
                    "\u4f9d\u983cNO",
                    "\u54c1\u540d(\u539f\u53cd)",
                    "\u4f7f\u7528\u539f\u53cd",
                    "\u539f\u53cd\u6570",
                    "\u54c1\u540d(\u88fd\u54c1)",
                    "\u88fd\u54c1\u540d",
                    "\u63db\u7b97\u6570\u91cf",
                    "\u5b9f\u52a0\u5de5\u6570",
                    "\u52a0\u5de5\u5185\u5bb9",
                    "\u5728\u5eab\u5834\u6240",
                    "\u539f\u53cd\u6295\u5165\u65e5",
                    "\u6307\u5b9a\u7d0d\u671f",
                    "\u56de\u7b54\u7d0d\u671f",
                    "\u52a0\u5de5\u5b8c\u4e86\u65e5",
                    "\u52a0\u5de5\u5b8c\u4e86\u533a\u5206",
                    "\u5b9f\u51fa\u6765\u9ad8",
                    "\u8a08\u753b\u5408\u8a08",
                    "\u539f\u53cd\u6295\u5165\u5834\u6240");

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
            case "\u53d7\u6ce8\u65e5",
                    "\u539f\u53cd\u6295\u5165\u65e5",
                    "\u6307\u5b9a\u7d0d\u671f",
                    "\u56de\u7b54\u7d0d\u671f",
                    "\u52a0\u5de5\u5b8c\u4e86\u65e5",
                    COL_DISPATCH_DATE -> true;
            default -> false;
        };
    }
}
