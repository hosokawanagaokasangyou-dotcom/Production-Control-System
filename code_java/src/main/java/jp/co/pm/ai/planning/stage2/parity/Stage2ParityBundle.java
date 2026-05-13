package jp.co.pm.ai.planning.stage2.parity;

import java.util.ArrayList;
import java.util.List;

/**
 * 同一検証の複数観点の結果。{@code null} の項目は「環境によりスキップ」で成功扱い。
 */
public record Stage2ParityBundle(
        Stage2ParityCheckResult planInputUiVsDisk,
        Stage2ParityCheckResult planPrimaryJson,
        Stage2ParityCheckResult memberJson,
        Stage2ParityCheckResult planWorkbookSemantics,
        Stage2ParityCheckResult memberWorkbookSemantics) {

    public boolean allPass() {
        return check(planInputUiVsDisk)
                && check(planPrimaryJson)
                && check(memberJson)
                && check(planWorkbookSemantics)
                && check(memberWorkbookSemantics);
    }

    private static boolean check(Stage2ParityCheckResult r) {
        return r == null || r.identical();
    }

    /** ダイアログ・ログ用の要約（不一致のみ列挙、すべて一致なら短文）。 */
    public String summary() {
        List<String> failures = new ArrayList<>();
        addIfFailed(failures, "配台計画タブの表と入力ファイル", planInputUiVsDisk);
        addIfFailed(failures, "計画 primary JSON", planPrimaryJson);
        addIfFailed(failures, "人員 JSON", memberJson);
        addIfFailed(failures, "計画ブック xlsx（セル内容）", planWorkbookSemantics);
        addIfFailed(failures, "人員ブック xlsx（セル内容）", memberWorkbookSemantics);
        if (failures.isEmpty()) {
            return "すべての検証項目が一致しました。\n\n"
                    + "（計画／人員の JSON・xlsx、配台計画タブの表と PM_AI_PLAN_INPUT_PATH の内容）";
        }
        return "以下の項目で不一致があります:\n\n- "
                + String.join("\n- ", failures)
                + "\n\n詳細は実行ログの [stage2-parity] を参照してください。";
    }

    private static void addIfFailed(
            List<String> failures, String label, Stage2ParityCheckResult r) {
        if (r != null && !r.identical()) {
            failures.add(label);
        }
    }

    /** ログ用: 各項目の一行サマリ。 */
    public List<String> logLines() {
        List<String> lines = new ArrayList<>();
        lines.add(line("plan_input_ui_vs_file", planInputUiVsDisk));
        lines.add(line("plan_primary_json", planPrimaryJson));
        lines.add(line("member_json", memberJson));
        lines.add(line("plan_xlsx_semantics", planWorkbookSemantics));
        lines.add(line("member_xlsx_semantics", memberWorkbookSemantics));
        return lines;
    }

    private static String line(String key, Stage2ParityCheckResult r) {
        if (r == null) {
            return "[stage2-parity] " + key + ": スキップ";
        }
        return "[stage2-parity] "
                + key
                + ": "
                + (r.identical() ? "一致" : "不一致")
                + " — "
                + firstLine(r.summary());
    }

    private static String firstLine(String s) {
        if (s == null) {
            return "";
        }
        int nl = s.indexOf('\n');
        return nl < 0 ? s : s.substring(0, nl);
    }
}
