package jp.co.pm.ai.kouchin.verify;

import java.util.Set;

/**
 * 判定ラベルと、Excel 出力での並び順・色。
 * 色は Python 版の {@code FILL_*} / {@code JUDGE_COLOR} と同じ値。
 */
public final class Judge {

    public static final String MATCH = "一致";
    public static final String MISMATCH = "不一致";
    public static final String NEXT_MONTH = "翌月記載";
    public static final String ONLY_1 = "①のみ";
    public static final String ONLY_2 = "②のみ";
    public static final String ONLY_3 = "③のみ";
    public static final String PREV_ADJUST = "前月調整";
    public static final String PREV_GAP = "前月過不足";
    public static final String MANUAL_1 = "①正(手動)";
    public static final String MANUAL_2 = "②正(手動)";
    public static final String RESOLVED = "月ずれ解消";
    public static final String BRANCH_MERGED = "枝番統合一致";
    public static final String BAD_FORMAT = "形式不正";

    /** 検証D */
    public static final String AMOUNT_DIFF = "金額差";
    public static final String REF_SHIFT = "参照ずれ";
    public static final String NOT_MAPPED = "未取込";
    public static final String BAD_FORMULA = "式異常";
    public static final String HARDCODED = "直接入力";

    /** 検証Dの「要修正」判定。 */
    public static final Set<String> D_ERRORS = Set.of(AMOUNT_DIFF, REF_SHIFT, NOT_MAPPED);

    private Judge() {
    }

    /**
     * 並び順の主キー。
     * 不一致 → 翌月記載 → 片側のみ → 前月調整 → 月ずれ解消・枝番統合 → 形式不正 → 一致
     */
    public static int rank(String judge) {
        if (judge == null) {
            return 6;
        }
        return switch (judge) {
            case MISMATCH, AMOUNT_DIFF -> 0;
            case NEXT_MONTH, REF_SHIFT, NOT_MAPPED, BAD_FORMULA -> 1;
            case ONLY_1, ONLY_2, ONLY_3, HARDCODED -> 2;
            case PREV_ADJUST, PREV_GAP, MANUAL_1, MANUAL_2 -> 3;
            case RESOLVED, BRANCH_MERGED -> 4;
            case BAD_FORMAT -> 5;
            default -> 6;
        };
    }

    /** 行の塗り色（RGB 16進・色なしは null）。 */
    public static String fillColor(String judge) {
        if (judge == null) {
            return null;
        }
        return switch (judge) {
            case MISMATCH, AMOUNT_DIFF, REF_SHIFT, NOT_MAPPED -> "FFC7CE";
            case NEXT_MONTH, BAD_FORMULA -> "F8CBAD";
            case ONLY_1, ONLY_2, ONLY_3, HARDCODED -> "FFEB9C";
            case PREV_ADJUST -> "BDD7EE";
            case PREV_GAP -> "DDEBF7";
            case MANUAL_1, MANUAL_2 -> "E4DFEC";
            case RESOLVED, BRANCH_MERGED -> "E2EFDA";
            case BAD_FORMAT -> "E7E6E6";
            default -> null;
        };
    }

    /** 判定セルの文字色。 */
    public static String fontColor(String judge) {
        if (judge == null) {
            return null;
        }
        return switch (judge) {
            case MISMATCH, AMOUNT_DIFF, REF_SHIFT, NOT_MAPPED -> "9C0006";
            case NEXT_MONTH, BAD_FORMULA -> "833C00";
            case ONLY_1, ONLY_2, ONLY_3, HARDCODED -> "7F5F00";
            case PREV_ADJUST, PREV_GAP -> "1F4E79";
            case MANUAL_1, MANUAL_2 -> "5B2C6F";
            case RESOLVED, BRANCH_MERGED -> "375623";
            case BAD_FORMAT -> "595959";
            default -> null;
        };
    }
}
