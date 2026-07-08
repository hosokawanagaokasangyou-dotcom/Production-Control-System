package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.List;

/**
 * アラジン・受注ファイル・依頼書原本目次・依頼書原本シートの4ソースの原反投入日が
 * すべて一致するかを厳格照合する。
 *
 * <p>判定は厳格: 4ソースすべてに非空の日付があり、{@link JuchuTransferDateMatcher} で
 * すべて等価なときのみ「一致」。1つでも欠落または相違があれば「不一致」。4ソースいずれも
 * 空のときは照合対象外として「―」を返す。
 */
public final class RawInputDateCrossSourceCheck {

    public static final String STATUS_MATCH = "一致";
    public static final String STATUS_MISMATCH = "不一致";
    public static final String STATUS_NA = "―";

    /** ソース種別（表示ラベル付き）。 */
    public enum Source {
        ALADDIN("アラジン"),
        JUCHU("受注ファイル"),
        INDEX("依頼書原本 目次"),
        SHEET("依頼書原本 シート");

        private final String label;

        Source(String label) {
            this.label = label;
        }

        public String label() {
            return label;
        }
    }

    /** 4ソースの原反投入日（strip 済み表示値）。 */
    public record SourceValues(String aladdin, String juchu, String index, String sheet) {}

    /** 照合結果。 */
    public record CrossSourceResult(String status, SourceValues values, String detailSummary) {

        public boolean matched() {
            return STATUS_MATCH.equals(status);
        }

        public boolean mismatched() {
            return STATUS_MISMATCH.equals(status);
        }
    }

    private RawInputDateCrossSourceCheck() {}

    /**
     * 4ソースの原反投入日を厳格照合する。
     *
     * @param aladdinJsonAvailable アラジン shaped JSON が読込済みか（未読込なら照合不能）
     */
    public static CrossSourceResult evaluate(
            String aladdin,
            String juchu,
            String index,
            String sheet,
            boolean aladdinJsonAvailable) {
        String a = strip(aladdin);
        String j = strip(juchu);
        String i = strip(index);
        String s = strip(sheet);
        SourceValues values = new SourceValues(a, j, i, s);

        if (a.isEmpty() && j.isEmpty() && i.isEmpty() && s.isEmpty()) {
            String reason =
                    aladdinJsonAvailable
                            ? "4ソースいずれも原反投入日なし"
                            : "アラジン未読込・原反投入日なし";
            return new CrossSourceResult(STATUS_NA, values, reason);
        }

        boolean allPresent = !a.isEmpty() && !j.isEmpty() && !i.isEmpty() && !s.isEmpty();
        if (!allPresent) {
            return new CrossSourceResult(
                    STATUS_MISMATCH, values, missingSummary(a, j, i, s, aladdinJsonAvailable));
        }

        boolean equal =
                JuchuTransferDateMatcher.datesMatch(a, j)
                        && JuchuTransferDateMatcher.datesMatch(a, i)
                        && JuchuTransferDateMatcher.datesMatch(a, s);
        if (equal) {
            return new CrossSourceResult(STATUS_MATCH, values, "4ソースの原反投入日が一致");
        }
        return new CrossSourceResult(STATUS_MISMATCH, values, diffSummary(values));
    }

    private static String missingSummary(
            String a, String j, String i, String s, boolean aladdinJsonAvailable) {
        List<String> missing = new ArrayList<>();
        if (a.isEmpty()) {
            missing.add(aladdinJsonAvailable ? Source.ALADDIN.label() : Source.ALADDIN.label() + "（未読込）");
        }
        if (j.isEmpty()) {
            missing.add(Source.JUCHU.label());
        }
        if (i.isEmpty()) {
            missing.add(Source.INDEX.label());
        }
        if (s.isEmpty()) {
            missing.add(Source.SHEET.label());
        }
        return "原反投入日が欠落: " + String.join("、", missing);
    }

    private static String diffSummary(SourceValues v) {
        return "原反投入日が相違: "
                + Source.ALADDIN.label() + "=" + display(v.aladdin())
                + " / " + Source.JUCHU.label() + "=" + display(v.juchu())
                + " / " + Source.INDEX.label() + "=" + display(v.index())
                + " / " + Source.SHEET.label() + "=" + display(v.sheet());
    }

    private static String display(String v) {
        String t = strip(v).replace("\n", " ");
        return t.isEmpty() ? "（空）" : t;
    }

    private static String strip(String v) {
        return v != null ? v.strip() : "";
    }
}
