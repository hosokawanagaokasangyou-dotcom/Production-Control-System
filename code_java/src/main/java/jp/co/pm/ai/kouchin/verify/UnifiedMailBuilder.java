package jp.co.pm.ai.kouchin.verify;

import java.util.ArrayList;
import java.util.List;

/**
 * 国分工場・湖南工場の2工場分をまとめた東レ宛 月次報告メールの下書きを生成する。
 * 片方の検証結果が無い場合、その工場の行は「（未検証・要再実行）」になる。
 */
public final class UnifiedMailBuilder {

    /** 未検証の工場に表示する文言。 */
    public static final String NOT_VERIFIED = "（未検証・要再実行）";
    /** コピー範囲の開始目印。 */
    public static final String COPY_MARKER = "▼▼ ここから下をコピーしてメール本文に貼り付け ▼▼";

    private UnifiedMailBuilder() {
    }

    /** メール下書きの各行を返す（国分→湖南の順は実際のメールに合わせて固定）。 */
    public static List<String> buildLines(MailSnapshot kokubu, MailSnapshot konan) {
        List<String> lines = new ArrayList<>();
        lines.add("【メール下書き】数字は本検証結果から自動反映。送信前に内容を確認・修正してください。");

        addNotes(lines, kokubu, "国分工場");
        addNotes(lines, konan, "湖南工場");

        lines.add(COPY_MARKER);
        lines.add("東レ株式会社");
        lines.add("トーレペフ事業部 御中");
        lines.add("");
        lines.add("いつも大変お世話になっております。");
        lines.add("");
        lines.add("掲題の件、下記にご報告申し上げます。");
        lines.add("");
        lines.add(kokubu != null ? kokubu.amountLine() : "国分工場　" + NOT_VERIFIED);
        lines.add(konan != null ? konan.amountLine() : "湖南工場　" + NOT_VERIFIED);
        lines.add("");
        lines.add("上記お支払いデータに対し");
        lines.add(kokubu != null ? kokubu.diffLine() : "国分工場　" + NOT_VERIFIED);
        lines.add(konan != null ? konan.diffLine() : "湖南工場　" + NOT_VERIFIED);
        lines.add("");
        lines.add("添付ファイルをご確認いただき、次月（" + nextLabel(kokubu, konan) + "）での");
        lines.add("ご調整をよろしくお願い申し上げます。");
        lines.add("今後とも何卒よろしくお願い申し上げます。");
        lines.add("");
        lines.add("以上");
        return lines;
    }

    /** メール下書きのテキスト。 */
    public static String buildText(MailSnapshot kokubu, MailSnapshot konan) {
        return String.join(System.lineSeparator(), buildLines(kokubu, konan));
    }

    /** {@link #COPY_MARKER} より上にある注意書きの行数。 */
    public static int noteLineCount(List<String> lines) {
        for (int i = 0; i < lines.size(); i++) {
            if (lines.get(i).startsWith("▼▼")) {
                return i + 1;
            }
        }
        return 0;
    }

    private static void addNotes(List<String> lines, MailSnapshot s, String label) {
        if (s == null) {
            lines.add("※" + label + "は未検証です。" + label + "側の検証を実行してから送信してください"
                    + "（本文の該当行は" + NOT_VERIFIED + "のままです）。");
            return;
        }
        lines.add("※" + label + " 当月差異の内訳: 金額不一致 " + s.mismatchCount() + "件 "
                + Fmt.n0(s.mismatchAmount()) + "円 / 翌月記載(②の記載月誤り) " + s.nextCount() + "件 "
                + Fmt.n0(s.nextAmount()) + "円 / ①のみ(②記載なし・未解明) " + s.only1Count() + "件 "
                + Fmt.n0(s.only1Amount()) + "円 / ②のみ(①検収なし) " + s.only2Count() + "件 "
                + Fmt.n0(-s.only2Amount()) + "円。未解明分の扱いは送信前に要確認。");
        if (s.nextCount() > 0) {
            lines.add("※" + label + " 翌月記載 " + s.nextCount()
                    + "件は②の2ファイル修正(当月へ追記・翌月から削除)で解消する見込み。"
                    + "②修正後に本検証を再実行し、数字を確定させてからメールを作成すること。");
        }
    }

    private static String nextLabel(MailSnapshot kokubu, MailSnapshot konan) {
        if (kokubu != null && kokubu.nextYmLabel() != null && !kokubu.nextYmLabel().isEmpty()) {
            return kokubu.nextYmLabel();
        }
        if (konan != null && konan.nextYmLabel() != null && !konan.nextYmLabel().isEmpty()) {
            return konan.nextYmLabel();
        }
        return "翌月";
    }
}
