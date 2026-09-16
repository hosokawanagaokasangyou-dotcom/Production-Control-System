package jp.co.pm.ai.kouchin.verify;

import jp.co.pm.ai.desktop.io.PoiWorkbookFileWriter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.GraphicsEnvironment;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * 検証結果Excel（XSSF）の書き出し。
 * シート構成: サマリ / レポートの見方 / 検証A / 検証B /（国分のみ）検証D / 報告メール下書き。
 */
public final class ResultExcelExporter {

    private static final String SHEET_SUMMARY = "サマリ";
    private static final String SHEET_GUIDE = "レポートの見方";
    private static final String SHEET_A = "検証A_契約NO(①vs②)";
    private static final String SHEET_B = "検証B_依頼NO(②vs③)";
    private static final String SHEET_D = "検証D_②まとめ整合性";
    private static final String SHEET_MAIL = "報告メール下書き";

    private static final String HEADER_FILL = "4472C4";
    private static final String ACCENT = "1F4E79";
    private static final String GRAY = "595959";
    private static final DateTimeFormatter FILE_STAMP = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");

    /** 見出し・本文フォント（無ければ游ゴシック） */
    private final String baseFont;
    /** 金額・件数用の等幅フォント */
    private final String numFont;

    private final XSSFWorkbook wb;
    private final Map<String, XSSFFont> fonts = new HashMap<>();
    private final Map<String, CellStyle> styles = new HashMap<>();
    private final short fmtInt;
    private final short fmtDec;

    private ResultExcelExporter(XSSFWorkbook wb) {
        this.wb = wb;
        this.baseFont = pickFont("BIZ UDPゴシック", "BIZ UDPGothic", "游ゴシック", "Yu Gothic");
        this.numFont = pickFont("BIZ UDゴシック", "BIZ UDGothic", "游ゴシック", "Yu Gothic");
        this.fmtInt = wb.createDataFormat().getFormat("#,##0;[Red]-#,##0");
        this.fmtDec = wb.createDataFormat().getFormat("#,##0.00;[Red]-#,##0.00");
    }

    /** 既定のファイル名（{@code 検証結果_国分工場_yyyyMMdd_HHmmss.xlsx}）。 */
    public static Path defaultOutputPath(KouchinPaths paths, FactoryProfile profile) {
        String name = "検証結果_" + profile.label() + "_" + LocalDateTime.now().format(FILE_STAMP) + ".xlsx";
        return paths.outputDir().resolve(name);
    }

    /** 他工場のメール数字は直近結果キャッシュから補う。 */
    public static Path export(VerifyResult result, Path outFile) throws IOException {
        return export(result, outFile, Map.of());
    }

    public static Path export(VerifyResult result, Path outFile, Map<String, String> ui)
            throws IOException {
        FactoryId other = result.factory() == FactoryId.KOKUBU ? FactoryId.KONAN : FactoryId.KOKUBU;
        MailSnapshot otherSnapshot = LastResultCache.load(other).orElse(null);
        MailSnapshot kokubu = result.factory() == FactoryId.KOKUBU ? result.mail() : otherSnapshot;
        MailSnapshot konan = result.factory() == FactoryId.KONAN ? result.mail() : otherSnapshot;
        return export(result, outFile, kokubu, konan, ui);
    }

    public static Path export(VerifyResult result, Path outFile, MailSnapshot kokubu, MailSnapshot konan)
            throws IOException {
        return export(result, outFile, kokubu, konan, Map.of());
    }

    public static Path export(
            VerifyResult result,
            Path outFile,
            MailSnapshot kokubu,
            MailSnapshot konan,
            Map<String, String> ui)
            throws IOException {
        try (XSSFWorkbook wb = buildWorkbook(result, kokubu, konan)) {
            save(outFile, wb, ui);
        }
        return outFile;
    }

    public static XSSFWorkbook buildWorkbook(VerifyResult result, MailSnapshot peer) {
        MailSnapshot kokubu = result.factory() == FactoryId.KOKUBU ? result.mail() : peer;
        MailSnapshot konan = result.factory() == FactoryId.KONAN ? result.mail() : peer;
        return buildWorkbook(result, kokubu, konan);
    }

    public static XSSFWorkbook buildWorkbook(VerifyResult result, MailSnapshot kokubu, MailSnapshot konan) {
        XSSFWorkbook wb = new XSSFWorkbook();
        ResultExcelExporter exporter = new ResultExcelExporter(wb);
        exporter.writeSummary(result);
        exporter.writeGuide(result);
        exporter.writeSheetA(result);
        exporter.writeSheetB(result);
        exporter.writeSheetD(result);
        exporter.writeSheetC(result);
        exporter.writeMail(kokubu, konan);
        exporter.writeRawSheets(result);
        return wb;
    }

    private static void save(Path outFile, XSSFWorkbook wb, Map<String, String> ui) throws IOException {
        Path parent = outFile.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        PoiWorkbookFileWriter.writeReplacing(outFile, wb, ui != null ? ui : Map.of());
    }

    // ---------------------------------------------------------------- サマリ

    private void writeSummary(VerifyResult r) {
        XSSFSheet ws = wb.createSheet(SHEET_SUMMARY);
        ws.setDisplayGridlines(false);
        ws.setTabColor(new XSSFColor(rgb(ACCENT), null));
        int[] widths = {2, 30, 24, 24, 24, 24, 24, 2};
        for (int i = 0; i < widths.length; i++) {
            ws.setColumnWidth(i, widths[i] * 256);
        }

        put(ws, 0, 1, "後加工工賃 金額整合性検証レポート【" + r.profile().label() + "】",
                style("title", ACCENT, null, true, 15, false, null));
        merge(ws, 0, 1, 6);
        put(ws, 1, 1, "対象月: " + r.str("対象月ラベル") + "　実行: " + r.str("実行日時")
                + " (" + r.str("実行者") + ")　許容差: ±" + r.dbl("許容差") + "円",
                style("sub", GRAY, null, false, 10, false, null));
        merge(ws, 1, 1, 6);

        // KPIカード（ラベル / 値 / 補足の3行）
        long tougetsu = r.num("報告する過不足");
        int warnCount = r.warnings().size();
        long zansa = r.num("報告残差");
        String[][] cards = {
                {"検証A 要確認 (①vs②)", r.requiredCheckA() + " 件",
                        "不一致" + r.num("A不一致") + "・翌月記載" + r.num("A翌月記載件数")
                                + "・片側のみ" + (r.num("A①のみ") + r.num("A②のみ")),
                        r.requiredCheckA() > 0 ? "bad" : "ok"},
                {"検証B 要確認 (②vs③)", r.requiredCheckB() + " 件",
                        "不一致" + r.num("B不一致") + "・片側のみ" + (r.num("B②のみ") + r.num("B③のみ")),
                        r.requiredCheckB() > 0 ? "bad" : "ok"},
                {"報告する過不足(当月)", Fmt.s0(tougetsu) + " 円", "検証A「報告計上額」列の合計", "neutral"},
                {"判明済み", r.knownCount() + " 件", "前月調整・翌月記載・月ずれ解消・枝番統合", "known"},
                {"警告", warnCount + " 件", "データ品質の注意", warnCount > 0 ? "bad" : "ok"},
                {"検算残差", Fmt.n0(zansa) + " 円", "0なら内訳の整合OK", zansa != 0 ? "bad" : "ok"},
        };
        for (int i = 0; i < cards.length; i++) {
            String[] palette = palette(cards[i][3]);
            int col = 1 + i;
            put(ws, 3, col, cards[i][0], style("kpiLabel" + cards[i][3], palette[1], palette[0], true, 10, true, null));
            put(ws, 4, col, cards[i][1], style("kpiValue" + cards[i][3], palette[1], palette[0], true, 16, true, null));
            put(ws, 5, col, cards[i][2], style("kpiSub" + cards[i][3], palette[1], palette[0], false, 8.5, true, null));
        }
        row(ws, 4).setHeightInPoints(24f);

        int rowIndex = 7;
        rowIndex = section(ws, rowIndex, "総額");
        rowIndex = kv(ws, rowIndex, "① 東レ検収 (お支払データ・" + r.str("入庫場所") + ")", r.num("①総額"));
        rowIndex = kv(ws, rowIndex, "② " + r.str("②名称") + " " + r.str("②金額列") + "合計", r.num("②総額"));
        rowIndex = kv(ws, rowIndex, "③ アラジン 加工金額合計", r.num("③総額"));
        rowIndex++;

        rowIndex = section(ws, rowIndex, "検証A　① vs ② (契約NO突合)");
        rowIndex = kv(ws, rowIndex, "共通契約NO", r.num("A共通") + " 件");
        rowIndex = kv(ws, rowIndex, "　一致 / 不一致", (r.num("A共通") - r.num("A不一致")) + " / " + r.num("A不一致"));
        rowIndex = kv(ws, rowIndex, "前月調整 (①のみ・過去月②に存在)",
                r.num("A前月調整件数") + "件 / " + Fmt.n0(r.dbl("A前月調整額")) + "円");
        rowIndex = kv(ws, rowIndex, "翌月記載 (①のみ・翌月②に存在 → ②要修正)",
                r.num("A翌月記載件数") + "件 / " + Fmt.n0(r.dbl("A翌月記載額")) + "円");
        rowIndex = kv(ws, rowIndex, "①のみ / ②のみ (金額あり)", r.num("A①のみ") + " / " + r.num("A②のみ"));
        rowIndex++;

        rowIndex = section(ws, rowIndex, "検証B　② vs ③ (依頼NO突合)");
        rowIndex = kv(ws, rowIndex, "共通依頼NO", r.num("B共通") + " 件");
        rowIndex = kv(ws, rowIndex, "　不一致(要確認)", String.valueOf(r.num("B不一致")));
        rowIndex = kv(ws, rowIndex, "月ずれ解消 (複数月累計一致・対象外)", r.num("B月ずれ解消") + "件");
        rowIndex = kv(ws, rowIndex, "枝番統合一致 (③枝番を親に合算して一致・対象外)", r.num("B枝番統合") + "件");
        rowIndex = kv(ws, rowIndex, "②のみ / ③のみ (金額あり)", r.num("B②のみ") + " / " + r.num("B③のみ"));
        rowIndex++;

        MatomeCheckResult d = r.checkD();
        if (d != null && !d.isSkipped()) {
            rowIndex = section(ws, rowIndex, "検証D　②「" + FactoryProfile.MATOME_SHEET + "」の内部整合性");
            rowIndex = kv(ws, rowIndex, "要修正 (金額差・参照ずれ・未取込)", d.errorCount() + " 件");
            rowIndex = kv(ws, rowIndex, "注意 (式異常・直接入力)", d.noticeCount() + " 件");
            for (MatomeCheckResult.SheetTotal t : d.totals()) {
                rowIndex = kv(ws, rowIndex, "　" + t.sheet() + " まとめ計 / 元シート計 / 差",
                        Fmt.n0(t.matomeSum()) + " / " + Fmt.n0(t.srcSum()) + " / " + Fmt.s0(t.diff()));
            }
            rowIndex++;
        }

        rowIndex = section(ws, rowIndex, "報告用内訳 (① = ②実売上 + 前月調整 + 翌月記載 + 当月差異 + 未解明)");
        rowIndex = kv(ws, rowIndex, "① 東レお支払データ", r.num("①総額"));
        rowIndex = kv(ws, rowIndex, "② 当月実売上 (契約NO合計)", r.num("報告②実売上"));
        rowIndex = kv(ws, rowIndex, "前月以前調整分", Math.round(r.dbl("A前月調整額")));
        rowIndex = kv(ws, rowIndex, "翌月記載 (②の記載月誤り)", Math.round(r.dbl("A翌月記載額")));
        rowIndex = kv(ws, rowIndex, "当月差異 (金額不一致の差計)", r.num("報告当月差異"));
        rowIndex = kv(ws, rowIndex, "未解明 (①のみ)", Math.round(r.dbl("報告①のみ額")));
        rowIndex = kv(ws, rowIndex, "未解明 (②のみ)", -Math.round(r.dbl("報告②のみ額")));
        rowIndex = kv(ws, rowIndex, "報告する過不足 (当月)", r.num("報告する過不足"));
        rowIndex = kv(ws, rowIndex, "検算残差 (0なら分解が完全)", r.num("報告残差"));
        rowIndex++;

        if (!r.warnings().isEmpty()) {
            rowIndex = section(ws, rowIndex, "【警告】データ品質の注意 (" + r.warnings().size() + "件)");
            CellStyle warnStyle = style("warn", "9C0006", "FDE9E9", false, 10, false, null);
            for (String warning : r.warnings()) {
                put(ws, rowIndex, 1, "⚠ " + warning, warnStyle);
                merge(ws, rowIndex, 1, 7);
                rowIndex++;
            }
            rowIndex++;
        }

        rowIndex = section(ws, rowIndex, "使用ファイル");
        rowIndex = kv(ws, rowIndex, "① 東レCSV", r.str("①ファイル") + " (入庫場所 " + r.str("入庫場所") + ")");
        rowIndex = kv(ws, rowIndex, "② " + r.str("②名称"), r.str("②ファイル"));
        rowIndex = kv(ws, rowIndex, "　過去月明細", r.str("前月ファイル"));
        rowIndex = kv(ws, rowIndex, "　翌月明細", r.str("翌月ファイル"));
        rowIndex = kv(ws, rowIndex, "③ アラジン", r.str("③ファイル") + " (対象年月: " + r.str("対象年月") + ")");
        rowIndex = kv(ws, rowIndex, "　他月照会", r.str("③他月照会"));
        rowIndex = kv(ws, rowIndex, "①の対象外入庫場所",
                r.str("対象外入庫場所").isEmpty() ? "なし" : r.str("対象外入庫場所"));
        rowIndex = kv(ws, rowIndex, "①の取消行(H列「-」)", r.str("①取消行"));
        kv(ws, rowIndex, "金額0円の無効行(出力対象外)", r.str("0円除外"));
    }

    private int section(XSSFSheet ws, int rowIndex, String title) {
        put(ws, rowIndex, 1, title, style("section", ACCENT, null, true, 12, false, BorderStyle.MEDIUM));
        merge(ws, rowIndex, 1, 7);
        return rowIndex + 1;
    }

    private int kv(XSSFSheet ws, int rowIndex, String label, Object value) {
        put(ws, rowIndex, 1, label, style("kvLabel", null, null, false, 10, false, null));
        Cell cell = row(ws, rowIndex).createCell(2);
        if (value instanceof Number n) {
            cell.setCellValue(n.doubleValue());
            cell.setCellStyle(numberStyle(null, false, true));
        } else {
            cell.setCellValue(String.valueOf(value));
            cell.setCellStyle(style("kvValue", null, null, false, 10, false, null));
        }
        merge(ws, rowIndex, 2, 7);
        return rowIndex + 1;
    }

    // ---------------------------------------------------------------- 検証A/B/D

    private void writeSheetA(VerifyResult r) {
        List<String> headers = List.of("契約NO", "依頼NO", "①東レ金額", "②金額", "差額(①-②)",
                "報告計上額", "判定", "備考");
        XSSFSheet ws = createDetailSheet(SHEET_A, headers, "C00000");
        int rowIndex = 1;
        for (RecordA rec : r.recordsA()) {
            Row row = ws.createRow(rowIndex++);
            String fill = Judge.fillColor(rec.judge());
            text(row, 0, rec.keiyaku(), fill);
            text(row, 1, rec.iraiNo(), fill);
            number(row, 2, rec.amount1(), fill, false);
            number(row, 3, rec.amount2(), fill, false);
            number(row, 4, rec.diff(), fill, Judge.MISMATCH.equals(rec.judge()));
            number(row, 5, rec.reportAmount(), fill, true);
            judge(row, 6, rec.judge(), fill);
            note(row, 7, rec.note(), fill);
        }
        finishDetailSheet(ws, headers.size(), rowIndex);
    }

    private void writeSheetB(VerifyResult r) {
        List<String> headers = List.of("依頼NO", "②金額", "③アラジン金額", "差額(②-③)", "判定", "備考");
        XSSFSheet ws = createDetailSheet(SHEET_B, headers, "ED7D31");
        int rowIndex = 1;
        for (RecordB rec : r.recordsB()) {
            Row row = ws.createRow(rowIndex++);
            String fill = Judge.fillColor(rec.judge());
            text(row, 0, rec.irai(), fill);
            number(row, 1, rec.amount2(), fill, false);
            number(row, 2, rec.amount3(), fill, false);
            number(row, 3, rec.diff(), fill, Judge.MISMATCH.equals(rec.judge()));
            judge(row, 4, rec.judge(), fill);
            note(row, 5, rec.note(), fill);
        }
        finishDetailSheet(ws, headers.size(), rowIndex);
    }

    private void writeSheetD(VerifyResult r) {
        MatomeCheckResult d = r.checkD();
        if (d == null) {
            return;
        }
        List<String> headers = List.of("場所", "依頼NO", "契約NO", "まとめAA", "元シートAA",
                "差額(まとめ-元)", "判定", "内容・対処");
        XSSFSheet ws = createDetailSheet(SHEET_D, headers, "7030A0");
        int rowIndex = 1;
        if (d.isSkipped()) {
            Row row = ws.createRow(rowIndex++);
            text(row, 0, "(スキップ)", null);
            judge(row, 6, Judge.BAD_FORMULA, Judge.fillColor(Judge.BAD_FORMULA));
            note(row, 7, d.skipped(), null);
        } else if (d.rows().isEmpty()) {
            Row row = ws.createRow(rowIndex++);
            text(row, 0, "問題なし", null);
            note(row, 7, "まとめの参照式・元シートの式・行別金額がすべて整合", null);
        } else {
            List<MatomeCheckResult.MatomeRow> sorted = new ArrayList<>(d.rows());
            sorted.sort((x, y) -> Integer.compare(Judge.rank(x.judge()), Judge.rank(y.judge())));
            for (MatomeCheckResult.MatomeRow mr : sorted) {
                Row row = ws.createRow(rowIndex++);
                String fill = Judge.fillColor(mr.judge());
                text(row, 0, mr.place(), fill);
                text(row, 1, mr.irai(), fill);
                text(row, 2, mr.keiyaku(), fill);
                number(row, 3, mr.matomeAa(), fill, false);
                number(row, 4, mr.srcAa(), fill, false);
                number(row, 5, mr.diff(), fill, true);
                judge(row, 6, mr.judge(), fill);
                note(row, 7, mr.detail(), fill);
            }
        }
        finishDetailSheet(ws, headers.size(), rowIndex);
    }

    private void writeSheetC(VerifyResult r) {
        CheckCResult c = r.checkC();
        if (c == null) {
            return;
        }
        List<String> headers = List.of("項目", "月次処理ファイル", "本検証", "差額", "判定", "備考");
        XSSFSheet ws = createDetailSheet("検証C_月次処理", headers, "1F4E79");
        int rowIndex = 1;
        if (c.isSkipped()) {
            Row row = ws.createRow(rowIndex++);
            text(row, 0, "(スキップ)", null);
            note(row, 5, c.skipped(), null);
        } else {
            for (CheckCResult.Row cr : c.rows()) {
                Row row = ws.createRow(rowIndex++);
                String fill = CheckCResult.MATCH.equals(cr.judge())
                        ? "C6EFCE"
                        : CheckCResult.EXPLAINABLE.equals(cr.judge())
                                ? "DDEBF7"
                                : CheckCResult.UNREADABLE.equals(cr.judge())
                                        ? "F8CBAD"
                                        : CheckCResult.SHEET_INTERNAL.equals(cr.judge())
                                                ? "FFF2CC"
                                                : "FFC7CE";
                text(row, 0, cr.item(), fill);
                number(row, 1, cr.monthlyValue(), fill, false);
                number(row, 2, cr.ours(), fill, false);
                number(row, 3, cr.diff(), fill, true);
                judge(row, 4, cr.judge(), fill);
                note(row, 5, cr.note(), fill);
            }
        }
        finishDetailSheet(ws, headers.size(), rowIndex);
    }

    private void writeRawSheets(VerifyResult r) {
        SourceRawSheets.append(wb, r, this::rawCellStyle);
    }

    CellStyle rawCellStyle(boolean highlight) {
        return style(highlight ? "rawHl" : "raw", null, highlight ? "FFF2CC" : null, false, 9, false, BorderStyle.THIN);
    }

    private XSSFSheet createDetailSheet(String title, List<String> headers, String tabColor) {
        XSSFSheet ws = wb.createSheet(title);
        ws.setTabColor(new XSSFColor(rgb(tabColor), null));
        CellStyle headerStyle = style("header", "FFFFFF", HEADER_FILL, true, 10.5, true, BorderStyle.THIN);
        Row row = ws.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(headerStyle);
        }
        row.setHeightInPoints(28f);
        return ws;
    }

    private void finishDetailSheet(XSSFSheet ws, int columns, int rowCount) {
        for (int i = 0; i < columns; i++) {
            int width = i == 0 ? 14 : (i == columns - 1 ? 70 : (i == columns - 2 ? 12 : 16));
            ws.setColumnWidth(i, width * 256);
        }
        ws.createFreezePane(0, 1);
        if (rowCount > 1) {
            ws.setAutoFilter(new CellRangeAddress(0, rowCount - 1, 0, columns - 1));
        }
        ws.setRepeatingRows(CellRangeAddress.valueOf("$1:$1"));
        ws.getPrintSetup().setLandscape(true);
        ws.setFitToPage(true);
        ws.getPrintSetup().setFitWidth((short) 1);
        ws.getPrintSetup().setFitHeight((short) 0);
    }

    // ---------------------------------------------------------------- ガイド・メール

    private void writeGuide(VerifyResult r) {
        XSSFSheet ws = wb.createSheet(SHEET_GUIDE);
        ws.setDisplayGridlines(false);
        ws.setTabColor(new XSSFColor(rgb("808080"), null));
        ws.setColumnWidth(0, 2 * 256);
        ws.setColumnWidth(1, 24 * 256);
        ws.setColumnWidth(2, 90 * 256);

        put(ws, 0, 1, "レポートの見方【" + r.profile().label() + "】",
                style("title", ACCENT, null, true, 15, false, null));
        int rowIndex = 2;
        String[][] items = {
                {"データの突合構造", "① 東レ検収CSV ─ 検証A(契約NO) ─ ② " + r.str("②名称")
                        + " ─ 検証B(依頼NO) ─ ③ アラジン"},
                {"検証Aのキー", r.profile().keyA()},
                {"検証Bのキー", r.profile().keyB()},
                {"不一致 (赤)", "両方に存在するが金額が異なる。最優先で原因確認。差額計が東レへ報告する「当月差異」"},
                {"翌月記載 (橙)", "①当月検収なのに②は翌月に記載。②2ファイルを修正(当月へ追記・翌月から削除)して再実行"},
                {"①②③のみ (黄)", "片側にしか存在しない。記載漏れ・計上月ずれを調査"},
                {"前月調整 (青)", "①のみだが過去月の②に存在。判明済みで参考確認のみ"},
                {"月ずれ解消・枝番統合一致 (緑)", "複数月累計または③の枝番合算で一致。対応不要"},
                {"形式不正 (灰)", "②のC列が契約NO形式でないのに金額がある行。記入漏れ疑い"},
                {"報告する過不足", "当月差異 + 翌月記載 + ①のみ − ②のみ。検証Aの「報告計上額」列の合計と一致する"},
                {"検算残差", "0であれば報告用内訳の内部整合が取れている"},
        };
        CellStyle labelStyle = style("guideLabel", null, null, true, 10, false, BorderStyle.THIN);
        CellStyle textStyle = style("guideText", null, null, false, 10, false, BorderStyle.THIN);
        for (String[] item : items) {
            put(ws, rowIndex, 1, item[0], labelStyle);
            put(ws, rowIndex, 2, item[1], textStyle);
            rowIndex++;
        }
    }

    private void writeMail(MailSnapshot kokubu, MailSnapshot konan) {
        XSSFSheet ws = wb.createSheet(SHEET_MAIL);
        ws.setDisplayGridlines(false);
        ws.setTabColor(new XSSFColor(rgb("70AD47"), null));
        ws.setColumnWidth(0, 100 * 256);

        List<String> lines = UnifiedMailBuilder.buildLines(kokubu, konan);
        int noteLines = UnifiedMailBuilder.noteLineCount(lines);
        CellStyle noteStyle = style("mailNote", "C00000", null, false, 10.5, false, null);
        CellStyle markerStyle = style("mailMarker", "375623", "C6EFCE", true, 10.5, false, null);
        CellStyle bodyStyle = style("mailBody", null, null, false, 10.5, false, null);

        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i);
            CellStyle style = i < noteLines - 1 ? noteStyle : (i == noteLines - 1 ? markerStyle : bodyStyle);
            put(ws, i, 0, line, style);
        }
    }

    // ---------------------------------------------------------------- セル書式

    private void text(Row row, int col, String value, String fill) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value == null ? "" : value);
        cell.setCellStyle(style("cell" + fill, null, fill, false, 10.5, false, BorderStyle.THIN));
    }

    private void number(Row row, int col, Double value, String fill, boolean bold) {
        Cell cell = row.createCell(col);
        if (value == null) {
            cell.setCellValue("");
            cell.setCellStyle(style("cell" + fill, null, fill, false, 10.5, false, BorderStyle.THIN));
            return;
        }
        double rounded = Fmt.round2(value);
        cell.setCellValue(rounded);
        cell.setCellStyle(numberStyle(fill, bold, rounded == Math.rint(rounded)));
    }

    private void judge(Row row, int col, String value, String fill) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value == null ? "" : value);
        cell.setCellStyle(style("judge" + value, Judge.fontColor(value), fill, true, 10.5, true, BorderStyle.THIN));
    }

    private void note(Row row, int col, String value, String fill) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value == null ? "" : value);
        cell.setCellStyle(style("note" + fill, null, fill, false, 10.5, false, BorderStyle.THIN, true));
    }

    private CellStyle numberStyle(String fill, boolean bold, boolean integral) {
        String key = "num/" + fill + "/" + bold + "/" + integral;
        return styles.computeIfAbsent(key, k -> {
            XSSFCellStyle cs = wb.createCellStyle();
            cs.setFont(font(bold, null, 10.5, true));
            cs.setDataFormat(integral ? fmtInt : fmtDec);
            cs.setAlignment(HorizontalAlignment.RIGHT);
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            applyFill(cs, fill);
            applyBorder(cs, BorderStyle.THIN);
            return cs;
        });
    }

    private CellStyle style(String key, String fontColor, String fill, boolean bold,
                            double size, boolean center, BorderStyle border) {
        return style(key, fontColor, fill, bold, size, center, border, false);
    }

    private CellStyle style(String key, String fontColor, String fill, boolean bold,
                            double size, boolean center, BorderStyle border, boolean wrap) {
        String cacheKey = String.join("/", key, String.valueOf(fontColor), String.valueOf(fill),
                String.valueOf(bold), String.valueOf(size), String.valueOf(center),
                String.valueOf(border), String.valueOf(wrap));
        return styles.computeIfAbsent(cacheKey, k -> {
            XSSFCellStyle cs = wb.createCellStyle();
            cs.setFont(font(bold, fontColor, size, false));
            if (center) {
                cs.setAlignment(HorizontalAlignment.CENTER);
            }
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            cs.setWrapText(wrap);
            applyFill(cs, fill);
            applyBorder(cs, border);
            return cs;
        });
    }

    private void applyFill(XSSFCellStyle cs, String fill) {
        if (fill == null || "null".equals(fill)) {
            return;
        }
        cs.setFillForegroundColor(new XSSFColor(rgb(fill), null));
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private void applyBorder(XSSFCellStyle cs, BorderStyle border) {
        if (border == null) {
            return;
        }
        cs.setBorderTop(border);
        cs.setBorderBottom(border);
        cs.setBorderLeft(border);
        cs.setBorderRight(border);
    }

    private XSSFFont font(boolean bold, String colorHex, double size, boolean numeric) {
        String key = bold + "/" + colorHex + "/" + size + "/" + numeric;
        return fonts.computeIfAbsent(key, k -> {
            XSSFFont f = wb.createFont();
            f.setFontName(numeric ? numFont : baseFont);
            f.setFontHeight(size);
            f.setBold(bold);
            if (colorHex != null) {
                f.setColor(new XSSFColor(rgb(colorHex), null));
            }
            return f;
        });
    }

    private void put(XSSFSheet ws, int rowIndex, int col, String value, CellStyle style) {
        Cell cell = row(ws, rowIndex).createCell(col);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static Row row(XSSFSheet ws, int rowIndex) {
        Row row = ws.getRow(rowIndex);
        return row != null ? row : ws.createRow(rowIndex);
    }

    private static void merge(XSSFSheet ws, int rowIndex, int firstCol, int lastCol) {
        if (lastCol > firstCol) {
            ws.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, firstCol, lastCol));
        }
    }

    private static String[] palette(String state) {
        return switch (state) {
            case "bad" -> new String[] {"FDE9E9", "9C0006"};
            case "ok" -> new String[] {"C6EFCE", "375623"};
            case "known" -> new String[] {"E2EFDA", "375623"};
            default -> new String[] {"DDEBF7", ACCENT};
        };
    }

    private static byte[] rgb(String hex) {
        return new byte[] {
                (byte) Integer.parseInt(hex.substring(0, 2), 16),
                (byte) Integer.parseInt(hex.substring(2, 4), 16),
                (byte) Integer.parseInt(hex.substring(4, 6), 16)};
    }

    /** インストール済みフォントから最初に見つかったものを使う。 */
    private static String pickFont(String... candidates) {
        Set<String> available = availableFonts();
        for (String name : candidates) {
            if (available.isEmpty() || available.contains(name)) {
                return name;
            }
        }
        return candidates[candidates.length - 1];
    }

    private static Set<String> availableFonts() {
        try {
            return new HashSet<>(Arrays.asList(GraphicsEnvironment.getLocalGraphicsEnvironment()
                    .getAvailableFontFamilyNames(Locale.JAPAN)));
        } catch (Throwable t) {
            return Set.of();
        }
    }
}
