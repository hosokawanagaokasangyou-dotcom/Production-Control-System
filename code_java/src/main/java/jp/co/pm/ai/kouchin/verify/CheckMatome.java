package jp.co.pm.ai.kouchin.verify;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 検証D（国分のみ）: ②「東レまとめ」の内部整合性。
 *
 * <p>まとめは元4シート（東レT/東レV.C/東レY/東レW.E）の各行を {@code =東レT!A6} のように
 * 列ごとに参照し、G（単価合計）・Z・AA（加工賃）だけ自行で再計算する構造。
 * 次を検出する。</p>
 * <ol>
 *   <li>まとめ参照式の行・列・シートずれ</li>
 *   <li>元シートのデータ行の取込漏れ</li>
 *   <li>元シート・まとめの G/AA の式崩れ（空欄・直接入力・別式）</li>
 *   <li>行別の まとめAA と 元シートAA の金額差</li>
 * </ol>
 */
public final class CheckMatome {

    /** データ開始行（1始まり） */
    private static final int DATA_START = 6;
    /** 元シート末尾の合計行（C列ラベル） */
    private static final Set<String> TOTAL_LABELS = Set.of("加工量合計", "加工賃合計", "営業入庫量");
    /** {@code =東レT!A6} / {@code =+'東レV.C'!A6}（POI は先頭の '=' を含まない） */
    private static final Pattern REF_PATTERN =
            Pattern.compile("^\\+?'?(東レ[^'!]+)'?!\\$?([A-Z]{1,2})\\$?(\\d+)$");
    /** 式中のセル参照の行番号 */
    private static final Pattern CELL_ROW_PATTERN =
            Pattern.compile("(?<![A-Za-z!$])\\$?[A-Z]{1,2}\\$?(\\d+)");
    /** AA列（1始まり） */
    private static final int COL_AA = 27;
    /** G列（1始まり） */
    private static final int COL_G = 7;

    private CheckMatome() {
    }

    public static MatomeCheckResult check(Path path, double tol) {
        Workbook wb;
        try {
            wb = ExcelValues.open(path);
        } catch (VerifyException e) {
            return MatomeCheckResult.skipped("数式を読めませんでした: " + e.getMessage());
        }
        try {
            return check(wb, tol);
        } catch (RuntimeException e) {
            return MatomeCheckResult.skipped("検証中にエラーが発生しました: " + e);
        } finally {
            try {
                wb.close();
            } catch (IOException ignored) {
                // 読み取り専用のため無視
            }
        }
    }

    private static MatomeCheckResult check(Workbook wb, double tol) {
        Sheet matome = wb.getSheet(FactoryProfile.MATOME_SHEET);
        if (matome == null) {
            return MatomeCheckResult.skipped("シート構成が想定外です(必要: "
                    + FactoryProfile.MATOME_SHEET + ", " + String.join(", ", FactoryProfile.MATOME_SRC_SHEETS) + ")");
        }
        for (String s : FactoryProfile.MATOME_SRC_SHEETS) {
            if (wb.getSheet(s) == null) {
                return MatomeCheckResult.skipped("シート構成が想定外です(必要: "
                        + FactoryProfile.MATOME_SHEET + ", " + String.join(", ", FactoryProfile.MATOME_SRC_SHEETS) + ")");
            }
        }

        Map<String, List<List<Object>>> values = new LinkedHashMap<>();
        values.put(FactoryProfile.MATOME_SHEET, ExcelValues.readSheet(matome));
        for (String s : FactoryProfile.MATOME_SRC_SHEETS) {
            values.put(s, ExcelValues.readSheet(wb.getSheet(s)));
        }

        List<MatomeCheckResult.MatomeRow> out = new ArrayList<>();
        Map<Integer, String[]> mapping = new TreeMap<>();   // まとめ行 -> [元シート, 元行]

        List<String> srcCols = new ArrayList<>();
        for (int c = 1; c <= 24; c++) {
            String letter = CellReference.convertNumToColString(c - 1);
            if (!"G".equals(letter)) {
                srcCols.add(letter);
            }
        }

        int lastRow = matome.getLastRowNum() + 1;
        for (int r = DATA_START; r <= lastRow; r++) {
            Map<String, String[]> refs = new LinkedHashMap<>();
            for (String col : srcCols) {
                String f = formula(matome, r, col);
                if (f == null) {
                    continue;
                }
                Matcher m = REF_PATTERN.matcher(f.replace(" ", ""));
                if (m.matches()) {
                    refs.put(col, new String[] {m.group(1), m.group(2), m.group(3)});
                }
            }
            if (refs.isEmpty()) {
                continue;
            }
            String sheetName = mostCommon(refs, 0);
            String srcRowText = mostCommon(refs, 2);
            int srcRow = Integer.parseInt(srcRowText);
            mapping.put(r, new String[] {sheetName, srcRowText});

            List<List<Object>> srcValues = values.get(sheetName);
            String[] ident = ident(srcValues, srcRow);

            List<String> problems = new ArrayList<>();
            List<String> sheetMix = new ArrayList<>();
            List<String> rowMix = new ArrayList<>();
            List<String> colMix = new ArrayList<>();
            for (Map.Entry<String, String[]> e : refs.entrySet()) {
                String[] ref = e.getValue();
                if (!ref[0].equals(sheetName)) {
                    sheetMix.add(e.getKey() + "→" + ref[0]);
                }
                if (!ref[2].equals(srcRowText)) {
                    rowMix.add(e.getKey() + "→" + ref[0] + "!" + ref[1] + ref[2]);
                }
                if (!ref[1].equals(e.getKey())) {
                    colMix.add(e.getKey() + "→" + ref[1]);
                }
            }
            if (!sheetMix.isEmpty()) {
                problems.add("参照シートが混在: " + String.join(", ", sheetMix));
            }
            if (!rowMix.isEmpty()) {
                problems.add("参照行がずれ: " + String.join(", ", rowMix));
            }
            if (!colMix.isEmpty()) {
                problems.add("参照列が自列と不一致: " + String.join(", ", colMix));
            }
            if (!problems.isEmpty()) {
                out.add(new MatomeCheckResult.MatomeRow(
                        FactoryProfile.MATOME_SHEET + "!" + r, ident[0], ident[1], null, null, null,
                        Judge.REF_SHIFT, String.join(" / ", problems)));
            }

            if (srcValues == null) {
                continue;
            }
            for (String col : new String[] {"G", "Z", "AA"}) {
                String f = formula(matome, r, col);
                String expected = expected(col, r);
                Object raw = rawValue(matome, r, col);
                if (f != null) {
                    Set<Integer> otherRows = otherRows(f, r);
                    if (!otherRows.isEmpty()) {
                        out.add(new MatomeCheckResult.MatomeRow(
                                FactoryProfile.MATOME_SHEET + "!" + col + r, ident[0], ident[1], null, null, null,
                                Judge.REF_SHIFT, "自行式が他行を参照: =" + f + " (期待 =" + expected + ")"));
                    } else if (!f.replace(" ", "").equals(expected.replace(" ", ""))) {
                        out.add(new MatomeCheckResult.MatomeRow(
                                FactoryProfile.MATOME_SHEET + "!" + col + r, ident[0], ident[1], null, null, null,
                                Judge.BAD_FORMULA, "定型と異なる式: =" + f + " (期待 =" + expected + ")"));
                    }
                } else if (raw != null && !Norm.isBlank(raw)
                        && (!ident[0].isEmpty() || !ident[1].isEmpty())) {
                    out.add(new MatomeCheckResult.MatomeRow(
                            FactoryProfile.MATOME_SHEET + "!" + col + r, ident[0], ident[1], null, null, null,
                            Judge.HARDCODED, "式ではなく値 " + Norm.text(raw) + " が入力されている (期待 ="
                            + expected + ")。元シートの単価欄と食い違う恐れ"));
                }
            }
        }

        if (mapping.isEmpty()) {
            return MatomeCheckResult.skipped("「" + FactoryProfile.MATOME_SHEET
                    + "」に元シートへの参照式が無い(値貼り付け)ため元シートとの突合ができません");
        }

        Map<String, Set<Integer>> covered = new HashMap<>();
        for (Map.Entry<Integer, String[]> e : mapping.entrySet()) {
            covered.computeIfAbsent(e.getValue()[0], k -> new HashSet<>())
                    .add(Integer.parseInt(e.getValue()[1]));
        }

        List<MatomeCheckResult.SheetTotal> totals = new ArrayList<>();
        List<List<Object>> matomeValues = values.get(FactoryProfile.MATOME_SHEET);
        for (String s : FactoryProfile.MATOME_SRC_SHEETS) {
            Sheet src = wb.getSheet(s);
            List<List<Object>> srcValues = values.get(s);
            double sumMatome = 0.0;
            double sumSrc = 0.0;
            Set<Integer> dataRows = new LinkedHashSet<>();

            for (int i = DATA_START; i <= srcValues.size(); i++) {
                String[] ident = ident(srcValues, i);
                if (ident[0].isEmpty() && ident[1].isEmpty()) {
                    continue;
                }
                dataRows.add(i);
                for (String col : new String[] {"G", "AA"}) {
                    String f = formula(src, i, col);
                    String expected = expected(col, i);
                    Object raw = rawValue(src, i, col);
                    double aa = ExcelValues.num1(srcValues, i, COL_AA);
                    if (f == null && (raw == null || Norm.isBlank(raw))) {
                        out.add(new MatomeCheckResult.MatomeRow(s + "!" + col + i, ident[0], ident[1], null, aa, null,
                                Judge.BAD_FORMULA, col + "列が空欄 (期待 =" + expected + ")。加工賃が0になる"));
                    } else if (f != null) {
                        if (!f.replace(" ", "").equals(expected.replace(" ", ""))) {
                            out.add(new MatomeCheckResult.MatomeRow(s + "!" + col + i, ident[0], ident[1], null, aa, null,
                                    Judge.BAD_FORMULA, "定型と異なる式: =" + f + " (期待 =" + expected + ")"));
                        }
                    } else {
                        out.add(new MatomeCheckResult.MatomeRow(s + "!" + col + i, ident[0], ident[1], null, aa, null,
                                Judge.HARDCODED, "式ではなく値 " + Norm.text(raw) + " が入力されている (期待 ="
                                + expected + ")。まとめは単価欄(H〜X)の合計で再計算するため食い違う恐れ"));
                    }
                }
            }

            Set<Integer> coveredRows = covered.getOrDefault(s, Set.of());
            for (Integer i : dataRows) {
                if (!coveredRows.contains(i)) {
                    String[] ident = ident(srcValues, i);
                    double aa = ExcelValues.num1(srcValues, i, COL_AA);
                    out.add(new MatomeCheckResult.MatomeRow(s + "!" + i, ident[0], ident[1], null, aa, null,
                            Judge.NOT_MAPPED, "元シートのデータ行がまとめに参照されていない (AA " + Fmt.n0(aa) + "円)"));
                }
            }

            for (Map.Entry<Integer, String[]> e : mapping.entrySet()) {
                if (!e.getValue()[0].equals(s)) {
                    continue;
                }
                int srcRow = Integer.parseInt(e.getValue()[1]);
                if (!dataRows.contains(srcRow)) {
                    continue;
                }
                int r = e.getKey();
                double matomeAa = ExcelValues.num1(matomeValues, r, COL_AA);
                double srcAa = ExcelValues.num1(srcValues, srcRow, COL_AA);
                sumMatome += matomeAa;
                sumSrc += srcAa;
                if (Math.abs(matomeAa - srcAa) > tol) {
                    String[] ident = ident(srcValues, srcRow);
                    double gm = ExcelValues.num1(matomeValues, r, COL_G);
                    double gs = ExcelValues.num1(srcValues, srcRow, COL_G);
                    out.add(new MatomeCheckResult.MatomeRow(
                            FactoryProfile.MATOME_SHEET + "!" + r + " ↔ " + s + "!" + srcRow,
                            ident[0], ident[1], matomeAa, srcAa, matomeAa - srcAa, Judge.AMOUNT_DIFF,
                            "まとめAA " + Fmt.n0(matomeAa) + " ≠ 元シートAA " + Fmt.n0(srcAa)
                                    + " (単価合計G: まとめ " + trim(gm) + " / 元 " + trim(gs) + ")。"
                                    + "元シートの単価を丸めるか、まとめの直接入力を式に戻して両者を一致させる"));
                }
            }
            totals.add(new MatomeCheckResult.SheetTotal(s, sumMatome, sumSrc, sumMatome - sumSrc));
        }

        return new MatomeCheckResult(List.copyOf(out), List.copyOf(totals), mapping.size(), null);
    }

    /** 元シート行の (依頼NO, 契約NO)。両方空ならデータ行ではない。 */
    private static String[] ident(List<List<Object>> rows, int row1) {
        if (rows == null) {
            return new String[] {"", ""};
        }
        String a = Norm.norm(ExcelValues.cell1(rows, row1, 1));
        String b = Norm.norm(ExcelValues.cell1(rows, row1, 2));
        String c = Norm.norm(ExcelValues.cell1(rows, row1, 3));
        String irai = (!a.isEmpty() && !b.isEmpty()) ? a + "-" + b : "";
        String keiyaku = TOTAL_LABELS.contains(c) ? "" : c;
        return new String[] {irai, keiyaku};
    }

    /** 元シート・まとめ共通の定型式（POI に合わせて先頭の '=' は含まない）。 */
    static String expected(String col, int row) {
        return switch (col) {
            case "G" -> "SUM(H" + row + ":X" + row + ")";
            case "Z" -> "ROUNDUP($D" + row + "*G" + row + ",0)";
            case "AA" -> "ROUNDUP($E" + row + "*G" + row + ",0)";
            default -> "";
        };
    }

    private static Set<Integer> otherRows(String formula, int row) {
        Set<Integer> rows = new LinkedHashSet<>();
        Matcher m = CELL_ROW_PATTERN.matcher(formula);
        while (m.find()) {
            int n = Integer.parseInt(m.group(1));
            if (n != row) {
                rows.add(n);
            }
        }
        return rows;
    }

    private static String mostCommon(Map<String, String[]> refs, int index) {
        Map<String, Integer> counts = new LinkedHashMap<>();
        for (String[] ref : refs.values()) {
            counts.merge(ref[index], 1, Integer::sum);
        }
        String best = null;
        int bestCount = -1;
        for (Map.Entry<String, Integer> e : counts.entrySet()) {
            if (e.getValue() > bestCount) {
                best = e.getKey();
                bestCount = e.getValue();
            }
        }
        return best;
    }

    private static Cell cell(Sheet sheet, int row1, String col) {
        Row row = sheet.getRow(row1 - 1);
        if (row == null) {
            return null;
        }
        return row.getCell(CellReference.convertColStringToIndex(col));
    }

    /** 数式文字列（'=' なし）。数式でなければ null。 */
    private static String formula(Sheet sheet, int row1, String col) {
        Cell c = cell(sheet, row1, col);
        if (c == null || c.getCellType() != CellType.FORMULA) {
            return null;
        }
        try {
            return c.getCellFormula();
        } catch (RuntimeException e) {
            return null;
        }
    }

    /** 数式でないセルの生値。 */
    private static Object rawValue(Sheet sheet, int row1, String col) {
        Cell c = cell(sheet, row1, col);
        if (c == null || c.getCellType() == CellType.FORMULA) {
            return null;
        }
        return ExcelValues.cellValue(c);
    }

    private static String trim(double v) {
        double r = Math.round(v * 10000.0) / 10000.0;
        return r == Math.rint(r) ? Long.toString((long) r) : Double.toString(r);
    }
}
