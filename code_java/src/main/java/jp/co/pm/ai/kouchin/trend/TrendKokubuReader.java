package jp.co.pm.ai.kouchin.trend;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.kouchin.verify.ExcelValues;
import jp.co.pm.ai.kouchin.verify.FactoryProfile;
import jp.co.pm.ai.kouchin.verify.VerifyException;
import jp.co.pm.ai.kouchin.verify.YearMonthKey;

/**
 * 国分② 後加工工賃明細（東レまとめ＋元4シート）の種類別・シート区分別集計。
 */
public final class TrendKokubuReader {

    private static final int COL_E = 4;
    private static final int COL_G = 6;
    private static final int COL_H = 7;
    private static final int COL_X = 23;
    private static final int COL_AA = 26;
    private static final int COL_AB = 27;
    private static final int COL_AT = 45;
    private static final int KIND_COUNT = COL_X - COL_H + 1;

    private TrendKokubuReader() {}

    public static TrendMonthData read(Path path) {
        List<String> warnings = new ArrayList<>();
        Map<String, TrendMetric> kinds = new LinkedHashMap<>();
        Map<String, TrendMetric> sheets = new LinkedHashMap<>();
        for (String s : FactoryProfile.MATOME_SRC_SHEETS) {
            sheets.put(s, new TrendMetric(0, 0));
        }
        YearMonthKey ym = YearMonthKey.parseGatsudo(path.getFileName().toString()).orElse(null);
        try (Workbook wb = ExcelValues.open(path)) {
            if (wb.getSheet(FactoryProfile.MATOME_SHEET) == null) {
                warnings.add(path.getFileName() + ": シート「" + FactoryProfile.MATOME_SHEET + "」がありません");
                return new TrendMonthData(ym, path, Map.of(), sheets, warnings);
            }
            List<List<Object>> matome = ExcelValues.readSheet(wb, FactoryProfile.MATOME_SHEET);
            Integer hdrI = findKindHeaderRow(matome);
            List<String> names;
            int dataStart;
            if (hdrI == null) {
                warnings.add(path.getFileName() + ": 東レまとめに工程名ヘッダーが見つかりません");
                names = List.of();
                dataStart = 0;
            } else {
                names = processNames(matome.get(hdrI));
                dataStart = hdrI + 1;
            }
            for (int r = dataStart; r < matome.size(); r++) {
                List<Object> row = matome.get(r);
                if (row == null || !isDataRow(row)) {
                    continue;
                }
                for (int i = 0; i < names.size(); i++) {
                    String name = names.get(i);
                    if (name == null || name.isEmpty()) {
                        continue;
                    }
                    add(kinds, name, ExcelValues.at(matome, r, COL_AB + i), ExcelValues.at(matome, r, COL_AT + i));
                }
            }
            for (String sname : FactoryProfile.MATOME_SRC_SHEETS) {
                if (wb.getSheet(sname) == null) {
                    warnings.add(path.getFileName() + ": シート「" + sname + "」がありません");
                    continue;
                }
                List<List<Object>> rows = ExcelValues.readSheet(wb, sname);
                for (List<Object> row : rows) {
                    if (row == null || !isDataRow(row)) {
                        continue;
                    }
                    add(sheets, sname, at(row, COL_AA), at(row, COL_E));
                }
            }
        } catch (IOException e) {
            throw new VerifyException("国分明細を閉じられませんでした: " + path + " (" + e.getMessage() + ")", e);
        }
        return new TrendMonthData(ym, path, kinds, sheets, warnings);
    }

    private static Integer findKindHeaderRow(List<List<Object>> rows) {
        int limit = Math.min(40, rows.size());
        for (int i = 0; i < limit; i++) {
            String a = TrendText.disp(ExcelValues.at(rows, i, 0));
            if (a.contains("加工依頼")) {
                if (i + 1 < rows.size() && processNames(rows.get(i + 1)).stream().anyMatch(n -> n != null)) {
                    return i + 1;
                }
            }
        }
        for (int i = 0; i < limit; i++) {
            if ("合計".equals(TrendText.disp(ExcelValues.at(rows, i, COL_G)))
                    && processNames(rows.get(i)).stream().anyMatch(n -> n != null)) {
                return i;
            }
        }
        return null;
    }

    private static List<String> processNames(List<Object> headerRow) {
        List<String> names = new ArrayList<>(KIND_COUNT);
        for (int i = 0; i < KIND_COUNT; i++) {
            Object v = headerRow != null && COL_H + i < headerRow.size() ? headerRow.get(COL_H + i) : null;
            String s = TrendText.disp(v);
            names.add(s.isEmpty() || "合計".equals(s) ? null : s);
        }
        return names;
    }

    private static boolean isDataRow(List<Object> row) {
        Object a = at(row, 0);
        Object b = at(row, 1);
        Object c = at(row, 2);
        String cS = TrendText.disp(c);
        if (TrendText.TOTAL_LABELS.contains(cS)) {
            return false;
        }
        String aS = TrendText.disp(a);
        if (!aS.isEmpty() && !aS.equals("0") && !aS.equals("0.0") && !isLooseNumber(aS)) {
            return true;
        }
        if (b != null && !TrendText.disp(b).isEmpty() && !TrendText.disp(b).equals("0") && !TrendText.disp(b).equals("0.0")) {
            if (!cS.isEmpty() && !TrendText.TOTAL_LABELS.contains(cS) && !cS.equals("0") && !cS.equals("0.0")) {
                return true;
            }
        }
        if (!cS.isEmpty() && !TrendText.TOTAL_LABELS.contains(cS) && !cS.equals("0") && !cS.equals("0.0")) {
            if (!aS.isEmpty()
                    || (b != null
                            && !TrendText.disp(b).isEmpty()
                            && !TrendText.disp(b).equals("0")
                            && !TrendText.disp(b).equals("0.0"))) {
                return true;
            }
        }
        return false;
    }

    private static boolean isLooseNumber(String s) {
        int dot = s.indexOf('.');
        if (dot < 0) {
            return s.chars().allMatch(Character::isDigit);
        }
        String once = s.substring(0, dot) + s.substring(dot + 1);
        return !once.isEmpty() && once.chars().allMatch(Character::isDigit);
    }

    private static Object at(List<Object> row, int idx) {
        return row != null && idx < row.size() ? row.get(idx) : null;
    }

    private static void add(Map<String, TrendMetric> map, String name, Object wage, Object qty) {
        TrendMetric cur = map.getOrDefault(name, new TrendMetric(0, 0));
        map.put(name, new TrendMetric(cur.wage() + TrendText.num(wage), cur.qty() + TrendText.num(qty)));
    }
}
