package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * 検証C（湖南）: 月次処理ファイルの東レ合計 vs 本検証の総額。
 */
public final class CheckVerifyC {

    public static final String ITEM_JISSEKI1 = "月次検証: 月次実績①(Excel)";
    public static final String ITEM_SHUKEI = "集計表: 東レ合計 加工金額";
    public static final String ITEM_JISSEKI2 = "月次検証: 月次実績表②(アラジン)";
    public static final String ITEM_URIAGE3 = "月次検証: 売上明細③";

    private static final Pattern YEAR_DIR = Pattern.compile("^\\d{4}年$");
    private static final Pattern MONTHLY_NAME = Pattern.compile("月次処理ファイル", Pattern.UNICODE_CHARACTER_CLASS);

    private CheckVerifyC() {}

    public static Path findMonthlyFile(Path folder, YearMonthKey targetYm) {
        if (folder == null || !Files.isDirectory(folder)) {
            return null;
        }
        List<Path> folders = new ArrayList<>();
        folders.add(folder);
        if (targetYm != null) {
            folders.add(folder.resolve(targetYm.year() + "年"));
        } else {
            try (DirectoryStream<Path> ds = Files.newDirectoryStream(folder)) {
                Path latestYear = null;
                for (Path d : ds) {
                    if (Files.isDirectory(d) && YEAR_DIR.matcher(Norm.norm(d.getFileName().toString())).matches()) {
                        if (latestYear == null
                                || d.getFileName().toString().compareTo(latestYear.getFileName().toString()) > 0) {
                            latestYear = d;
                        }
                    }
                }
                if (latestYear != null) {
                    folders.add(latestYear);
                }
            } catch (IOException ignored) {
            }
        }
        Path best = null;
        long bestMtime = Long.MIN_VALUE;
        for (Path dir : folders) {
            if (!Files.isDirectory(dir)) {
                continue;
            }
            try (DirectoryStream<Path> ds = Files.newDirectoryStream(dir)) {
                for (Path f : ds) {
                    String name = f.getFileName().toString();
                    if (name.startsWith("~$") || !MONTHLY_NAME.matcher(name).find()) {
                        continue;
                    }
                    String lower = name.toLowerCase();
                    if (!(lower.endsWith(".xlsx") || lower.endsWith(".xlsm") || lower.endsWith(".xls"))) {
                        continue;
                    }
                    if (targetYm != null) {
                        YearMonthKey ym = YearMonthKey.parseGatsudo(name)
                                .or(() -> YearMonthKey.parseYearMonth(name))
                                .orElse(null);
                        if (ym == null || !ym.equals(targetYm)) {
                            continue;
                        }
                    }
                    long mt = 0L;
                    try {
                        mt = Files.getLastModifiedTime(f).toMillis();
                    } catch (IOException ignored) {
                    }
                    if (best == null || mt > bestMtime) {
                        best = f;
                        bestMtime = mt;
                    }
                }
            } catch (IOException ignored) {
            }
        }
        return best;
    }

    public static Map<String, Double> readMonthlyFile(Path path) {
        Map<String, Double> res = new LinkedHashMap<>();
        Workbook wb;
        try {
            wb = ExcelValues.open(path);
        } catch (VerifyException e) {
            return res;
        }
        try {
            List<List<Object>> monthly = ExcelValues.readSheet(wb, "月次検証");
            if (!monthly.isEmpty()) {
                Integer hdr = null;
                for (int i = 0; i < Math.min(5, monthly.size()); i++) {
                    for (Object v : monthly.get(i)) {
                        if (Norm.norm(v).contains("月次実績")) {
                            hdr = i;
                            break;
                        }
                    }
                    if (hdr != null) {
                        break;
                    }
                }
                List<Object> toray = null;
                for (List<Object> r : monthly) {
                    if (!r.isEmpty() && "東レ".equals(Norm.norm(r.get(0)))) {
                        toray = r;
                        break;
                    }
                }
                if (hdr != null && toray != null) {
                    List<String> h = new ArrayList<>();
                    for (Object v : monthly.get(hdr)) {
                        h.add(Norm.norm(v));
                    }
                    res.put(ITEM_JISSEKI1, col(h, toray, "月次実績1"));
                    res.put(ITEM_JISSEKI2, col(h, toray, "月次実績表2"));
                    res.put(ITEM_URIAGE3, col(h, toray, "売上明細3"));
                }
            }
            List<List<Object>> shukei = ExcelValues.readSheet(wb, "集計表");
            if (!shukei.isEmpty()) {
                Integer hdr = null;
                for (int i = 0; i < Math.min(5, shukei.size()); i++) {
                    for (Object v : shukei.get(i)) {
                        if ("加工金額".equals(Norm.norm(v))) {
                            hdr = i;
                            break;
                        }
                    }
                    if (hdr != null) {
                        break;
                    }
                }
                List<Object> row = null;
                for (List<Object> r : shukei) {
                    if (!r.isEmpty() && "東レ合計".equals(Norm.norm(r.get(0)))) {
                        row = r;
                        break;
                    }
                }
                if (hdr != null && row != null) {
                    int j = -1;
                    List<Object> hrow = shukei.get(hdr);
                    for (int i = 0; i < hrow.size(); i++) {
                        if ("加工金額".equals(Norm.norm(hrow.get(i)))) {
                            j = i;
                            break;
                        }
                    }
                    if (j >= 0 && j < row.size()) {
                        Double v = Norm.number(row.get(j));
                        if (v != null) {
                            res.put(ITEM_SHUKEI, v);
                        }
                    }
                }
            }
        } finally {
            try {
                wb.close();
            } catch (IOException ignored) {
            }
        }
        return res;
    }

    public static CheckCResult evaluate(
            Path monthlyFile,
            Map<String, Double> monthly,
            double total2Irai,
            double total3,
            double total1,
            double explained,
            double tol) {
        if (monthlyFile == null) {
            return CheckCResult.skipped("月次処理ファイルがありません");
        }
        List<CheckCResult.Row> rows = new ArrayList<>();
        addCompare(rows, monthly, ITEM_JISSEKI1, total2Irai, "②加工賃試算 東レ3シート合計", tol);
        addCompare(rows, monthly, ITEM_SHUKEI, total2Irai, "②加工賃試算 東レ3シート合計", tol);
        addCompare(rows, monthly, ITEM_JISSEKI2, total3, "③アラジン 東レ(049006)合計", tol);

        Double mv = monthly.get(ITEM_URIAGE3);
        if (mv == null) {
            rows.add(new CheckCResult.Row(ITEM_URIAGE3, null, total1, null, CheckCResult.UNREADABLE, "①東レCSV総額 と比較できず"));
        } else {
            double d = total1 - mv;
            String judge;
            String note;
            if (Math.abs(d) <= tol) {
                judge = CheckCResult.MATCH;
                note = "vs ①東レCSV総額";
            } else if (Math.abs(d - explained) <= tol) {
                judge = CheckCResult.EXPLAINABLE;
                note = "vs ①東レCSV総額。差 " + Fmt.s0(d) + "円 = 報告用内訳 " + Fmt.s0(explained) + "円";
            } else {
                judge = CheckCResult.NEED_CHECK;
                note = "vs ①東レCSV総額。差 " + Fmt.s0(d) + "円 のうち報告用内訳で説明できるのは "
                        + Fmt.s0(explained) + "円 (未説明 " + Fmt.s0(d - explained) + "円)";
            }
            rows.add(new CheckCResult.Row(ITEM_URIAGE3, mv, total1, d, judge, note));
        }
        return new CheckCResult(monthlyFile, null, List.copyOf(rows));
    }

    private static void addCompare(
            List<CheckCResult.Row> rows,
            Map<String, Double> monthly,
            String key,
            double ours,
            String oursLabel,
            double tol) {
        Double mv = monthly.get(key);
        if (mv == null) {
            rows.add(new CheckCResult.Row(key, null, ours, null, CheckCResult.UNREADABLE, oursLabel + " と比較できず"));
            return;
        }
        double d = ours - mv;
        String judge = Math.abs(d) <= tol ? CheckCResult.MATCH : CheckCResult.MISMATCH;
        rows.add(new CheckCResult.Row(key, mv, ours, d, judge, "vs " + oursLabel));
    }

    private static Double col(List<String> h, List<Object> toray, String keyNormContains) {
        int j = -1;
        for (int i = 0; i < h.size(); i++) {
            if (h.get(i).contains(keyNormContains)) {
                j = i;
                break;
            }
        }
        if (j < 0 || j >= toray.size()) {
            return null;
        }
        return Norm.number(toray.get(j));
    }
}
