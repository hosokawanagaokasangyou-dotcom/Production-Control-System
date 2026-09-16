package jp.co.pm.ai.kouchin.verify;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;

/**
 * ② 湖南工場 加工賃試算（*加工賃試算*.xlsm）の東レ3シート読み取り。
 *
 * <ul>
 *   <li>A列「加工依頼No.」を含む行がヘッダー。C列は「契約No.」。</li>
 *   <li>金額列はヘッダー行の「加工種類別」かつ「加工賃」を含む列。</li>
 *   <li>ヘッダー行の3行下が合計行、4行下からデータ行。合計行とデータ行合計を検算する。</li>
 * </ul>
 */
public final class ShisanReader {

    private static final double TOLERANCE = 0.5;

    private ShisanReader() {
    }

    public static MoneyMaps read(Path path) {
        try (Workbook wb = ExcelValues.open(path)) {
            MoneyMaps maps = new MoneyMaps();
            String fileName = path.getFileName().toString();
            int found = 0;
            for (String name : FactoryProfile.SHISAN_SHEETS) {
                if (wb.getSheet(name) == null) {
                    maps.addWarning("②" + fileName + ": シート「" + name + "」が無いためスキップしました");
                    continue;
                }
                found++;
                readSheet(ExcelValues.readSheet(wb, name), name, fileName, path, maps);
            }
            if (found == 0) {
                throw new VerifyException("②に東レシート("
                        + String.join("/", FactoryProfile.SHISAN_SHEETS) + ")が1つもありません: " + path);
            }
            if (maps.isEmpty()) {
                throw new VerifyException("②の東レシートからデータ行を1件も取り込めませんでした: " + path);
            }
            return maps;
        } catch (IOException e) {
            throw new VerifyException("②を閉じられませんでした: " + path + " (" + e.getMessage() + ")", e);
        }
    }

    private static void readSheet(List<List<Object>> rows, String sheetName, String fileName,
                                  Path path, MoneyMaps maps) {
        int header = -1;
        for (int i = 0; i < Math.min(20, rows.size()); i++) {
            Object first = ExcelValues.at(rows, i, 0);
            if (first != null && Norm.norm(first).contains("加工依頼")) {
                header = i;
                break;
            }
        }
        if (header < 0 || rows.get(header).size() < 3
                || !Norm.norm(ExcelValues.at(rows, header, 2)).contains("契約")) {
            throw new VerifyException("②「" + sheetName
                    + "」のヘッダー(A列「加工依頼No.」/C列「契約No.」)が見つかりません: " + path);
        }

        int amtCol = -1;
        List<Object> headerRow = rows.get(header);
        for (int j = 0; j < headerRow.size(); j++) {
            String v = Norm.norm(headerRow.get(j));
            if (v.contains("加工賃") && v.contains("加工種類別")) {
                amtCol = j;
                break;
            }
        }
        if (amtCol < 0) {
            throw new VerifyException("②「" + sheetName + "」に「加工種類別 加工賃」列が見つかりません: " + path);
        }

        Double sheetTotal = Norm.number(ExcelValues.at(rows, header + 3, amtCol));
        double sheetSum = 0.0;

        for (int i = header + 4; i < rows.size(); i++) {
            int rowNo = i + 1;
            String irai = Norm.norm(ExcelValues.at(rows, i, 0));
            String c = Norm.keiyaku(ExcelValues.at(rows, i, 2));
            Object amtCell = ExcelValues.at(rows, i, amtCol);

            if (irai.isEmpty() && c.isEmpty()) {
                continue;
            }
            Double amt = Norm.number(amtCell);
            if (amt == null) {
                if (!Norm.isBlank(amtCell)) {
                    maps.addWarning("②" + fileName + "「" + sheetName + "」" + rowNo + "行目: 依頼NO " + irai
                            + " 契約NO " + c + " の加工賃「" + Norm.text(amtCell) + "」が数値でないため除外");
                }
                continue;
            }
            sheetSum += amt;
            if (!irai.isEmpty()) {
                maps.addIrai(irai, amt);
            }
            if (Norm.isKeiyaku(c)) {
                maps.addKeiyaku(c, amt);
                if (!irai.isEmpty()) {
                    maps.linkIrai(c, irai);
                }
            } else if (!c.isEmpty() && !"0".equals(c)) {
                maps.addBadKeiyaku("「" + sheetName + "」" + rowNo, c, amt);
            } else if (c.isEmpty() && Math.abs(amt) > TOLERANCE) {
                maps.addBadKeiyaku("「" + sheetName + "」" + rowNo, "(契約NO空欄)", amt);
            }
        }

        if (sheetTotal != null && Math.abs(sheetTotal - sheetSum) > TOLERANCE) {
            maps.addWarning("②" + fileName + "「" + sheetName + "」: 合計行(" + Fmt.n0(sheetTotal)
                    + "円)とデータ行合計(" + Fmt.n0(sheetSum) + "円)が " + Fmt.s0(sheetSum - sheetTotal)
                    + "円 不一致。シート内の集計式を確認");
        }
    }

    /**
     * 湖南②の対象年月を東レシート先頭の「yyyy年m月度」から取得する（ファイル名に年が無いため）。
     */
    public static Optional<YearMonthKey> targetYm(Path path) {
        try (Workbook wb = ExcelValues.open(path)) {
            for (String name : FactoryProfile.SHISAN_SHEETS) {
                if (wb.getSheet(name) == null) {
                    continue;
                }
                List<List<Object>> rows = ExcelValues.readSheet(wb, name);
                for (int i = 0; i < Math.min(3, rows.size()); i++) {
                    for (Object v : rows.get(i)) {
                        if (v instanceof String s) {
                            Optional<YearMonthKey> ym = YearMonthKey.parseGatsudo(s);
                            if (ym.isPresent()) {
                                return ym;
                            }
                        }
                    }
                }
            }
            return Optional.empty();
        } catch (IOException | VerifyException e) {
            return Optional.empty();
        }
    }
}
