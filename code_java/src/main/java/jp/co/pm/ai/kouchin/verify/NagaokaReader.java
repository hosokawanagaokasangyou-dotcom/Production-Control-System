package jp.co.pm.ai.kouchin.verify;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

/**
 * ② 国分工場 長岡明細（後加工工賃明細*.xlsx）の「東レまとめ」シート読み取り。
 *
 * <ul>
 *   <li>データは6行目以降。</li>
 *   <li>依頼NO = A列 &amp; "-" &amp; B列（B列は数値セルなので {@code 72.0 → 72} に整数化）。</li>
 *   <li>契約NO = C列（ハイフン除去）。</li>
 *   <li>金額 = AA列（27列目・0始まりで26）。</li>
 * </ul>
 */
public final class NagaokaReader {

    /** AA列（0始まり） */
    private static final int COL_AA = 26;
    /** データ開始行（0始まり）= 6行目 */
    private static final int DATA_START_INDEX = 5;

    private NagaokaReader() {
    }

    public static MoneyMaps read(Path path) {
        try (Workbook wb = ExcelValues.open(path)) {
            List<List<Object>> rows = ExcelValues.readSheet(wb, FactoryProfile.MATOME_SHEET);
            return read(rows, path.getFileName().toString());
        } catch (IOException e) {
            throw new VerifyException("②を閉じられませんでした: " + path + " (" + e.getMessage() + ")", e);
        }
    }

    /** 読み込み済みのシート値から集計する（テスト用に公開）。 */
    public static MoneyMaps read(List<List<Object>> rows, String fileName) {
        if (rows.size() < 6
                || !"合計".equals(Norm.norm(ExcelValues.at(rows, 3, COL_AA)))
                || !Norm.norm(ExcelValues.at(rows, 2, 2)).contains("契約")) {
            throw new VerifyException("②「東レまとめ」の構成が想定と異なります"
                    + "(C3セル「契約No.」/AA4セル「合計」を確認): " + fileName);
        }

        MoneyMaps maps = new MoneyMaps();
        for (int i = DATA_START_INDEX; i < rows.size(); i++) {
            int rowNo = i + 1;
            String a = Norm.norm(ExcelValues.at(rows, i, 0));
            String b = Norm.norm(ExcelValues.at(rows, i, 1));
            Object aaCell = ExcelValues.at(rows, i, COL_AA);

            if (a.isEmpty() || b.isEmpty() || "0".equals(a) || "0".equals(b)) {
                continue;
            }
            Double aa = Norm.number(aaCell);
            if (aa == null) {
                if (!Norm.isBlank(aaCell)) {
                    maps.addWarning("②" + fileName + " " + rowNo + "行目: 依頼NO " + a + "-" + b
                            + " のAA列「" + Norm.text(aaCell) + "」が数値でないため除外");
                }
                continue;
            }

            String irai = a + "-" + b;
            if (VerifyScope.outOfScopeIrai(irai)) {
                continue;
            }
            maps.addIrai(irai, aa);
            String c = Norm.keiyaku(ExcelValues.at(rows, i, 2));
            if (Norm.isKeiyaku(c)) {
                maps.addKeiyaku(c, aa);
                maps.linkIrai(c, irai);
            } else if (!c.isEmpty() && !"0".equals(c)) {
                maps.addBadKeiyaku(String.valueOf(rowNo), c, aa);
            }
        }

        if (maps.byIrai().isEmpty()) {
            throw new VerifyException("②「東レまとめ」からデータ行を1件も取り込めませんでした: " + fileName);
        }
        return maps;
    }
}
