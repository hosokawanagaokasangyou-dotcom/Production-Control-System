package jp.co.pm.ai.kouchin.verify;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * ③ アラジン「依頼NO別問合せ*.xlsx」の読み取り。
 *
 * <ul>
 *   <li>「依頼NO」と「項目」を含む行がヘッダー。</li>
 *   <li>項目=「加工金額」の行の「--合計--」列が依頼NO別加工金額。</li>
 *   <li>得意先コードを指定するとその行だけを集計する（湖南は 049006 東ﾚ自材部のみ）。</li>
 *   <li>倉庫名キーワードを指定するとその工場の倉庫だけを集計する（湖南③は国分倉庫が混在するため）。</li>
 *   <li>TPI（得意先 049052 / 依頼NO先頭 TPI）と自社加工（依頼NO先頭 2）は常に対象外。</li>
 * </ul>
 */
public final class AladdinReader {

    private static final int HEADER_SEARCH_ROWS = 10;

    private AladdinReader() {
    }

    /**
     * @param path     依頼NO別問合せ*.xlsx
     * @param customer 得意先コード（絞り込み不要なら null）
     */
    public static AladdinData read(Path path, String customer) {
        return read(path, customer, null);
    }

    /**
     * @param warehouseContains 倉庫名にこの文字列を含む行だけ集計（湖南/国分。列が無ければ無視）
     */
    public static AladdinData read(Path path, String customer, String warehouseContains) {
        try (Workbook wb = ExcelValues.open(path)) {
            return read(ExcelValues.readFirstSheet(wb), customer, warehouseContains, path.toString());
        } catch (IOException e) {
            throw new VerifyException("③を閉じられませんでした: " + path + " (" + e.getMessage() + ")", e);
        }
    }

    /** 読み込み済みのシート値から集計する（テスト用に公開）。 */
    public static AladdinData read(List<List<Object>> rows, String customer, String source) {
        return read(rows, customer, null, source);
    }

    public static AladdinData read(List<List<Object>> rows, String customer, String warehouseContains, String source) {
        String taisho = "";
        int headerIndex = -1;
        for (int i = 0; i < Math.min(HEADER_SEARCH_ROWS, rows.size()); i++) {
            List<Object> row = rows.get(i);
            StringBuilder joined = new StringBuilder();
            for (Object v : row) {
                joined.append(Norm.text(v));
            }
            String line = joined.toString();
            if (line.contains("対象年月")) {
                int colon = Math.max(line.lastIndexOf(':'), line.lastIndexOf('：'));
                taisho = (colon >= 0 ? line.substring(colon + 1) : line).trim();
            }
            boolean hasIrai = false;
            boolean hasItem = false;
            for (Object v : row) {
                String n = Norm.norm(v);
                hasIrai |= "依頼NO".equals(n);
                hasItem |= "項目".equals(n);
            }
            if (hasIrai && hasItem) {
                headerIndex = i;
                break;
            }
        }
        if (headerIndex < 0) {
            throw new VerifyException("③のヘッダー行(依頼NO/項目)が見つかりません: " + source);
        }

        List<String> header = new ArrayList<>();
        for (Object v : rows.get(headerIndex)) {
            header.add(Norm.norm(v));
        }
        int iIrai = header.indexOf("依頼NO");
        int iItem = header.indexOf("項目");
        int iTotal = header.indexOf("--合計--");
        if (iIrai < 0 || iItem < 0 || iTotal < 0) {
            throw new VerifyException("③のヘッダーに想定列が見つかりません(依頼NO/項目/--合計--): " + source);
        }
        int iCustomer = header.indexOf("得意先");
        if (customer != null && !customer.isEmpty()) {
            if (iCustomer < 0) {
                throw new VerifyException("③のヘッダーに「得意先」列が無いため得意先 " + customer
                        + " で絞り込めません: " + source);
            }
        }
        int iWarehouseName = header.indexOf("倉庫名");
        String warehouseKey = Norm.norm(warehouseContains);

        String customerKey = Norm.norm(customer);
        Map<String, Double> result = new LinkedHashMap<>();
        for (int i = headerIndex + 1; i < rows.size(); i++) {
            List<Object> row = rows.get(i);
            if (row.size() <= iTotal) {
                continue;
            }
            if (!"加工金額".equals(Norm.norm(ExcelValues.at(rows, i, iItem)))) {
                continue;
            }
            String rowCustomer = iCustomer >= 0 ? Norm.norm(ExcelValues.at(rows, i, iCustomer)) : "";
            if (!customerKey.isEmpty() && !customerKey.equals(rowCustomer)) {
                continue;
            }
            if (!warehouseKey.isEmpty() && iWarehouseName >= 0) {
                String warehouseName = Norm.norm(ExcelValues.at(rows, i, iWarehouseName));
                if (!warehouseName.contains(warehouseKey)) {
                    continue;
                }
            }
            String irai = Norm.norm(ExcelValues.at(rows, i, iIrai));
            if (VerifyScope.outOfScope(irai, rowCustomer)) {
                continue;
            }
            Double v = Norm.number(ExcelValues.at(rows, i, iTotal));
            if (v != null) {
                result.merge(irai, v, Double::sum);
            }
        }
        if (result.isEmpty()) {
            throw new VerifyException("③から「加工金額」行を1件も取り込めませんでした"
                    + (customer != null && !customer.isEmpty() ? "(得意先 " + customer + " で絞り込み)" : "")
                    + "。フォーマット変更の可能性: " + source);
        }
        return new AladdinData(result, taisho, YearMonthKey.parseYearMonth(taisho).orElse(null));
    }

    /** シート内の「対象年月 : yyyy年mm月」だけを読む（ファイル選択用）。 */
    public static Optional<YearMonthKey> targetYm(Path path) {
        try (Workbook wb = ExcelValues.open(path)) {
            List<List<Object>> rows = ExcelValues.readFirstSheet(wb);
            for (int i = 0; i < Math.min(HEADER_SEARCH_ROWS, rows.size()); i++) {
                StringBuilder joined = new StringBuilder();
                for (Object v : rows.get(i)) {
                    joined.append(Norm.text(v));
                }
                if (joined.indexOf("対象年月") >= 0) {
                    Optional<YearMonthKey> ym = YearMonthKey.parseYearMonth(joined.toString());
                    if (ym.isPresent()) {
                        return ym;
                    }
                }
            }
            return Optional.empty();
        } catch (IOException | VerifyException e) {
            return Optional.empty();
        }
    }
}
