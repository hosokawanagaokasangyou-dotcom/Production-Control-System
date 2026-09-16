package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.CharsetDecoder;
import java.nio.charset.CodingErrorAction;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

/**
 * ① 東レ送付CSV（cp932）の読み取り。
 *
 * <ul>
 *   <li>A列（入庫場所）はブロック先頭行にのみ入り、以降の行へ引き継がれる（前方補完）。</li>
 *   <li>H列（8列目）が「-」の行は取消行。金額列（J列=10列目）は絶対値なので符号を反転して集計する。</li>
 *   <li>発注No.（C列）が {@code \d{3}-\d{3}[A-Z0-9]} 形式の行だけがデータ行。</li>
 *   <li>T行（ブロック小計）とデータ行合計を検算し、不一致を警告として返す。</li>
 * </ul>
 */
public final class TorayCsvReader {

    /** ①の発注No. 例: 191-352R */
    public static final Pattern KEIYAKU_CSV_PATTERN = Pattern.compile("^\\d{3}-\\d{3}[A-Z0-9]$");

    /** 金額列（0始まり）= J列 */
    private static final int COL_AMOUNT = 9;
    /** 符号列（0始まり）= H列 */
    private static final int COL_SIGN = 7;
    /** 発注No.列（0始まり）= C列 */
    private static final int COL_KEIYAKU = 2;
    /** 入庫月日列（0始まり）= D列 */
    private static final int COL_DATE = 3;
    /** 入庫場所列（0始まり）= A列 */
    private static final int COL_BASHO = 0;
    /** データ行として扱う最小列数 */
    private static final int MIN_COLUMNS = 12;

    private static final double SUBTOTAL_TOLERANCE = 0.5;

    private TorayCsvReader() {
    }

    /**
     * ①CSVを読み、対象入庫場所の契約NO別金額を集計する。
     *
     * @param path  RVSHEETyyyymm.csv
     * @param basho 対象入庫場所（A010 / A010P）
     */
    public static TorayCsvData read(Path path, String basho) {
        String targetBasho = Norm.norm(basho);
        List<List<String>> rows = parseCsv(decodeCp932(path));

        boolean headerOk = false;
        for (int i = 0; i < Math.min(3, rows.size()); i++) {
            List<String> row = rows.get(i);
            if (row.size() > COL_AMOUNT && row.get(COL_AMOUNT).contains("金額")) {
                headerOk = true;
                break;
            }
        }
        if (!headerOk) {
            throw new VerifyException("①CSVの10列目に「金額」ヘッダーが見つかりません。フォーマット変更の可能性: " + path);
        }

        Map<String, Double> result = new LinkedHashMap<>();
        Map<String, Double> byBasho = new LinkedHashMap<>();
        Map<String, Set<String>> nyukoDates = new LinkedHashMap<>();
        Map<String, List<TorayCsvData.MinusRow>> minusRows = new LinkedHashMap<>();
        List<String> warnings = new ArrayList<>();
        List<TorayCsvData.SubtotalError> subtotalErrors = new ArrayList<>();

        String currentBasho = null;
        double blockSum = 0.0;
        Integer blockStart = null;
        List<TorayCsvData.Candidate> blockRows = new ArrayList<>();

        for (int index = 0; index < rows.size(); index++) {
            int rowNo = index + 1;
            List<String> cols = rows.get(index);
            if (cols.size() < MIN_COLUMNS) {
                continue;
            }
            String keiyakuRaw = cols.get(COL_KEIYAKU).trim();

            if (KEIYAKU_CSV_PATTERN.matcher(keiyakuRaw).matches()) {
                String a = cols.get(COL_BASHO).trim();
                if (!a.isEmpty()) {
                    currentBasho = Norm.norm(a);
                }
                Double parsed = parseAmount(cols.get(COL_AMOUNT));
                if (parsed == null) {
                    warnings.add("①CSV " + rowNo + "行目: 発注No." + keiyakuRaw
                            + " の金額「" + cols.get(COL_AMOUNT) + "」を数値化できず除外");
                    continue;
                }
                double amount = parsed;
                String sign = cols.get(COL_SIGN).trim();
                if ("-".equals(sign)) {
                    amount = -amount;
                } else if (!sign.isEmpty()) {
                    warnings.add("①CSV " + rowNo + "行目: 発注No." + keiyakuRaw
                            + " のH列(符号)に想定外の値「" + sign + "」。プラスとして集計");
                }
                String bashoKey = currentBasho == null ? "(不明)" : currentBasho;
                byBasho.merge(bashoKey, amount, Double::sum);
                blockSum += amount;
                if (blockStart == null) {
                    blockStart = rowNo;
                }
                String key = Norm.keiyaku(keiyakuRaw);
                blockRows.add(new TorayCsvData.Candidate(key, bashoKey, amount));
                if (bashoKey.equals(targetBasho)) {
                    result.merge(key, amount, Double::sum);
                    nyukoDates.computeIfAbsent(key, k -> new LinkedHashSet<>()).add(cols.get(COL_DATE).trim());
                    if ("-".equals(sign)) {
                        minusRows.computeIfAbsent(key, k -> new ArrayList<>())
                                .add(new TorayCsvData.MinusRow(rowNo, cols.get(COL_DATE).trim(), amount));
                    }
                }
            } else if ("T".equals(keiyakuRaw)) {
                Double tValue = parseAmount(cols.get(COL_AMOUNT));
                if (tValue != null && Math.abs(blockSum - tValue) > SUBTOTAL_TOLERANCE) {
                    subtotalErrors.add(buildSubtotalError(
                            blockStart == null ? rowNo : blockStart, rowNo, tValue, blockSum, blockRows));
                }
                blockSum = 0.0;
                blockStart = null;
                blockRows = new ArrayList<>();
            }
        }

        if (result.isEmpty()) {
            throw new VerifyException("①CSVから入庫場所 " + targetBasho + " のデータ行を1件も取り込めませんでした。"
                    + "フォーマット変更または入庫場所の指定を確認してください: " + path);
        }

        List<List<String>> rawRows = List.copyOf(rows);
        return new TorayCsvData(result, byBasho, nyukoDates, minusRows, subtotalErrors, warnings, rawRows);
    }

    private static TorayCsvData.SubtotalError buildSubtotalError(
            int blockStart, int tRow, double tValue, double blockSum, List<TorayCsvData.Candidate> blockRows) {
        double diff = blockSum - tValue;
        Map<String, Double> perKey = new LinkedHashMap<>();
        Map<String, String> bashoOf = new LinkedHashMap<>();
        Set<String> bashos = new java.util.TreeSet<>();
        for (TorayCsvData.Candidate row : blockRows) {
            perKey.merge(row.keiyaku(), row.amount(), Double::sum);
            bashoOf.put(row.keiyaku(), row.basho());
            bashos.add(row.basho());
        }
        List<TorayCsvData.Candidate> candidates = new ArrayList<>();
        for (Map.Entry<String, Double> e : perKey.entrySet()) {
            if (Math.abs(e.getValue() - diff) <= SUBTOTAL_TOLERANCE) {
                candidates.add(new TorayCsvData.Candidate(e.getKey(), bashoOf.get(e.getKey()), e.getValue()));
            }
        }
        return new TorayCsvData.SubtotalError(
                blockStart, tRow, String.join("/", bashos), tValue, blockSum, diff, candidates);
    }

    private static Double parseAmount(String cell) {
        String s = cell.trim().replace(",", "");
        if (s.isEmpty()) {
            return null;
        }
        try {
            return Double.valueOf(s);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    /** cp932（Windows-31J）でデコードする。 */
    static String decodeCp932(Path path) {
        Charset cp932 = charsetCp932();
        try {
            byte[] bytes = Files.readAllBytes(path);
            CharsetDecoder decoder = cp932.newDecoder()
                    .onMalformedInput(CodingErrorAction.REPORT)
                    .onUnmappableCharacter(CodingErrorAction.REPORT);
            return decoder.decode(java.nio.ByteBuffer.wrap(bytes)).toString();
        } catch (java.nio.charset.CharacterCodingException e) {
            throw new VerifyException("①CSVがcp932でデコードできません: " + path + "\n" + e.getMessage(), e);
        } catch (IOException e) {
            throw new VerifyException("①CSVを読み込めません: " + path + " (" + e.getMessage() + ")", e);
        }
    }

    private static Charset charsetCp932() {
        for (String name : new String[] {"windows-31j", "Shift_JIS", "MS932"}) {
            try {
                return Charset.forName(name);
            } catch (RuntimeException ignored) {
                // 次の候補へ
            }
        }
        return StandardCharsets.UTF_8;
    }

    /** RFC4180 準拠の簡易CSVパーサ（引用符内の改行・エスケープに対応）。 */
    static List<List<String>> parseCsv(String text) {
        List<List<String>> rows = new ArrayList<>();
        List<String> row = new ArrayList<>();
        StringBuilder field = new StringBuilder();
        boolean inQuotes = false;
        boolean rowStarted = false;
        int i = 0;
        int n = text.length();

        while (i < n) {
            char ch = text.charAt(i);
            if (inQuotes) {
                if (ch == '"') {
                    if (i + 1 < n && text.charAt(i + 1) == '"') {
                        field.append('"');
                        i += 2;
                    } else {
                        inQuotes = false;
                        i++;
                    }
                } else {
                    field.append(ch);
                    i++;
                }
                continue;
            }
            switch (ch) {
                case '"' -> {
                    inQuotes = true;
                    rowStarted = true;
                    i++;
                }
                case ',' -> {
                    row.add(field.toString());
                    field.setLength(0);
                    rowStarted = true;
                    i++;
                }
                case '\r', '\n' -> {
                    i += (ch == '\r' && i + 1 < n && text.charAt(i + 1) == '\n') ? 2 : 1;
                    row.add(field.toString());
                    field.setLength(0);
                    rows.add(row);
                    row = new ArrayList<>();
                    rowStarted = false;
                }
                default -> {
                    field.append(ch);
                    rowStarted = true;
                    i++;
                }
            }
        }
        if (rowStarted || field.length() > 0 || !row.isEmpty()) {
            row.add(field.toString());
            rows.add(row);
        }
        return rows;
    }
}
