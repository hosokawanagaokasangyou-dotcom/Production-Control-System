package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

/**
 * {@code 手動判定.csv} / {@code 前月過不足.csv} の読込。
 * 既定パスは国分固定の {@code ●自動検証}（{@link KouchinPaths#judgmentDir}）。BASE には追従しない。
 */
public final class JudgmentCsv {

    public static final String MANUAL_FILE = "手動判定.csv";
    public static final String PRIOR_FILE = "前月過不足.csv";

    private JudgmentCsv() {}

    public record ManualEntry(int side, String reason) {}

    public record ManualLoad(Map<String, ManualEntry> byKeiyaku, List<String> warnings) {}

    public record PriorRow(
            String keiyaku, String irai, double amount1, double amount2, double diff, String reason) {}

    public record PriorLoad(List<PriorRow> rows, List<String> warnings) {}

    public static Path manualFile(KouchinPaths paths) {
        return paths.judgmentDir().resolve(MANUAL_FILE);
    }

    public static Path priorFile(KouchinPaths paths) {
        return paths.judgmentDir().resolve(PRIOR_FILE);
    }

    public static ManualLoad loadManual(Path path, String factoryLabel, YearMonthKey targetYm) {
        Map<String, ManualEntry> result = new LinkedHashMap<>();
        List<String> warnings = new ArrayList<>();
        if (path == null || !Files.isRegularFile(path)) {
            return new ManualLoad(Map.of(), List.of());
        }
        List<List<String>> rows;
        try {
            rows = readCsvRows(path);
        } catch (IOException e) {
            warnings.add(path.getFileName() + ": 読めません: " + e.getMessage());
            return new ManualLoad(Map.of(), List.copyOf(warnings));
        }
        if (rows.isEmpty()) {
            return new ManualLoad(Map.of(), List.of());
        }
        List<String> hdr = rows.get(0).stream().map(Norm::norm).toList();
        if (!hdr.contains("契約NO") || !hdr.contains("正とする側")) {
            warnings.add(path.getFileName() + ": ヘッダーが想定と異なるため無視しました (必要: 工場,対象月,契約NO,正とする側,理由)");
            return new ManualLoad(Map.of(), List.copyOf(warnings));
        }
        Map<String, Integer> col = indexOf(hdr);
        for (int n = 1; n < rows.size(); n++) {
            List<String> r = rows.get(n);
            if (!factoryMatch(cell(r, col, "工場"), factoryLabel)) {
                continue;
            }
            String ymText = cell(r, col, "対象月");
            Optional<YearMonthKey> ym = YearMonthKey.parseFlexible(ymText);
            if (ym.isEmpty()) {
                warnings.add(path.getFileName() + " " + (n + 1) + "行目: 対象月「" + ymText
                        + "」を読めないため無視しました (例: 2026年8月度)");
                continue;
            }
            if (targetYm != null && !ym.get().equals(targetYm)) {
                continue;
            }
            String k = Norm.keiyaku(cell(r, col, "契約NO"));
            String sideTxt = Norm.norm(cell(r, col, "正とする側"));
            Integer side = parseSide(sideTxt);
            if (k.isEmpty() || side == null) {
                warnings.add(path.getFileName() + " " + (n + 1) + "行目: 契約NO「" + cell(r, col, "契約NO")
                        + "」正とする側「" + cell(r, col, "正とする側") + "」が不正のため無視しました (正とする側は ① か ②)");
                continue;
            }
            result.put(k, new ManualEntry(side, cell(r, col, "理由")));
        }
        return new ManualLoad(Map.copyOf(result), List.copyOf(warnings));
    }

    public static PriorLoad loadPrior(Path path, String factoryLabel, YearMonthKey targetYm) {
        List<PriorRow> result = new ArrayList<>();
        List<String> warnings = new ArrayList<>();
        if (path == null || !Files.isRegularFile(path)) {
            return new PriorLoad(List.of(), List.of());
        }
        List<List<String>> rows;
        try {
            rows = readCsvRows(path);
        } catch (IOException e) {
            warnings.add(path.getFileName() + ": 読めません: " + e.getMessage());
            return new PriorLoad(List.of(), List.copyOf(warnings));
        }
        if (rows.isEmpty()) {
            return new PriorLoad(List.of(), List.of());
        }
        List<String> hdr = rows.get(0).stream().map(Norm::norm).toList();
        Map<String, Integer> col = new LinkedHashMap<>();
        for (int i = 0; i < hdr.size(); i++) {
            col.put(priorAlias(hdr.get(i)), i);
        }
        if (!col.containsKey("契約NO") || !col.containsKey("①東レ金額") || !col.containsKey("②長岡金額")) {
            warnings.add(path.getFileName() + ": ヘッダーが想定と異なるため無視しました (必要: 工場,対象月,契約NO,依頼NO,①東レ金額,②長岡金額,理由)");
            return new PriorLoad(List.of(), List.copyOf(warnings));
        }
        for (int n = 1; n < rows.size(); n++) {
            List<String> r = rows.get(n);
            if (!factoryMatch(cell(r, col, "工場"), factoryLabel)) {
                continue;
            }
            String ymText = cell(r, col, "対象月");
            Optional<YearMonthKey> ym = YearMonthKey.parseFlexible(ymText);
            if (ym.isEmpty()) {
                warnings.add(path.getFileName() + " " + (n + 1) + "行目: 対象月「" + ymText
                        + "」を読めないため無視しました (例: 2026年8月度)");
                continue;
            }
            if (targetYm != null && ym.get().compareTo(targetYm) > 0) {
                continue;
            }
            String k = Norm.keiyaku(cell(r, col, "契約NO"));
            String irai = Norm.norm(cell(r, col, "依頼NO"));
            Double a1;
            Double a2;
            try {
                a1 = parseAmount(cell(r, col, "①東レ金額"));
                a2 = parseAmount(cell(r, col, "②長岡金額"));
            } catch (NumberFormatException e) {
                warnings.add(path.getFileName() + " " + (n + 1) + "行目: 金額を数値化できないため無視しました");
                continue;
            }
            if (k.isEmpty() || a1 == null || a2 == null) {
                warnings.add(path.getFileName() + " " + (n + 1) + "行目: 契約NO・①東レ金額・②長岡金額は必須です");
                continue;
            }
            result.add(new PriorRow(k, irai, a1, a2, a1 - a2, cell(r, col, "理由")));
        }
        return new PriorLoad(List.copyOf(result), List.copyOf(warnings));
    }

    /**
     * 当月差異の契約NOを、次月の検証で読む {@code 前月過不足.csv} へ足す。
     * 同じ工場・次月・契約NOが既にあれば足さない。
     *
     * @return 追加した件数
     */
    public static int appendCarryForward(Path path, VerifyResult result) throws IOException {
        if (path == null || result == null || result.targetYm() == null || result.recordsA() == null) {
            return 0;
        }
        YearMonthKey next = result.targetYm().plusMonths(1);
        String nextLabel = next.gatsudoLabel();
        String factory = result.profile() == null ? "" : result.profile().label();
        List<List<String>> existing = Files.isRegularFile(path) ? readCsvRows(path) : new ArrayList<>();
        Set<String> seen = new LinkedHashSet<>();
        for (int i = 1; i < existing.size(); i++) {
            List<String> row = existing.get(i);
            if (row.size() < 3) {
                continue;
            }
            seen.add(keyOf(row.get(0), row.get(1), row.get(2)));
        }
        List<String> added = new ArrayList<>();
        for (RecordA rec : result.recordsA()) {
            if (rec == null || !carryJudge(rec.judge()) || rec.keiyaku() == null || rec.keiyaku().isBlank()) {
                continue;
            }
            String key = keyOf(factory, nextLabel, rec.keiyaku());
            if (!seen.add(key)) {
                continue;
            }
            double a1 = rec.amount1() == null ? 0 : rec.amount1();
            double a2 = rec.amount2() == null ? 0 : rec.amount2();
            String reason = result.targetYm().gatsudoLabel() + "の当月差異（" + rec.judge()
                    + "）。次月（" + nextLabel + "）の検証へ繰越";
            added.add(String.join(",",
                    csv(factory),
                    csv(nextLabel),
                    csv(rec.keiyaku()),
                    csv(rec.iraiNo() == null ? "" : rec.iraiNo()),
                    csv(money(a1)),
                    csv(money(a2)),
                    csv(reason)));
        }
        if (added.isEmpty()) {
            return 0;
        }
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        StringBuilder body = new StringBuilder();
        if (existing.isEmpty()) {
            body.append("工場,対象月,契約NO,依頼NO,①東レ金額,②長岡金額,理由\n");
        } else {
            byte[] raw = Files.readAllBytes(path);
            String text = decodeCsv(raw);
            body.append(text);
            if (!text.endsWith("\n")) {
                body.append('\n');
            }
        }
        for (String line : added) {
            body.append(line).append('\n');
        }
        byte[] out = body.toString().getBytes(StandardCharsets.UTF_8);
        byte[] withBom = new byte[out.length + 3];
        withBom[0] = (byte) 0xEF;
        withBom[1] = (byte) 0xBB;
        withBom[2] = (byte) 0xBF;
        System.arraycopy(out, 0, withBom, 3, out.length);
        Files.write(path, withBom);
        return added.size();
    }

    private static boolean carryJudge(String judge) {
        return Judge.MISMATCH.equals(judge)
                || Judge.MANUAL_2.equals(judge)
                || Judge.NEXT_MONTH.equals(judge)
                || Judge.ONLY_1.equals(judge)
                || Judge.ONLY_2.equals(judge);
    }

    private static String keyOf(String factory, String ym, String keiyaku) {
        return Norm.norm(factory) + "|" + Norm.norm(ym) + "|" + Norm.keiyaku(keiyaku);
    }

    private static String money(double value) {
        return String.format(Locale.US, "%.0f", value);
    }

    private static String csv(String value) {
        String text = value == null ? "" : value;
        if (text.contains(",") || text.contains("\"") || text.contains("\n")) {
            return "\"" + text.replace("\"", "\"\"") + "\"";
        }
        return text;
    }

    private static String decodeCsv(byte[] raw) {
        if (raw.length >= 3 && (raw[0] & 0xFF) == 0xEF && (raw[1] & 0xFF) == 0xBB && (raw[2] & 0xFF) == 0xBF) {
            return new String(raw, 3, raw.length - 3, StandardCharsets.UTF_8);
        }
        if (looksLikeUtf8(raw)) {
            return new String(raw, StandardCharsets.UTF_8);
        }
        return new String(raw, Charset.forName("MS932"));
    }

    static boolean factoryMatch(String facCell, String factoryLabel) {
        String cell = Norm.norm(facCell);
        if (cell.isEmpty()) {
            return true;
        }
        String fl = Norm.norm(factoryLabel);
        return cell.equals(fl) || cell.equals(fl.replace("工場", ""));
    }

    static Integer parseSide(String sideTxt) {
        if (sideTxt == null) {
            return null;
        }
        if (sideTxt.equals("1") || sideTxt.equals("①") || sideTxt.equals("東レ") || sideTxt.equals("トウレ")) {
            return 1;
        }
        if (sideTxt.equals("2") || sideTxt.equals("②") || sideTxt.equals("長岡") || sideTxt.equals("ナガオカ")) {
            return 2;
        }
        return null;
    }

    static Double parseAmount(String text) {
        String t = (text == null ? "" : text).trim().replace(",", "").replace("円", "");
        if (t.isEmpty()) {
            return null;
        }
        return Double.parseDouble(t);
    }

    static List<List<String>> readCsvRows(Path path) throws IOException {
        byte[] raw = Files.readAllBytes(path);
        String text;
        if (raw.length >= 3 && (raw[0] & 0xFF) == 0xEF && (raw[1] & 0xFF) == 0xBB && (raw[2] & 0xFF) == 0xBF) {
            text = new String(raw, 3, raw.length - 3, StandardCharsets.UTF_8);
        } else if (looksLikeUtf8(raw)) {
            text = new String(raw, StandardCharsets.UTF_8);
        } else {
            text = new String(raw, Charset.forName("MS932"));
        }
        List<List<String>> out = new ArrayList<>();
        for (String line : text.split("\\r?\\n")) {
            List<String> cols = parseCsvLine(line);
            boolean any = false;
            for (String c : cols) {
                if (c != null && !c.isBlank()) {
                    any = true;
                    break;
                }
            }
            if (any) {
                out.add(cols);
            }
        }
        return out;
    }

    private static boolean looksLikeUtf8(byte[] raw) {
        try {
            String s = new String(raw, StandardCharsets.UTF_8);
            byte[] back = s.getBytes(StandardCharsets.UTF_8);
            if (raw.length >= 3 && (raw[0] & 0xFF) == 0xEF) {
                return true;
            }
            return java.util.Arrays.equals(raw, back) || s.indexOf('\uFFFD') < 0;
        } catch (RuntimeException e) {
            return false;
        }
    }

    static List<String> parseCsvLine(String line) {
        List<String> out = new ArrayList<>();
        StringBuilder cur = new StringBuilder();
        boolean quoted = false;
        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);
            if (quoted) {
                if (c == '"') {
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') {
                        cur.append('"');
                        i++;
                    } else {
                        quoted = false;
                    }
                } else {
                    cur.append(c);
                }
            } else if (c == '"') {
                quoted = true;
            } else if (c == ',') {
                out.add(cur.toString());
                cur.setLength(0);
            } else {
                cur.append(c);
            }
        }
        out.add(cur.toString());
        return out;
    }

    private static Map<String, Integer> indexOf(List<String> hdr) {
        Map<String, Integer> col = new LinkedHashMap<>();
        for (int i = 0; i < hdr.size(); i++) {
            col.put(hdr.get(i), i);
        }
        return col;
    }

    private static String cell(List<String> r, Map<String, Integer> col, String name) {
        Integer i = col.get(name);
        if (i == null || i >= r.size()) {
            return "";
        }
        return r.get(i) == null ? "" : r.get(i).trim();
    }

    private static String priorAlias(String h) {
        return switch (h) {
            case "1東レ金額", "①金額", "1金額", "東レ金額", "①", "1" -> "①東レ金額";
            case "2長岡金額", "②金額", "2金額", "長岡金額", "②", "2" -> "②長岡金額";
            default -> h;
        };
    }
}
