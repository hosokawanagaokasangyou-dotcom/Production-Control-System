package jp.co.pm.ai.kouchin.trend;

import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.TreeMap;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.kouchin.verify.ExcelValues;
import jp.co.pm.ai.kouchin.verify.FactoryProfile;
import jp.co.pm.ai.kouchin.verify.FileDiscovery;
import jp.co.pm.ai.kouchin.verify.VerifyException;
import jp.co.pm.ai.kouchin.verify.YearMonthKey;

/**
 * 国分年度フォルダ／湖南試算ルートから月次ファイルを列挙する。
 * {@code .lnk} は読まない。絶対パスの実フォルダのみ。
 */
public final class TrendScan {

    private static final Pattern NENDO_DIR = Pattern.compile("^(\\d{4})年度試算");
    private static final String KOKUBU_GLOB = "*後加工工賃明細*.xls*";
    private static final String KONAN_GLOB = "*加工賃試算*.xls*";

    public record Period(List<YearMonthKey> months, Map<YearMonthKey, Path> files) {
        public Period {
            months = months == null ? List.of() : List.copyOf(months);
            files = files == null ? Map.of() : Map.copyOf(files);
        }
    }

    private TrendScan() {}

    public static Map<YearMonthKey, Path> listKokubuFiles(Path folder) {
        if (folder == null) {
            return Map.of();
        }
        return listKokubuFiles(List.of(folder));
    }

    public static Map<YearMonthKey, Path> listKokubuFiles(Collection<Path> folders) {
        Map<YearMonthKey, Path> best = new TreeMap<>();
        if (folders == null) {
            return best;
        }
        for (Path dir : folders) {
            if (dir == null || !Files.isDirectory(dir)) {
                continue;
            }
            for (Path file : listFiles(dir, KOKUBU_GLOB)) {
                Optional<YearMonthKey> ym = YearMonthKey.parseGatsudo(file.getFileName().toString());
                if (ym.isEmpty()) {
                    continue;
                }
                Path cur = best.get(ym.get());
                if (cur == null || lastModified(file) > lastModified(cur)) {
                    best.put(ym.get(), file);
                }
            }
        }
        return best;
    }

    public static Map<YearMonthKey, Path> listKonanFiles(Path root) {
        Map<YearMonthKey, Path> best = new TreeMap<>();
        if (root == null || !Files.isDirectory(root)) {
            return best;
        }
        List<Path> cands = new ArrayList<>(listFiles(root, KONAN_GLOB));
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(root)) {
            for (Path dir : stream) {
                if (!Files.isDirectory(dir)) {
                    continue;
                }
                String name = dir.getFileName().toString();
                if ("temp".equals(name) || "バックアップ".equals(name) || name.contains("バックアップ")) {
                    continue;
                }
                String nfkc = Normalizer.normalize(name, Normalizer.Form.NFKC);
                if (!NENDO_DIR.matcher(nfkc).find()) {
                    continue;
                }
                cands.addAll(listFiles(dir, KONAN_GLOB));
            }
        } catch (IOException ignored) {
            // サブフォルダが読めなければ直下のみ
        }
        for (Path file : cands) {
            Optional<YearMonthKey> ym = shisanYm(file);
            if (ym.isEmpty()) {
                continue;
            }
            Path cur = best.get(ym.get());
            if (cur == null || lastModified(file) > lastModified(cur)) {
                best.put(ym.get(), file);
            }
        }
        return best;
    }

    public static Period pickPeriod(Map<YearMonthKey, Path> available, int n) {
        return pickPeriod(available, n, null);
    }

    public static Period pickPeriod(Map<YearMonthKey, Path> available, int n, YearMonthKey end) {
        Map<YearMonthKey, Path> src = available == null ? Map.of() : available;
        if (src.isEmpty() && end == null) {
            return new Period(List.of(), Map.of());
        }
        YearMonthKey endYm = end != null ? end : src.keySet().stream().max(YearMonthKey::compareTo).orElse(null);
        if (endYm == null) {
            return new Period(List.of(), Map.of());
        }
        int count = Math.max(1, n);
        List<YearMonthKey> months = monthRange(endYm, count);
        Map<YearMonthKey, Path> files = new LinkedHashMap<>();
        for (YearMonthKey ym : months) {
            Path p = src.get(ym);
            if (p != null) {
                files.put(ym, p);
            }
        }
        return new Period(months, files);
    }

    public static List<YearMonthKey> monthRange(YearMonthKey end, int n) {
        List<YearMonthKey> months = new ArrayList<>(n);
        for (int i = 0; i < n; i++) {
            months.add(end.plusMonths(i - (n - 1)));
        }
        return months;
    }

    @SafeVarargs
    public static YearMonthKey resolveEndYm(Map<YearMonthKey, Path>... availables) {
        YearMonthKey max = null;
        if (availables == null) {
            return null;
        }
        for (Map<YearMonthKey, Path> d : availables) {
            if (d == null) {
                continue;
            }
            for (YearMonthKey ym : d.keySet()) {
                if (max == null || ym.compareTo(max) > 0) {
                    max = ym;
                }
            }
        }
        return max;
    }

    public static Optional<YearMonthKey> shisanYm(Path path) {
        if (path == null) {
            return Optional.empty();
        }
        try (Workbook wb = ExcelValues.open(path)) {
            for (String name : FactoryProfile.SHISAN_SHEETS) {
                if (wb.getSheet(name) == null) {
                    continue;
                }
                List<List<Object>> rows = ExcelValues.readSheet(wb, name);
                int limit = Math.min(15, rows.size());
                for (int i = 0; i < limit; i++) {
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

    private static List<Path> listFiles(Path dir, String glob) {
        try {
            List<Path> files = new ArrayList<>();
            for (Path f : FileDiscovery.list(dir, glob)) {
                String name = f.getFileName().toString();
                if (name.toLowerCase(Locale.ROOT).endsWith(".lnk")) {
                    continue;
                }
                files.add(f);
            }
            return files;
        } catch (VerifyException e) {
            return List.of();
        }
    }

    private static long lastModified(Path path) {
        try {
            return Files.getLastModifiedTime(path).toMillis();
        } catch (IOException e) {
            return 0L;
        }
    }
}
