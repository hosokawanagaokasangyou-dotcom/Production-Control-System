package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 各データソースの対象月ファイルを自動検出する。
 * 対象月は①のファイル名 {@code RVSHEETyyyymm} を基準とする。
 */
public final class FileDiscovery {

    /**
     * 過去月明細として読む範囲（当月から何か月前まで）。
     * 年を含まない依頼NO（C8-1 等）が前年同月と衝突しないよう 12 未満にする。
     */
    public static final int PREV_MONTH_WINDOW = 6;

    /** 「2026年度試算　湖南」（暦年。H30年度等の旧フォルダは対象外） */
    private static final Pattern NENDO_DIR_PATTERN = Pattern.compile("^(\\d{4})年度試算");
    /** アラジンのファイル名 {@code _yyyymmdd_hhmmss} */
    private static final Pattern ALADDIN_STAMP_PATTERN = Pattern.compile("_(\\d{8})_(\\d{6})");
    private static final Pattern YEAR_DIR_PATTERN = Pattern.compile("^\\d{4}年$");

    private FileDiscovery() {
    }

    // ---- ① 東レ送付CSV ----

    /** ①をファイル名 RVSHEETyyyymm の年月最大で選択する。 */
    public static Path findTorayCsv(Path folder) {
        List<Path> files = list(folder, "RVSHEET*.csv");
        Path best = null;
        YearMonthKey bestYm = null;
        for (Path f : files) {
            Optional<YearMonthKey> ym = YearMonthKey.parseRvsheet(f.getFileName().toString());
            if (ym.isEmpty()) {
                continue;
            }
            if (bestYm == null || ym.get().compareTo(bestYm) > 0) {
                bestYm = ym.get();
                best = f;
            }
        }
        if (best == null) {
            throw new VerifyException("RVSHEETyyyymm.csv が見つかりません: " + folder);
        }
        return best;
    }

    /** ①ファイル名から対象年月を取得する。 */
    public static Optional<YearMonthKey> torayTargetYm(Path csv) {
        return YearMonthKey.parseRvsheet(csv.getFileName().toString());
    }

    // ---- ② 工場別 ----

    /** 工場プロファイルに応じて②ファイルを当月・過去月・翌月に振り分ける。 */
    public static MonthlyFileSet findSource2(FactoryProfile profile, Path folder, YearMonthKey targetYm) {
        return profile.id() == FactoryId.KOKUBU
                ? findNagaokaFiles(folder, targetYm)
                : findShisanFiles(folder, targetYm);
    }

    /** ②ファイル1件の年月（国分=ファイル名 / 湖南=シート内）。 */
    public static Optional<YearMonthKey> ymOfSource2(FactoryProfile profile, Path file) {
        return profile.id() == FactoryId.KOKUBU ? ymFromFilename(file) : ShisanReader.targetYm(file);
    }

    /** ファイル名の「yyyy年m月度」。 */
    public static Optional<YearMonthKey> ymFromFilename(Path path) {
        return YearMonthKey.parseGatsudo(path.getFileName().toString());
    }

    /** 国分② 後加工工賃明細*.xlsx をファイル名の年月で振り分ける。 */
    public static MonthlyFileSet findNagaokaFiles(Path folder, YearMonthKey targetYm) {
        Map<YearMonthKey, List<Path>> dated = new TreeMap<>();
        for (Path f : list(folder, "後加工工賃明細*.xlsx")) {
            ymFromFilename(f).ifPresent(ym -> dated.computeIfAbsent(ym, k -> new ArrayList<>()).add(f));
        }
        if (dated.isEmpty()) {
            throw new VerifyException("「yyyy年m月度」を含む 後加工工賃明細*.xlsx が見つかりません: " + folder);
        }
        return pickByYm(dated, targetYm, folder);
    }

    /** 湖南② *加工賃試算*.xlsm をシート内の年月で振り分ける。 */
    public static MonthlyFileSet findShisanFiles(Path root, YearMonthKey targetYm) {
        Map<YearMonthKey, List<Path>> dated = new TreeMap<>();
        for (Path f : shisanCandidates(root, targetYm)) {
            ShisanReader.targetYm(f).ifPresent(ym -> dated.computeIfAbsent(ym, k -> new ArrayList<>()).add(f));
        }
        if (dated.isEmpty()) {
            throw new VerifyException("東レシートに「yyyy年m月度」を持つ *加工賃試算*.xlsm が見つかりません: "
                    + root + " (直下および「yyyy年度試算　湖南」フォルダ)");
        }
        return pickByYm(dated, targetYm, root);
    }

    /**
     * 湖南②の候補ファイル。
     * ルート直下の「月度加工賃試算.xlsm」= 当月の作業中ファイル、
     * 「yyyy年度試算　湖南\m月度加工賃試算.xlsm」= 確定済みの各月。
     */
    private static List<Path> shisanCandidates(Path root, YearMonthKey targetYm) {
        List<Path> files = new ArrayList<>(list(root, "*加工賃試算*.xls*"));
        Map<Integer, Path> yearDirs = new TreeMap<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(root)) {
            for (Path d : stream) {
                if (!Files.isDirectory(d)) {
                    continue;
                }
                Matcher m = NENDO_DIR_PATTERN.matcher(Norm.norm(d.getFileName().toString()));
                if (m.find()) {
                    yearDirs.put(Integer.parseInt(m.group(1)), d);
                }
            }
        } catch (IOException e) {
            throw new VerifyException("②のフォルダを走査できません: " + root + " (" + e.getMessage() + ")", e);
        }

        List<Integer> years = new ArrayList<>();
        if (targetYm != null) {
            years.add(targetYm.year() - 1);
            years.add(targetYm.year());
            years.add(targetYm.year() + 1);
        } else {
            List<Integer> all = new ArrayList<>(yearDirs.keySet());
            years.addAll(all.subList(Math.max(0, all.size() - 2), all.size()));
        }
        for (Integer y : years) {
            Path dir = yearDirs.get(y);
            if (dir != null) {
                files.addAll(list(dir, "*加工賃試算*.xls*"));
            }
        }
        return files;
    }

    /**
     * 年月→ファイル群から 当月/過去月/翌月以降 を振り分ける（同一年月は更新日時最新を採用）。
     * 過去月は直近 {@link #PREV_MONTH_WINDOW} か月に限る。
     */
    public static MonthlyFileSet pickByYm(Map<YearMonthKey, List<Path>> dated, YearMonthKey targetYm, Path folder) {
        List<YearMonthKey> keys = new ArrayList<>(dated.keySet());
        keys.sort(Comparator.naturalOrder());

        YearMonthKey current;
        if (targetYm == null) {
            current = keys.get(keys.size() - 1);
        } else if (dated.containsKey(targetYm)) {
            current = targetYm;
        } else {
            StringBuilder sb = new StringBuilder();
            for (YearMonthKey k : keys) {
                if (sb.length() > 0) {
                    sb.append(", ");
                }
                sb.append(k.gatsudoLabel());
            }
            throw new VerifyException("②に " + targetYm.ymLabel() + "度 のファイルがありません: " + folder
                    + "\n  存在する月度: " + sb);
        }

        Path currentFile = newest(dated.get(current));
        List<DatedFile> prevs = new ArrayList<>();
        List<DatedFile> nexts = new ArrayList<>();
        for (YearMonthKey k : keys) {
            int delta = current.monthIndex() - k.monthIndex();
            if (delta > 0 && delta <= PREV_MONTH_WINDOW) {
                prevs.add(new DatedFile(k.gatsudoLabel(), k, newest(dated.get(k))));
            } else if (delta < 0) {
                nexts.add(new DatedFile(k.gatsudoLabel(), k, newest(dated.get(k))));
            }
        }
        prevs.sort(Comparator.comparing((DatedFile f) -> f.ym().monthIndex()).reversed());
        nexts.sort(Comparator.comparingInt(f -> f.ym().monthIndex()));
        return new MonthlyFileSet(currentFile, List.copyOf(prevs), List.copyOf(nexts));
    }

    // ---- ③ アラジン ----

    /**
     * ③の全照会ファイルを対象年月（シート内）ごとに整理する。
     * 同一年月に複数あればファイル名 {@code _yyyymmdd_hhmmss} が最新のものを採用する。
     */
    public static Map<YearMonthKey, Path> findAladdinAll(Path folder) {
        List<Path> files = list(folder, "依頼NO別問合せ*.xlsx");
        if (files.isEmpty()) {
            throw new VerifyException("依頼NO別問合せ*.xlsx が見つかりません: " + folder);
        }
        files.sort(Comparator.comparing(FileDiscovery::aladdinSortKey));
        Map<YearMonthKey, Path> result = new TreeMap<>();
        for (Path f : files) {
            AladdinReader.targetYm(f).ifPresent(ym -> result.put(ym, f));
        }
        if (result.isEmpty()) {
            throw new VerifyException("③のシート内「対象年月」を読めるファイルがありません: " + folder);
        }
        return result;
    }

    /** 対象年月が①の対象月と一致する③を選ぶ。 */
    public static Path findAladdin(Path folder, YearMonthKey targetYm) {
        Map<YearMonthKey, Path> all = findAladdinAll(folder);
        if (targetYm == null) {
            YearMonthKey last = null;
            for (YearMonthKey k : all.keySet()) {
                last = k;
            }
            return all.get(last);
        }
        Path hit = all.get(targetYm);
        if (hit != null) {
            return hit;
        }
        StringBuilder listing = new StringBuilder();
        for (Map.Entry<YearMonthKey, Path> e : all.entrySet()) {
            if (listing.length() > 0) {
                listing.append(", ");
            }
            listing.append(e.getValue().getFileName()).append('(').append(e.getKey().ymLabel()).append(')');
        }
        throw new VerifyException("③に対象年月 " + targetYm.ymLabel()
                + " の依頼NO別問合せファイルがありません: " + folder + "\n  存在するファイル: " + listing);
    }

    /** ③の取得日時（ファイル名 {@code _yyyymmdd_hhmmss}）。無ければ空。 */
    public static Optional<String> aladdinStamp(Path path) {
        Matcher m = ALADDIN_STAMP_PATTERN.matcher(path.getFileName().toString());
        return m.find() ? Optional.of(m.group(1) + m.group(2)) : Optional.empty();
    }

    private static String aladdinSortKey(Path path) {
        return aladdinStamp(path).orElseGet(() -> "0" + String.format(Locale.US, "%015d", lastModified(path)));
    }

    // ---- 湖南 月次処理ファイル（検証C用・任意） ----

    /** 対象月の月次処理ファイルを探す。無ければ null。 */
    public static Path findMonthlyFile(Path folder, YearMonthKey targetYm) {
        if (folder == null || !Files.isDirectory(folder)) {
            return null;
        }
        List<Path> folders = new ArrayList<>();
        folders.add(folder);
        if (targetYm != null) {
            folders.add(folder.resolve(targetYm.year() + "年"));
        } else {
            try (DirectoryStream<Path> stream = Files.newDirectoryStream(folder)) {
                List<Path> years = new ArrayList<>();
                for (Path d : stream) {
                    if (Files.isDirectory(d) && YEAR_DIR_PATTERN.matcher(Norm.norm(d.getFileName().toString())).matches()) {
                        years.add(d);
                    }
                }
                years.sort(Comparator.comparing((Path p) -> p.getFileName().toString()).reversed());
                if (!years.isEmpty()) {
                    folders.add(years.get(0));
                }
            } catch (IOException ignored) {
                // 年フォルダが読めなければ直下だけを見る
            }
        }
        List<Path> candidates = new ArrayList<>();
        for (Path dir : folders) {
            if (!Files.isDirectory(dir)) {
                continue;
            }
            for (Path f : list(dir, "*月次処理ファイル*.xls*")) {
                if (targetYm == null || ymFromFilename(f).map(targetYm::equals).orElse(false)) {
                    candidates.add(f);
                }
            }
        }
        return candidates.isEmpty() ? null : newest(candidates);
    }

    // ---- 共通 ----

    /** 一時ファイル（{@code ~$}）を除いた glob 検索。UNC はディレクトリ mtime でキャッシュする。 */
    public static List<Path> list(Path folder, String glob) {
        return KouchinDirListingCache.list(folder, glob);
    }

    public static void invalidateListingCache() {
        KouchinDirListingCache.invalidateAll();
    }

    static List<Path> listUncached(Path folder, String glob) {
        if (folder == null || !Files.isDirectory(folder)) {
            throw new VerifyException("フォルダにアクセスできません: " + folder);
        }
        List<Path> files = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(folder, glob)) {
            for (Path f : stream) {
                if (Files.isRegularFile(f) && !f.getFileName().toString().startsWith("~$")) {
                    files.add(f);
                }
            }
        } catch (IOException e) {
            throw new VerifyException("フォルダを走査できません: " + folder + " (" + e.getMessage() + ")", e);
        }
        files.sort(Comparator.comparing(p -> p.getFileName().toString()));
        return files;
    }

    /** 更新日時が最新のファイル。 */
    public static Path newest(List<Path> files) {
        Path best = null;
        long bestTime = Long.MIN_VALUE;
        for (Path f : files) {
            long t = lastModified(f);
            if (best == null || t > bestTime) {
                best = f;
                bestTime = t;
            }
        }
        return best;
    }

    private static long lastModified(Path path) {
        try {
            return Files.getLastModifiedTime(path).toMillis();
        } catch (IOException e) {
            return 0L;
        }
    }

    /** ②ファイルの表示名（参照フォルダ配下ならサブフォルダ付き）。 */
    public static String displayName(Path dir, Path file) {
        try {
            return dir.relativize(file).toString();
        } catch (IllegalArgumentException e) {
            return file.getFileName().toString();
        }
    }

    /** 表示名の一覧をカンマ区切りで返す。 */
    public static String displayNames(Path dir, List<DatedFile> files) {
        if (files.isEmpty()) {
            return "なし";
        }
        List<String> names = new ArrayList<>(files.size());
        for (DatedFile f : files) {
            names.add(displayName(dir, f.path()));
        }
        return String.join(", ", names);
    }

    /** 空でない要素だけを {@code sep} で連結する。 */
    public static String joinNonEmpty(String sep, String... parts) {
        List<String> kept = new ArrayList<>();
        for (String p : parts) {
            if (p != null && !p.isEmpty()) {
                kept.add(p);
            }
        }
        return String.join(sep, kept);
    }
}
