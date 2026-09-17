package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * 検証実行前の①②③検出（UI表用）。欠落は {@code missing=true}。
 * 湖南②はファイル名から当月を特定し、作業中 xlsm を開かない。
 * {@link Row#path()} は表の短縮表示、{@link Row#fullPath()} はツールチップ用の絶対パス。
 */
public final class KouchinDiscovery {

    public static final String ROLE_1 = "①東レCSV";
    /** ③の検出表・案内の表示名（依頼NO別問合せ＝月次実績）。 */
    public static final String ROLE_3 = "③月次実績";

    public record Row(String role, String path, String fullPath, String ym, boolean missing, String note) {}

    private KouchinDiscovery() {}

    public static List<Row> scan(FactoryId factory, KouchinPaths paths) {
        FactoryProfile profile = FactoryProfile.of(factory);
        List<Row> rows = new ArrayList<>();
        YearMonthKey ym = detectTargetYm(paths);
        Path dir1 = paths.source1Dir(factory);
        try {
            if (Files.isDirectory(dir1)) {
                Path toray = FileDiscovery.findTorayCsv(dir1);
                YearMonthKey fileYm = FileDiscovery.torayTargetYm(toray).orElse(ym);
                rows.add(found(ROLE_1, dir1, toray, fileYm));
            } else {
                rows.add(missing(ROLE_1, dir1, "フォルダなし"));
            }
        } catch (RuntimeException e) {
            rows.add(missing(ROLE_1, dir1, e.getMessage()));
        }
        Path dir2 = paths.source2Dir(factory);
        try {
            Path current2 = FileDiscovery.findSource2CurrentForUi(profile, dir2, ym);
            rows.add(found("②" + profile.name2(), dir2, current2, ym));
        } catch (RuntimeException e) {
            rows.add(missing("②" + profile.name2(), dir2, e.getMessage()));
        }
        Path dir3 = paths.source3Dir(factory);
        try {
            Path a = FileDiscovery.findAladdin(dir3, ym);
            rows.add(found(ROLE_3, dir3, a, ym, Source3TargetMonthCheck.noteIfUncovered(a, ym, profile)));
        } catch (RuntimeException e) {
            rows.add(missing(ROLE_3, dir3, Source3TargetMonthCheck.warningNote(ym, e.getMessage())));
        }
        if (factory == FactoryId.KONAN) {
            Path monthlyDir = paths.konanMonthlyDir();
            Path monthly = CheckVerifyC.findMonthlyFile(monthlyDir, ym);
            if (monthly == null) {
                rows.add(missing("湖南 月次処理", monthlyDir, "見つかりません"));
            } else {
                rows.add(found("湖南 月次処理", monthlyDir, monthly, ym));
            }
        }
        return List.copyOf(rows);
    }

    static YearMonthKey detectTargetYm(KouchinPaths paths) {
        if (paths == null) {
            return null;
        }
        try {
            Path dir = paths.resolveTorayCsvDir();
            if (dir != null && Files.isDirectory(dir)) {
                return FileDiscovery.torayTargetYm(FileDiscovery.findTorayCsv(dir)).orElse(null);
            }
        } catch (RuntimeException ignored) {
        }
        return null;
    }

    private static Row found(String role, Path dir, Path file, YearMonthKey ym) {
        return found(role, dir, file, ym, "");
    }

    private static Row found(String role, Path dir, Path file, YearMonthKey ym, String note) {
        return new Row(
                role,
                FileDiscovery.uiDisplayPath(dir, file),
                file.toString(),
                ym == null ? "" : ym.gatsudoLabel(),
                false,
                note == null ? "" : note);
    }

    private static Row missing(String role, Path dir, String note) {
        return new Row(
                role,
                FileDiscovery.uiDisplayFolder(dir),
                dir == null ? "" : dir.toString(),
                "",
                true,
                note == null ? "" : note);
    }
}
