package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * 検証実行前の①②③検出（UI表用）。欠落は {@code missing=true}。
 * 湖南②はファイル名から当月を特定し、作業中 xlsm を開かない。
 */
public final class KouchinDiscovery {

    public record Row(String role, String path, String ym, boolean missing, String note) {}

    private KouchinDiscovery() {}

    public static List<Row> scan(FactoryId factory, KouchinPaths paths) {
        FactoryProfile profile = FactoryProfile.of(factory);
        List<Row> rows = new ArrayList<>();
        Path dir1 = paths.torayCsvDir();
        Path toray = null;
        YearMonthKey ym = null;
        try {
            if (Files.isDirectory(dir1)) {
                toray = FileDiscovery.findTorayCsv(dir1);
                ym = FileDiscovery.torayTargetYm(toray).orElse(null);
                rows.add(new Row("①東レCSV", toray.toString(), ym == null ? "" : ym.gatsudoLabel(), false, ""));
            } else {
                rows.add(new Row("①東レCSV", dir1.toString(), "", true, "フォルダなし"));
            }
        } catch (RuntimeException e) {
            rows.add(new Row("①東レCSV", dir1.toString(), "", true, e.getMessage()));
        }
        Path dir2 = paths.source2Dir(factory);
        try {
            Path current2 = FileDiscovery.findSource2CurrentForUi(profile, dir2, ym);
            rows.add(new Row("②" + profile.name2(), current2.toString(),
                    ym == null ? "" : ym.gatsudoLabel(), false, ""));
        } catch (RuntimeException e) {
            rows.add(new Row("②" + profile.name2(), dir2.toString(), "", true, e.getMessage()));
        }
        Path dir3 = paths.source3Dir(factory);
        try {
            Path a = FileDiscovery.findAladdin(dir3, ym);
            rows.add(new Row("③アラジン", a.toString(), ym == null ? "" : ym.gatsudoLabel(), false, ""));
        } catch (RuntimeException e) {
            rows.add(new Row("③アラジン", dir3.toString(), "", true, e.getMessage()));
        }
        if (factory == FactoryId.KONAN) {
            Path monthly = CheckVerifyC.findMonthlyFile(paths.konanMonthlyDir(), ym);
            if (monthly == null) {
                rows.add(new Row("湖南 月次処理", String.valueOf(paths.konanMonthlyDir()), "", true, "見つかりません"));
            } else {
                rows.add(new Row("湖南 月次処理", monthly.toString(), ym == null ? "" : ym.gatsudoLabel(), false, ""));
            }
        }
        return List.copyOf(rows);
    }
}
