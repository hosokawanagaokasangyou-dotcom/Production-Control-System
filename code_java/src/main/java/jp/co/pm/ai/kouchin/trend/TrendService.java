package jp.co.pm.ai.kouchin.trend;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.kouchin.verify.DualWriteFiles;
import jp.co.pm.ai.kouchin.verify.KouchinOutputDirs;
import jp.co.pm.ai.kouchin.verify.KouchinPaths;
import jp.co.pm.ai.kouchin.verify.ResultArchive;
import jp.co.pm.ai.kouchin.verify.VerifyException;
import jp.co.pm.ai.kouchin.verify.YearMonthKey;

/**
 * 後加工工賃の月次トレンド入口。①③は突合しない。②のみ。
 */
public final class TrendService {

    private static final DateTimeFormatter STAMP = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");

    private TrendService() {}

    public static TrendResult run(TrendRequest req) {
        TrendRequest request = req == null ? new TrendRequest(0, TrendRequest.FACTORY_BOTH, null, Map.of()) : req;
        Map<String, String> ui = request.ui();
        KouchinPaths paths = request.paths();
        List<String> warnings = new ArrayList<>();
        Map<YearMonthKey, Path> kokubuAvail = Map.of();
        Map<YearMonthKey, Path> konanAvail = Map.of();

        List<Path> kokubuDirs = kokubuYearDirs(ui);
        if (request.wantKokubu()) {
            List<String> missing = new ArrayList<>();
            boolean any = false;
            for (Path d : kokubuDirs) {
                if (Files.isDirectory(d)) {
                    any = true;
                } else {
                    missing.add(d.toString());
                }
            }
            for (String m : missing) {
                warnings.add("国分フォルダにアクセスできません: " + m);
            }
            if (any) {
                kokubuAvail = TrendScan.listKokubuFiles(kokubuDirs);
            }
            if (kokubuAvail.isEmpty() && TrendRequest.FACTORY_KOKUBU.equals(request.factory())) {
                throw new VerifyException("国分の後加工工賃明細が見つかりません");
            }
        }

        Path konanRoot = paths.konanShisanDir();
        boolean konanRootOk = konanRoot != null && Files.isDirectory(konanRoot);
        if (request.wantKonan()) {
            if (!konanRootOk) {
                String msg = "湖南試算ルートにアクセスできません: " + konanRoot;
                if (TrendRequest.FACTORY_KONAN.equals(request.factory())) {
                    throw new VerifyException(msg);
                }
                warnings.add(msg);
            } else {
                konanAvail = TrendScan.listKonanFiles(konanRoot);
            }
        }

        YearMonthKey end =
                TrendScan.resolveEndYm(
                        request.wantKokubu() ? kokubuAvail : Map.of(),
                        request.wantKonan() ? konanAvail : Map.of());
        if (end == null) {
            throw new VerifyException("対象工場で読める月次ファイルがありません");
        }
        TrendScan.Period kPeriod = TrendScan.pickPeriod(kokubuAvail, request.months(), end);
        TrendScan.Period nPeriod = TrendScan.pickPeriod(konanAvail, request.months(), end);
        List<YearMonthKey> months = kPeriod.months();
        List<TrendMonthData> kokubuData = new ArrayList<>();
        List<TrendMonthData> konanData = new ArrayList<>();
        for (YearMonthKey ym : months) {
            if (kPeriod.files().containsKey(ym)) {
                TrendMonthData agg = TrendKokubuReader.read(kPeriod.files().get(ym));
                warnings.addAll(agg.warnings());
                kokubuData.add(agg);
            } else if (request.wantKokubu()) {
                warnings.add("国分: " + ym.ymLabel() + " の明細がありません");
            }
            if (nPeriod.files().containsKey(ym)) {
                TrendMonthData agg = TrendKonanReader.read(nPeriod.files().get(ym));
                warnings.addAll(agg.warnings());
                konanData.add(agg);
            } else if (request.wantKonan() && (konanRootOk || !konanAvail.isEmpty())) {
                warnings.add("湖南: " + ym.ymLabel() + " の試算がありません");
            }
        }
        if (kokubuData.isEmpty() && konanData.isEmpty()) {
            throw new VerifyException("対象期間に集計できるデータがありません");
        }

        List<TrendMonthData> kokubuDisp = new ArrayList<>();
        for (TrendMonthData d : kokubuData) {
            kokubuDisp.add(toDisplay(d));
        }
        List<TrendMonthData> konanDisp = new ArrayList<>();
        for (TrendMonthData d : konanData) {
            konanDisp.add(toDisplay(d));
        }
        TrendCrossMatrix wage = TrendAggregate.buildGroupedCross(kokubuDisp, konanDisp, months, "wage");
        TrendCrossMatrix qty = TrendAggregate.buildGroupedCross(kokubuDisp, konanDisp, months, "qty");
        List<TrendShiftRow> suspects = TrendAggregate.rankShiftSuspects(wage, "wage");
        List<String> focus = TrendAggregate.pickFocusProcesses(wage, qty, 6);
        TrendModel model =
                new TrendModel(months, focus, wage, qty, suspects, kokubuDisp, konanDisp, List.copyOf(warnings));
        String stamp = LocalDateTime.now().format(STAMP);
        String fileName = "月トレンド_" + stamp + ".xlsx";
        List<Path> dirs = KouchinOutputDirs.resolveAll(ui);
        List<Path> targets = KouchinOutputDirs.withNames(dirs, fileName);
        XSSFWorkbook workbook = TrendExcelWriter.buildWorkbook(model);
        DualWriteFiles.WriteOutcome xlsx = DualWriteFiles.writeWorkbook(workbook, targets, ui);
        if (xlsx.anyFailed()) {
            warnings.addAll(xlsx.failures());
        }
        if (xlsx.allFailed()) {
            try {
                workbook.close();
            } catch (IOException ignored) {
                // 閉じられなくても失敗内容を優先
            }
            throw new VerifyException("月トレンドExcelを書けません: " + String.join("; ", xlsx.failures()));
        }
        for (Path dir : dirs) {
            Path base = dir.toAbsolutePath().normalize();
            List<Path> keepHere = new ArrayList<>();
            for (Path p : xlsx.succeeded()) {
                Path n = p.toAbsolutePath().normalize();
                if (n.startsWith(base)) {
                    keepHere.add(n);
                }
            }
            try {
                ResultArchive.archiveOldTrends(dir, keepHere);
            } catch (IOException e) {
                warnings.add("過去月トレンドへの退避に失敗: " + dir + " (" + e.getMessage() + ")");
            }
        }
        Path preferred = xlsx.succeeded().isEmpty() ? null : xlsx.succeeded().get(0);
        return new TrendResult(workbook, xlsx, preferred, stamp, model, warnings);
    }

    static TrendMonthData toDisplay(TrendMonthData src) {
        java.util.LinkedHashMap<String, TrendMetric> kinds = new java.util.LinkedHashMap<>();
        for (var e : src.kinds().entrySet()) {
            kinds.put(e.getKey(), display(e.getValue()));
        }
        java.util.LinkedHashMap<String, TrendMetric> sheets = new java.util.LinkedHashMap<>();
        for (var e : src.sheets().entrySet()) {
            sheets.put(e.getKey(), display(e.getValue()));
        }
        return new TrendMonthData(src.ym(), src.path(), kinds, sheets, src.warnings());
    }

    private static TrendMetric display(TrendMetric m) {
        return new TrendMetric(
                TrendText.roundDisp(m.wage() / 1000.0), TrendText.roundDisp(m.qty() / 1000.0));
    }

    static List<Path> kokubuYearDirs(Map<String, String> ui) {
        String raw = ui == null ? null : ui.get(AppPaths.KEY_PM_AI_KOUCHIN_KOKUBU_YEAR_DIRS);
        if (raw == null || raw.isBlank()) {
            raw = AppPaths.DEFAULT_KOUCHIN_KOKUBU_YEAR_DIRS;
        }
        List<Path> dirs = new ArrayList<>();
        for (String p : raw.split(";")) {
            if (p != null && !p.isBlank()) {
                dirs.add(Path.of(p.trim()));
            }
        }
        return List.copyOf(dirs);
    }
}
