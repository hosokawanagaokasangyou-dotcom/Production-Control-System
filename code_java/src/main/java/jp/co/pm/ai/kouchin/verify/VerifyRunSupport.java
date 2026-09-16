package jp.co.pm.ai.kouchin.verify;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

/**
 * 検証後の Excel / 統合メール / HTML を両工場 ●自動検証へ書き、出力先ごとに archive する。
 */
public final class VerifyRunSupport {

    private static final DateTimeFormatter STAMP = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");

    public record Written(
            DualWriteFiles.WriteOutcome xlsx,
            DualWriteFiles.WriteOutcome mailTxt,
            DualWriteFiles.WriteOutcome mailHtml,
            String stamp,
            String fileBase) {}

    private VerifyRunSupport() {}

    public static Written writeBoth(
            VerifyResult kokubu,
            VerifyResult konan,
            BothResult both,
            Map<String, String> ui,
            AtomicBoolean cancel)
            throws Exception {
        String stamp = LocalDateTime.now().format(STAMP);
        List<Path> dirs = KouchinOutputDirs.resolveAll(ui);
        List<Path> keepAll = new ArrayList<>();
        DualWriteFiles.WriteOutcome xlsxOut = new DualWriteFiles.WriteOutcome(List.of(), List.of());
        DualWriteFiles.WriteOutcome txtOut;
        DualWriteFiles.WriteOutcome htmlOut;

        if (cancel != null && cancel.get()) {
            throw new VerifyException("中断されました");
        }
        if (kokubu != null) {
            String name = "検証結果_国分工場_" + stamp + ".xlsx";
            Workbook wb = ResultExcelExporter.buildWorkbook(kokubu, peerMail(konan, ui, FactoryId.KONAN));
            try {
                xlsxOut = DualWriteFiles.writeWorkbook(wb, KouchinOutputDirs.withNames(dirs, name), ui);
                keepAll.addAll(xlsxOut.succeeded());
            } finally {
                wb.close();
            }
        }
        if (cancel != null && cancel.get()) {
            throw new VerifyException("中断されました");
        }
        if (konan != null) {
            String name = "検証結果_湖南工場_" + stamp + ".xlsx";
            Workbook wb = ResultExcelExporter.buildWorkbook(konan, peerMail(kokubu, ui, FactoryId.KOKUBU));
            try {
                DualWriteFiles.WriteOutcome more =
                        DualWriteFiles.writeWorkbook(wb, KouchinOutputDirs.withNames(dirs, name), ui);
                xlsxOut = merge(xlsxOut, more);
                keepAll.addAll(more.succeeded());
            } finally {
                wb.close();
            }
        }

        MailSnapshot kMail = kokubu == null ? null : kokubu.mail();
        MailSnapshot nMail = konan == null ? null : konan.mail();
        String mail = both != null ? both.unifiedMail() : UnifiedMailBuilder.buildText(kMail, nMail);
        String html = MailHtmlBuilder.build(kMail, nMail);
        String mailName = "報告メール_統合_" + stamp;
        txtOut = DualWriteFiles.writeBytes(
                mail.getBytes(StandardCharsets.UTF_8),
                KouchinOutputDirs.withNames(dirs, mailName + ".txt"));
        htmlOut = DualWriteFiles.writeBytes(
                html.getBytes(StandardCharsets.UTF_8),
                KouchinOutputDirs.withNames(dirs, mailName + ".html"));
        keepAll.addAll(txtOut.succeeded());
        keepAll.addAll(htmlOut.succeeded());

        for (Path dir : dirs) {
            List<Path> keepHere = new ArrayList<>();
            for (Path p : keepAll) {
                if (p != null && p.getParent() != null && p.getParent().equals(dir.toAbsolutePath().normalize())) {
                    keepHere.add(p);
                } else if (p != null && p.startsWith(dir.toAbsolutePath().normalize())) {
                    keepHere.add(p);
                }
            }
            ResultArchive.archiveOldVerifyResults(dir, keepHere);
        }
        return new Written(xlsxOut, txtOut, htmlOut, stamp, mailName);
    }

    public static Path preferredOpenExcel(Written written, FactorySite site) {
        List<Path> ok = written.xlsx() == null ? List.of() : written.xlsx().succeeded();
        if (ok.isEmpty()) {
            return null;
        }
        String prefer = site == FactorySite.KOKUBU ? "国分工場" : "湖南工場";
        for (Path p : ok) {
            if (p.getFileName().toString().contains(prefer)) {
                return p;
            }
        }
        return ok.get(0);
    }

    public static Path preferredOpenDir(Written written, FactorySite site) {
        Path excel = preferredOpenExcel(written, site);
        if (excel != null && excel.getParent() != null) {
            return excel.getParent();
        }
        List<Path> dirs = KouchinOutputDirs.resolveAll();
        if (site == FactorySite.KOKUBU) {
            return dirs.get(0);
        }
        return dirs.size() > 1 ? dirs.get(1) : dirs.get(0);
    }

    private static MailSnapshot peerMail(VerifyResult peer, Map<String, String> ui, FactoryId want) {
        if (peer != null) {
            return peer.mail();
        }
        return LastResultCache.load(want).orElse(null);
    }

    private static DualWriteFiles.WriteOutcome merge(
            DualWriteFiles.WriteOutcome a, DualWriteFiles.WriteOutcome b) {
        List<Path> ok = new ArrayList<>(a.succeeded());
        ok.addAll(b.succeeded());
        List<String> fail = new ArrayList<>(a.failures());
        fail.addAll(b.failures());
        return new DualWriteFiles.WriteOutcome(List.copyOf(ok), List.copyOf(fail));
    }
}
