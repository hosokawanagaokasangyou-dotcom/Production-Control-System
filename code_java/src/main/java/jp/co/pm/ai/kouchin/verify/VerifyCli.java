package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

/**
 * JavaFX なしで検証パイプラインを実行する CLI 入口（スモークテスト・自動実行用）。
 *
 * <pre>
 * mvn -q validate compile exec:java@kouchin-cli -Dexec.args="--factory both"
 * </pre>
 */
public final class VerifyCli {

    private static final DateTimeFormatter STAMP = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");

    private enum FactoryMode {
        KOKUBU, KONAN, BOTH
    }

    private VerifyCli() {
    }

    public static void main(String[] args) {
        int exitCode;
        try {
            CliOptions options = CliOptions.parse(args);
            exitCode = run(options);
        } catch (IllegalArgumentException e) {
            System.err.println("引数エラー: " + e.getMessage());
            printUsage();
            exitCode = 1;
        } catch (Exception e) {
            System.err.println("失敗: " + (e.getMessage() == null ? e.toString() : e.getMessage()));
            e.printStackTrace(System.err);
            exitCode = 1;
        }
        System.exit(exitCode);
    }

    private static int run(CliOptions options) throws IOException {
        KouchinPaths paths = KouchinPaths.defaults();
        ensureOutputDir(paths.outputDir());

        System.out.println("後加工工賃検証 CLI");
        System.out.println("工場: " + options.factoryLabel());
        System.out.println("許容差: ±" + options.tolerance() + " 円");
        System.out.println("出力先: " + paths.outputDir());
        System.out.println();

        List<Path> outputs = new ArrayList<>();

        switch (options.factory()) {
            case KOKUBU -> outputs.add(runOne(FactoryId.KOKUBU, paths, options.tolerance()));
            case KONAN -> outputs.add(runOne(FactoryId.KONAN, paths, options.tolerance()));
            case BOTH -> outputs.addAll(runBoth(paths, options.tolerance()));
        }

        if (outputs.isEmpty()) {
            System.out.println("出力ファイルはありません。");
            return 1;
        }

        System.out.println();
        System.out.println("完了: " + outputs.size() + " ファイル");
        for (Path out : outputs) {
            System.out.println("  " + out);
        }
        return 0;
    }

    private static Path runOne(FactoryId factory, KouchinPaths paths, double tol) throws IOException {
        System.out.println("開始: " + factoryLabel(factory));
        VerifyResult result = VerifyService.run(factory, paths, tol);
        Path out = ResultExcelExporter.defaultOutputPath(paths, result.profile());
        ResultExcelExporter.export(result, out);
        printSummary(result, out);
        return out;
    }

    private static List<Path> runBoth(KouchinPaths paths, double tol) throws IOException {
        System.out.println("開始: まとめて生成（国分 → 湖南）");
        BothResult both = VerifyService.runBoth(paths, tol);
        List<Path> outputs = new ArrayList<>();

        MailSnapshots snaps = new MailSnapshots(
                both.kokubu() != null ? both.kokubu().mail() : null,
                both.konan() != null ? both.konan().mail() : null);

        if (both.kokubu() != null) {
            Path out = ResultExcelExporter.defaultOutputPath(paths, both.kokubu().profile());
            ResultExcelExporter.export(both.kokubu(), out, snaps.kokubu(), snaps.konan());
            outputs.add(out);
            printSummary(both.kokubu(), out);
        } else {
            System.out.println("国分失敗: " + both.kokubuError());
        }

        if (both.konan() != null) {
            Path out = ResultExcelExporter.defaultOutputPath(paths, both.konan().profile());
            ResultExcelExporter.export(both.konan(), out, snaps.kokubu(), snaps.konan());
            outputs.add(out);
            printSummary(both.konan(), out);
        } else {
            System.out.println("湖南失敗: " + both.konanError());
        }

        Path mailTxt = paths.outputDir().resolve("報告メール_統合_" + LocalDateTime.now().format(STAMP) + ".txt");
        Files.writeString(mailTxt, both.unifiedMail() + System.lineSeparator(), StandardCharsets.UTF_8);
        outputs.add(mailTxt);
        System.out.println("統合メール: " + mailTxt);

        return outputs;
    }

    private static void printSummary(VerifyResult result, Path out) {
        System.out.println("出力: " + out);
        System.out.println(summaryLine(result));
    }

    private static String summaryLine(VerifyResult r) {
        return r.profile().label()
                + " A要確認=" + r.requiredCheckA()
                + " B要確認=" + r.requiredCheckB()
                + " 警告=" + r.warnings().size();
    }

    private static String factoryLabel(FactoryId factory) {
        return factory == FactoryId.KOKUBU ? "国分工場" : "湖南工場";
    }

    private static void ensureOutputDir(Path dir) throws IOException {
        if (!Files.isDirectory(dir)) {
            Files.createDirectories(dir);
        }
    }

    private static void printUsage() {
        System.err.println();
        System.err.println("使い方:");
        System.err.println("  --factory kokubu|konan|both  対象工場（既定: both）");
        System.err.println("  --tol <数値>                   許容差（円、既定: "
                + VerifyService.DEFAULT_TOLERANCE + "）");
    }

    private record MailSnapshots(MailSnapshot kokubu, MailSnapshot konan) {
    }

    private record CliOptions(FactoryMode factory, double tolerance) {

        String factoryLabel() {
            return switch (factory) {
                case KOKUBU -> "国分工場";
                case KONAN -> "湖南工場";
                case BOTH -> "両工場（まとめて生成）";
            };
        }

        static CliOptions parse(String[] args) {
            FactoryMode factory = FactoryMode.BOTH;
            double tol = VerifyService.DEFAULT_TOLERANCE;

            for (int i = 0; i < args.length; i++) {
                String arg = args[i];
                switch (arg) {
                    case "--factory" -> {
                        if (++i >= args.length) {
                            throw new IllegalArgumentException("--factory の値がありません");
                        }
                        factory = parseFactory(args[i]);
                    }
                    case "--tol" -> {
                        if (++i >= args.length) {
                            throw new IllegalArgumentException("--tol の値がありません");
                        }
                        try {
                            tol = Double.parseDouble(args[i]);
                        } catch (NumberFormatException e) {
                            throw new IllegalArgumentException("--tol は数値で指定してください: " + args[i]);
                        }
                    }
                    case "--help", "-h" -> {
                        printUsage();
                        System.exit(0);
                    }
                    default -> throw new IllegalArgumentException("不明な引数: " + arg);
                }
            }
            return new CliOptions(factory, tol);
        }

        private static FactoryMode parseFactory(String value) {
            return switch (value.toLowerCase()) {
                case "kokubu" -> FactoryMode.KOKUBU;
                case "konan" -> FactoryMode.KONAN;
                case "both" -> FactoryMode.BOTH;
                default -> throw new IllegalArgumentException(
                        "--factory は kokubu / konan / both のいずれか: " + value);
            };
        }
    }
}
