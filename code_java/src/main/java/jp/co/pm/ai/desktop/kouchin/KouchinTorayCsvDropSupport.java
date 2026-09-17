package jp.co.pm.ai.desktop.kouchin;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * ①東レCSVのドラッグ＆ドロップ取り込み。ディレクトリ指定は維持し、.csv だけコピーする。
 */
public final class KouchinTorayCsvDropSupport {

    public record Outcome(List<Path> copied, List<String> warnings, List<String> errors) {
        public boolean copiedAny() {
            return !copied.isEmpty();
        }
    }

    private KouchinTorayCsvDropSupport() {}

    public static Outcome copyCsvFiles(
            List<Path> dropped, Path destDir, AtomicBoolean cancel, boolean overwriteExisting) {
        List<Path> copied = new ArrayList<>();
        List<String> warnings = new ArrayList<>();
        List<String> errors = new ArrayList<>();
        if (destDir == null) {
            errors.add("取り込み先フォルダが未設定です");
            return new Outcome(List.of(), List.of(), List.copyOf(errors));
        }
        try {
            Files.createDirectories(destDir);
        } catch (IOException e) {
            errors.add("取り込み先を作成できません: " + destDir + " (" + e.getMessage() + ")");
            return new Outcome(List.of(), List.of(), List.copyOf(errors));
        }
        List<Path> csvs = expandCsv(dropped, warnings);
        for (Path src : csvs) {
            if (cancel != null && cancel.get()) {
                warnings.add("中断のため残りは取り込みませんでした");
                break;
            }
            if (!Files.isRegularFile(src)) {
                continue;
            }
            Path dest = destDir.resolve(src.getFileName().toString());
            if (Files.exists(dest) && !overwriteExisting) {
                warnings.add("上書きせずスキップ: " + src.getFileName());
                continue;
            }
            try {
                Files.copy(src, dest, StandardCopyOption.REPLACE_EXISTING);
                copied.add(dest);
                String name = src.getFileName().toString();
                if (!name.toUpperCase(Locale.ROOT).matches("RVSHEET\\d{6}\\.CSV")) {
                    warnings.add("ファイル名が RVSHEETyyyyMM.csv ではありません（コピーは実施）: " + name);
                }
            } catch (IOException e) {
                errors.add("書込失敗（コピーなし）: " + src.getFileName() + " → " + dest + " (" + e.getMessage() + ")");
            }
        }
        return new Outcome(List.copyOf(copied), List.copyOf(warnings), List.copyOf(errors));
    }

    /** 取り込み先に同名があるCSVファイル名。 */
    public static List<String> existingDestFileNames(List<Path> dropped, Path destDir) {
        if (destDir == null || !Files.isDirectory(destDir)) {
            return List.of();
        }
        List<String> names = new ArrayList<>();
        for (Path src : expandCsv(dropped, new ArrayList<>())) {
            if (src == null || src.getFileName() == null) {
                continue;
            }
            String name = src.getFileName().toString();
            if (Files.exists(destDir.resolve(name))) {
                names.add(name);
            }
        }
        return List.copyOf(names);
    }

    static List<Path> expandCsv(List<Path> dropped, List<String> warnings) {
        List<Path> csvs = new ArrayList<>();
        if (dropped == null) {
            return csvs;
        }
        for (Path p : dropped) {
            if (p == null) {
                continue;
            }
            if (Files.isDirectory(p)) {
                try (var stream = Files.list(p)) {
                    stream.filter(Files::isRegularFile)
                            .filter(f -> f.getFileName() != null)
                            .filter(f -> f.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".csv"))
                            .forEach(csvs::add);
                } catch (IOException e) {
                    errorsSafe(warnings, "フォルダを読めません: " + p + " (" + e.getMessage() + ")");
                }
            } else if (Files.isRegularFile(p)) {
                String n = p.getFileName().toString().toLowerCase(Locale.ROOT);
                if (n.endsWith(".csv")) {
                    csvs.add(p);
                } else {
                    warnings.add("拒否（csv以外）: " + p.getFileName());
                }
            }
        }
        return csvs;
    }

    private static void errorsSafe(List<String> warnings, String msg) {
        warnings.add(msg);
    }
}
