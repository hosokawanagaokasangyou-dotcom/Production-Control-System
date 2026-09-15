package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.stream.Stream;

/** 検査表フォルダを再帰走査し CSV 索引を増分更新する。 */
public final class InspectionSheetIndexScanner {

    public record Result(List<InspectionSheetIndexStore.Row> rows, List<String> warnings, int readExcelCount) {}

    @FunctionalInterface
    public interface Progress {
        void onProgress(String phase, int processed, int total);
    }

    @FunctionalInterface
    public interface Checkpoint {
        void onRows(List<InspectionSheetIndexStore.Row> rowsSoFar, int readExcelCount);
    }

    private InspectionSheetIndexScanner() {}

    public static Result scan(Path root, List<InspectionSheetIndexStore.Row> previous, Progress progress)
            throws IOException {
        return scan(root == null ? List.of() : List.of(root), previous, progress);
    }

    public static Result scan(
            List<Path> roots, List<InspectionSheetIndexStore.Row> previous, Progress progress)
            throws IOException {
        return scan(roots, previous, progress, null);
    }

    public static Result scan(
            List<Path> roots,
            List<InspectionSheetIndexStore.Row> previous,
            Progress progress,
            Checkpoint checkpoint)
            throws IOException {
        List<String> warnings = new ArrayList<>();
        List<Path> usable = new ArrayList<>();
        if (roots != null) {
            for (Path root : roots) {
                if (root != null && Files.isDirectory(root)) {
                    usable.add(root);
                }
            }
        }
        if (usable.isEmpty()) {
            return new Result(List.of(), List.of("検査表フォルダにアクセスできません: " + roots), 0);
        }
        notifyProgress(progress, InspectionSheetIndexProgress.PHASE_WALK, 0, 0);
        List<Path> files = new ArrayList<>();
        java.util.LinkedHashSet<String> seen = new java.util.LinkedHashSet<>();
        for (Path root : usable) {
            for (Path file : listExcelFiles(root, progress)) {
                String abs = file.toAbsolutePath().normalize().toString();
                if (seen.add(abs)) {
                    files.add(file);
                }
            }
        }
        Map<String, InspectionSheetIndexStore.Row> prevByPath = new HashMap<>();
        if (previous != null) {
            for (InspectionSheetIndexStore.Row row : previous) {
                if (row != null && row.filePath() != null && !row.filePath().isBlank()) {
                    prevByPath.put(row.filePath(), row);
                }
            }
        }
        List<InspectionSheetIndexStore.Row> out = new ArrayList<>();
        int processed = 0;
        int readExcel = 0;
        String indexedAt = Instant.now().toString();
        int total = files.size();
        notifyProgress(progress, InspectionSheetIndexProgress.PHASE_INDEX, 0, total);
        for (Path file : files) {
            processed++;
            notifyProgress(progress, InspectionSheetIndexProgress.PHASE_INDEX, processed, total);
            try {
                String abs = file.toAbsolutePath().normalize().toString();
                long mtime = Files.getLastModifiedTime(file).toMillis();
                long size = Files.size(file);
                InspectionSheetIndexStore.Row prev = prevByPath.get(abs);
                if (prev != null && prev.fileMtimeEpoch() == mtime && prev.fileSize() == size) {
                    out.add(prev);
                    notifyCheckpoint(checkpoint, out, readExcel);
                    continue;
                }
                InspectionSheetHeaderReader.Header header = InspectionSheetHeaderReader.read(file);
                readExcel++;
                LocalDate date = header.processingDate();
                String irai = header.iraiNo() != null ? header.iraiNo() : "";
                if (irai.isBlank()) {
                    irai = InspectionSheetIraiNo.extractFromFileName(file.getFileName().toString()).orElse("");
                }
                if (irai.isBlank()) {
                    warnings.add("依頼NOを読めません: " + file.getFileName());
                    continue;
                }
                out.add(
                        new InspectionSheetIndexStore.Row(
                                irai,
                                date,
                                InspectionSheetProcessingDate.yearMonth(date),
                                abs,
                                file.getFileName().toString(),
                                mtime,
                                size,
                                indexedAt));
                notifyCheckpoint(checkpoint, out, readExcel);
            } catch (Exception ex) {
                warnings.add("読込エラー " + file.getFileName() + ": " + ex.getMessage());
            }
        }
        return new Result(List.copyOf(out), List.copyOf(warnings), readExcel);
    }

    static List<Path> listExcelFiles(Path root) throws IOException {
        return listExcelFiles(root, null);
    }

    static List<Path> listExcelFiles(Path root, Progress progress) throws IOException {
        List<Path> files = new ArrayList<>();
        try (Stream<Path> walk = Files.walk(root)) {
            walk.filter(Files::isRegularFile)
                    .filter(InspectionSheetIndexScanner::isInspectionExcel)
                    .forEach(
                            path -> {
                                files.add(path);
                                notifyProgress(
                                        progress,
                                        InspectionSheetIndexProgress.PHASE_WALK,
                                        files.size(),
                                        0);
                            });
        }
        files.sort(Path::compareTo);
        return files;
    }

    private static void notifyProgress(Progress progress, String phase, int processed, int total) {
        if (progress == null) {
            return;
        }
        progress.onProgress(phase, processed, total);
    }

    private static void notifyCheckpoint(
            Checkpoint checkpoint, List<InspectionSheetIndexStore.Row> rows, int readExcelCount) {
        if (checkpoint == null) {
            return;
        }
        checkpoint.onRows(List.copyOf(rows), readExcelCount);
    }

    static boolean isInspectionExcel(Path path) {
        if (path == null) {
            return false;
        }
        Path namePath = path.getFileName();
        if (namePath == null) {
            return false;
        }
        String name = namePath.toString();
        if (name.startsWith("~$")) {
            return false;
        }
        String lower = name.toLowerCase(Locale.ROOT);
        return lower.endsWith(".xlsx") || lower.endsWith(".xlsm");
    }
}
