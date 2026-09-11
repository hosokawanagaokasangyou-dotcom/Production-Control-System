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
        void onProgress(int processed, int total);
    }

    private InspectionSheetIndexScanner() {}

    public static Result scan(Path root, List<InspectionSheetIndexStore.Row> previous, Progress progress)
            throws IOException {
        List<String> warnings = new ArrayList<>();
        if (root == null || !Files.isDirectory(root)) {
            return new Result(List.of(), List.of("検査表フォルダにアクセスできません: " + root), 0);
        }
        List<Path> files = listExcelFiles(root);
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
        for (Path file : files) {
            processed++;
            if (progress != null) {
                progress.onProgress(processed, files.size());
            }
            try {
                String abs = file.toAbsolutePath().normalize().toString();
                long mtime = Files.getLastModifiedTime(file).toMillis();
                long size = Files.size(file);
                InspectionSheetIndexStore.Row prev = prevByPath.get(abs);
                if (prev != null && prev.fileMtimeEpoch() == mtime && prev.fileSize() == size) {
                    out.add(prev);
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
            } catch (Exception ex) {
                warnings.add("読込エラー " + file.getFileName() + ": " + ex.getMessage());
            }
        }
        return new Result(List.copyOf(out), List.copyOf(warnings), readExcel);
    }

    static List<Path> listExcelFiles(Path root) throws IOException {
        List<Path> files = new ArrayList<>();
        try (Stream<Path> walk = Files.walk(root)) {
            walk.filter(Files::isRegularFile)
                    .filter(InspectionSheetIndexScanner::isInspectionExcel)
                    .forEach(files::add);
        }
        files.sort(Path::compareTo);
        return files;
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
