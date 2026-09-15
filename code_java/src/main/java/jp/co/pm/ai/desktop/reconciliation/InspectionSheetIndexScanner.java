package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.Instant;
import java.time.LocalDate;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

/** 検査表フォルダを再帰走査し CSV 索引を増分更新する。 */
public final class InspectionSheetIndexScanner {

    static final int COOPERATIVE_YIELD_EVERY = 256;
    static final int CHECKPOINT_EVERY_FILES = 128;

    /** Box / UNC の待ち時間を隠すため、CPU 数より多めの I/O 並列。 */
    static final int PARALLELISM =
            Math.max(4, Math.min(12, Runtime.getRuntime().availableProcessors() * 2));

    public record Result(List<InspectionSheetIndexStore.Row> rows, List<String> warnings, int readExcelCount) {}

    @FunctionalInterface
    public interface Progress {
        void onProgress(String phase, int processed, int total);
    }

    @FunctionalInterface
    public interface Checkpoint {
        void onRows(List<InspectionSheetIndexStore.Row> rowsSoFar, int readExcelCount);
    }

    private record Item(InspectionSheetIndexStore.Row row, String warning, boolean excelRead) {}

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
        int readExcel = 0;
        String indexedAt = Instant.now().toString();
        int total = files.size();
        notifyProgress(progress, InspectionSheetIndexProgress.PHASE_INDEX, 0, total);
        ExecutorService pool = newIndexPool(Math.min(PARALLELISM, Math.max(1, total)));
        ArrayDeque<CompletableFuture<Item>> pending = new ArrayDeque<>();
        RuntimeException abort = null;
        try {
            for (int i = 0; i < files.size(); i++) {
                int processed = i + 1;
                notifyProgress(progress, InspectionSheetIndexProgress.PHASE_INDEX, processed, total);
                maybeYield(processed);
                Path file = files.get(i);
                pending.add(
                        CompletableFuture.supplyAsync(
                                () -> processOne(file, prevByPath, indexedAt), pool));
                while (pending.size() >= PARALLELISM) {
                    Item item = take(pending);
                    readExcel += applyItem(item, out, warnings);
                    maybeCheckpoint(checkpoint, out, readExcel, processed, total, item.excelRead);
                }
            }
        } catch (RuntimeException ex) {
            abort = ex;
        } finally {
            try {
                while (!pending.isEmpty()) {
                    Item item = take(pending);
                    readExcel += applyItem(item, out, warnings);
                    maybeCheckpoint(checkpoint, out, readExcel, out.size(), total, item.excelRead);
                }
            } finally {
                pool.shutdownNow();
            }
        }
        if (abort != null) {
            notifyCheckpoint(checkpoint, out, readExcel);
            throw abort;
        }
        notifyCheckpoint(checkpoint, out, readExcel);
        return new Result(List.copyOf(out), List.copyOf(warnings), readExcel);
    }

    static List<Path> listExcelFiles(Path root) throws IOException {
        return listExcelFiles(root, null);
    }

    static List<Path> listExcelFiles(Path root, Progress progress) throws IOException {
        List<Path> subDirs = new ArrayList<>();
        List<Path> files = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(root)) {
            for (Path path : stream) {
                if (Files.isDirectory(path)) {
                    subDirs.add(path);
                } else if (isInspectionExcel(path)) {
                    files.add(path);
                }
            }
        }
        if (subDirs.isEmpty()) {
            files.sort(Path::compareTo);
            for (int i = 0; i < files.size(); i++) {
                notifyProgress(progress, InspectionSheetIndexProgress.PHASE_WALK, i + 1, 0);
            }
            return files;
        }
        ConcurrentLinkedQueue<Path> acc = new ConcurrentLinkedQueue<>(files);
        AtomicInteger found = new AtomicInteger(files.size());
        if (found.get() > 0) {
            notifyProgress(progress, InspectionSheetIndexProgress.PHASE_WALK, found.get(), 0);
        }
        int workers = Math.min(PARALLELISM, subDirs.size());
        ExecutorService pool = newIndexPool(workers);
        try {
            List<Future<?>> jobs = new ArrayList<>(subDirs.size());
            for (Path dir : subDirs) {
                jobs.add(
                        pool.submit(
                                () -> {
                                    try (Stream<Path> walk = Files.walk(dir)) {
                                        walk.filter(InspectionSheetIndexScanner::isInspectionExcel)
                                                .filter(Files::isRegularFile)
                                                .forEach(
                                                        path -> {
                                                            acc.add(path);
                                                            notifyProgress(
                                                                    progress,
                                                                    InspectionSheetIndexProgress
                                                                            .PHASE_WALK,
                                                                    found.incrementAndGet(),
                                                                    0);
                                                        });
                                    } catch (IOException ex) {
                                        throw new UncheckedIOException(ex);
                                    }
                                }));
            }
            for (Future<?> job : jobs) {
                try {
                    job.get();
                } catch (InterruptedException ex) {
                    Thread.currentThread().interrupt();
                    throw new IOException("検査表フォルダ走査が中断されました", ex);
                } catch (ExecutionException ex) {
                    Throwable cause = ex.getCause() != null ? ex.getCause() : ex;
                    if (cause instanceof UncheckedIOException uio) {
                        throw uio.getCause();
                    }
                    if (cause instanceof IOException io) {
                        throw io;
                    }
                    throw new IOException(cause.getMessage(), cause);
                }
            }
        } finally {
            pool.shutdownNow();
        }
        List<Path> out = new ArrayList<>(acc);
        out.sort(Path::compareTo);
        return out;
    }

    private static Item processOne(
            Path file, Map<String, InspectionSheetIndexStore.Row> prevByPath, String indexedAt) {
        try {
            String abs = file.toAbsolutePath().normalize().toString();
            BasicFileAttributes attrs = Files.readAttributes(file, BasicFileAttributes.class);
            long mtime = attrs.lastModifiedTime().toMillis();
            long size = attrs.size();
            InspectionSheetIndexStore.Row prev = prevByPath.get(abs);
            if (prev != null && prev.fileMtimeEpoch() == mtime && prev.fileSize() == size) {
                return new Item(prev, null, false);
            }
            InspectionSheetHeaderReader.Header header = InspectionSheetHeaderReader.read(file);
            String irai = header.iraiNo() != null ? header.iraiNo() : "";
            if (irai.isBlank()) {
                irai =
                        InspectionSheetIraiNo.extractFromFileName(file.getFileName().toString())
                                .orElse("");
            }
            if (irai.isBlank()) {
                return new Item(null, "依頼NOを読めません: " + file.getFileName(), true);
            }
            LocalDate date = header.processingDate();
            return new Item(
                    new InspectionSheetIndexStore.Row(
                            irai,
                            date,
                            InspectionSheetProcessingDate.yearMonth(date),
                            abs,
                            file.getFileName().toString(),
                            mtime,
                            size,
                            indexedAt),
                    null,
                    true);
        } catch (Exception ex) {
            return new Item(null, "読込エラー " + file.getFileName() + ": " + ex.getMessage(), false);
        }
    }

    private static Item take(ArrayDeque<CompletableFuture<Item>> pending) {
        return pending.removeFirst().join();
    }

    private static int applyItem(Item item, List<InspectionSheetIndexStore.Row> out, List<String> warnings) {
        if (item.warning() != null) {
            warnings.add(item.warning());
        }
        if (item.row() != null) {
            out.add(item.row());
        }
        return item.excelRead() ? 1 : 0;
    }

    private static void maybeCheckpoint(
            Checkpoint checkpoint,
            List<InspectionSheetIndexStore.Row> out,
            int readExcel,
            int processed,
            int total,
            boolean excelRead) {
        if (excelRead || processed % CHECKPOINT_EVERY_FILES == 0 || processed == total) {
            notifyCheckpoint(checkpoint, out, readExcel);
        }
    }

    private static ExecutorService newIndexPool(int n) {
        int size = Math.max(1, n);
        ThreadFactory factory =
                r -> {
                    Thread t = new Thread(r, "inspection-sheet-index-worker");
                    t.setDaemon(true);
                    t.setPriority(Thread.NORM_PRIORITY);
                    return t;
                };
        return Executors.newFixedThreadPool(size, factory);
    }

    private static void notifyProgress(Progress progress, String phase, int processed, int total) {
        if (progress == null) {
            return;
        }
        progress.onProgress(phase, processed, total);
    }

    private static void maybeYield(int count) {
        if (count > 0 && count % COOPERATIVE_YIELD_EVERY == 0) {
            Thread.yield();
        }
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
