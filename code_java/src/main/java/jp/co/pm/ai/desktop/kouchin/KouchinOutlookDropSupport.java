package jp.co.pm.ai.desktop.kouchin;

import java.io.File;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javafx.scene.input.DataFormat;
import javafx.scene.input.Dragboard;

/**
 * Outlook添付ドロップ。JavaFXの {@code hasFiles()} には載らない。
 * {@code message/external-body} のバイト列を一時CSVにし、
 * 取れないときは TEMP と Content.Outlook の直近 {@code RVSHEET*.csv} を使う。
 */
public final class KouchinOutlookDropSupport {

    static final Duration RECENT_TEMP_MAX_AGE = Duration.ofSeconds(120);
    private static final Pattern RVSHEET_CSV = Pattern.compile("RVSHEET.*\\.CSV", Pattern.CASE_INSENSITIVE);
    private static final Pattern SAFE_NAME = Pattern.compile("[A-Za-z0-9._-]+");
    private static final Pattern NYUKO_YYMMDD =
            Pattern.compile("(?<![0-9])(\\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\\d|3[01])(?![0-9])");
    private static final Pattern MIME_FILENAME =
            Pattern.compile("name\\s*=\\s*\"([^\"]+)\"", Pattern.CASE_INSENSITIVE);

    private KouchinOutlookDropSupport() {}

    public static List<Path> resolveDroppedFiles(Dragboard db) {
        if (db == null) {
            return List.of();
        }
        List<String> originals = originalNamesFromDragboard(db);
        List<Path> files = existingJavaFxFiles(db);
        if (files.isEmpty()) {
            files = existingPathFiles(pathsFromText(db));
        }
        if (files.isEmpty()) {
            files = materializeExternalBody(mimeIds(db), externalBodyContent(db), systemTempDir());
        }
        if (files.isEmpty()) {
            files = recentTempCsvFiles(systemTempDir(), Instant.now());
        }
        if (files.isEmpty()) {
            files = tempFilesNamed(systemTempDir(), originals);
        }
        if (files.isEmpty()) {
            Path cached = newestRecentRvsheetUnder(outlookContentCacheDir(), Instant.now());
            if (cached != null) {
                files = List.of(cached);
            }
        }
        if (files.isEmpty() && !originals.isEmpty()) {
            Path newest = newestTempRvsheetCsv(systemTempDir());
            if (newest != null) {
                files = List.of(newest);
            }
        }
        return applyOriginalName(files, originals);
    }

    /** ドロップ完了後の再取得。Dragboard は触らないので FX スレッド以外から呼べる。 */
    public static List<Path> resolveLateTempFiles(List<String> originals) {
        return resolveLateTempFiles(originals, systemTempDir(), outlookContentCacheDir(), Instant.now());
    }

    static List<Path> resolveLateTempFiles(List<String> originals, Path tempDir, Path cacheDir, Instant now) {
        List<String> names = originals == null ? List.of() : originals;
        Instant at = now == null ? Instant.now() : now;
        List<Path> named = tempFilesNamed(tempDir, names);
        if (!named.isEmpty()) {
            return applyOriginalName(named, names);
        }
        List<Path> recent = recentTempCsvFiles(tempDir, at);
        if (!recent.isEmpty()) {
            return applyOriginalName(recent, names);
        }
        Path cached = newestRecentRvsheetUnder(cacheDir, at);
        if (cached != null) {
            return applyOriginalName(List.of(cached), names);
        }
        if (!names.isEmpty()) {
            Path newest = newestTempRvsheetCsv(tempDir);
            if (newest != null) {
                return applyOriginalName(List.of(newest), names);
            }
        }
        return List.of();
    }

    public static List<String> originalNames(Dragboard db) {
        return List.copyOf(originalNamesFromDragboard(db));
    }

    /** ドロップ完了後に Outlook が TEMP へ書き終わる場合の再取得。 */
    public static List<Path> resolveAfterDropCompleted(Dragboard db) {
        List<Path> first = resolveDroppedFiles(db);
        if (!first.isEmpty()) {
            return first;
        }
        List<String> originals = db == null ? List.of() : originalNamesFromDragboard(db);
        List<Path> named = tempFilesNamed(systemTempDir(), originals);
        Path newest = newestTempRvsheetCsv(systemTempDir());
        if (newest != null) {
            return applyOriginalName(List.of(newest), originals);
        }
        if (!named.isEmpty()) {
            return applyOriginalName(named, originals);
        }
        return applyOriginalName(recentTempCsvFiles(systemTempDir(), Instant.now()), originals);
    }

    public static Path systemTempDir() {
        String raw = System.getProperty("java.io.tmpdir");
        return Paths.get(raw == null || raw.isBlank() ? "." : raw);
    }

    public static List<Path> recentTempCsvFiles(Path tempDir, Instant now) {
        if (tempDir == null || now == null || !Files.isDirectory(tempDir)) {
            return List.of();
        }
        List<Path> found = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(tempDir, p -> {
            Path name = p.getFileName();
            return name != null && RVSHEET_CSV.matcher(name.toString()).matches();
        })) {
            for (Path p : stream) {
                try {
                    if (!Files.isRegularFile(p) || p.getFileName() == null) {
                        continue;
                    }
                    Instant mtime = Files.getLastModifiedTime(p).toInstant();
                    if (Duration.between(mtime, now).compareTo(RECENT_TEMP_MAX_AGE) > 0) {
                        continue;
                    }
                    if (Duration.between(mtime, now).isNegative()
                            && Duration.between(now, mtime).compareTo(Duration.ofSeconds(5)) > 0) {
                        continue;
                    }
                    found.add(p.toAbsolutePath().normalize());
                } catch (IOException ignored) {
                }
            }
        } catch (IOException ignored) {
            return List.of();
        }
        found.sort(Comparator.comparing((Path p) -> datedRvsheet(p.getFileName().toString()) ? 0 : 1)
                .thenComparing(p -> p.getFileName().toString()));
        return List.copyOf(found);
    }

    /**
     * Outlook の {@code message/external-body} バイト列を一時CSVにする。
     * FileGroupDescriptor の構造体はそのままではCSVにしない。
     */
    public static List<Path> materializeExternalBody(Iterable<String> mimeIds, Object content, Path tempDir) {
        byte[] raw = payloadBytes(content);
        if (raw == null || tempDir == null || !fileNamesFromDescriptor(raw).isEmpty()) {
            return List.of();
        }
        List<String> names = fileNamesFromContentTypeIds(mimeIds);
        String name = sanitizeDropFileName(names.isEmpty() ? null : names.get(0));
        try {
            Path dir = tempDir.resolve("pm-ai-outlook-drop");
            Files.createDirectories(dir);
            Path dest = dir.resolve(name);
            Files.write(dest, raw);
            return List.of(dest.toAbsolutePath().normalize());
        } catch (IOException e) {
            return List.of();
        }
    }

    /** Content.Outlook 直下の各フォルダから、更新が新しい {@code RVSHEET*.csv} を1件返す。 */
    static Path newestRecentRvsheetUnder(Path root, Instant now) {
        Path best = null;
        Instant bestTime = Instant.EPOCH;
        for (Path p : recentRvsheetUnder(root, now)) {
            try {
                Instant mt = Files.getLastModifiedTime(p).toInstant();
                if (best == null || mt.isAfter(bestTime)) {
                    best = p;
                    bestTime = mt;
                }
            } catch (IOException ignored) {
            }
        }
        return best == null ? null : best.toAbsolutePath().normalize();
    }

    static List<Path> recentRvsheetUnder(Path root, Instant now) {
        if (root == null || now == null || !Files.isDirectory(root)) {
            return List.of();
        }
        List<Path> found = new ArrayList<>();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(root)) {
            for (Path child : stream) {
                if (Files.isDirectory(child)) {
                    found.addAll(recentTempCsvFiles(child, now));
                }
            }
        } catch (IOException e) {
            return List.of();
        }
        return List.copyOf(found);
    }

    static Path outlookContentCacheDir() {
        String local = System.getenv("LOCALAPPDATA");
        if (local == null || local.isBlank()) {
            return null;
        }
        try {
            return Path.of(local, "Microsoft", "Windows", "INetCache", "Content.Outlook");
        } catch (RuntimeException e) {
            return null;
        }
    }

    static String sanitizeDropFileName(String raw) {
        if (raw == null || raw.isBlank()) {
            return "RVSHEET.csv";
        }
        String name = raw.trim();
        int slash = Math.max(name.lastIndexOf('/'), name.lastIndexOf('\\'));
        if (slash >= 0 && slash + 1 < name.length()) {
            name = name.substring(slash + 1);
        }
        name = name.replaceAll("[\\\\/:*?\"<>|\\p{Cntrl}]", "_").trim();
        if (name.isBlank() || ".".equals(name) || "..".equals(name)) {
            return "RVSHEET.csv";
        }
        return name;
    }

    /** Outlook MIME の name= で指名された TEMP ファイル。更新時刻は問わない。 */
    public static List<Path> tempFilesNamed(Path tempDir, List<String> names) {
        if (tempDir == null || names == null || names.isEmpty() || !Files.isDirectory(tempDir)) {
            return List.of();
        }
        List<Path> found = new ArrayList<>();
        for (String raw : names) {
            if (raw == null || raw.isBlank()) {
                continue;
            }
            String name = Path.of(raw.trim()).getFileName().toString();
            if (name.isBlank() || !SAFE_NAME.matcher(name).matches()) {
                continue;
            }
            Path p = tempDir.resolve(name);
            if (Files.isRegularFile(p)) {
                found.add(p.toAbsolutePath().normalize());
            }
        }
        return List.copyOf(found);
    }

    /** TEMP 上の最新 {@code RVSHEET*.csv}（{@code RVSHEET (1).csv} を含む）。 */
    public static Path newestTempRvsheetCsv(Path tempDir) {
        if (tempDir == null || !Files.isDirectory(tempDir)) {
            return null;
        }
        Path best = null;
        Instant bestTime = Instant.EPOCH;
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(tempDir, p -> {
            Path name = p.getFileName();
            return name != null && RVSHEET_CSV.matcher(name.toString()).matches();
        })) {
            for (Path p : stream) {
                try {
                    if (!Files.isRegularFile(p)) {
                        continue;
                    }
                    Instant mt = Files.getLastModifiedTime(p).toInstant();
                    if (best == null || mt.isAfter(bestTime)) {
                        best = p;
                        bestTime = mt;
                    }
                } catch (IOException ignored) {
                }
            }
        } catch (IOException ignored) {
            return null;
        }
        return best == null ? null : best.toAbsolutePath().normalize();
    }

    public static List<String> fileNamesFromContentTypeIds(Iterable<String> ids) {
        if (ids == null) {
            return List.of();
        }
        List<String> names = new ArrayList<>();
        for (String id : ids) {
            if (id == null || id.isBlank()) {
                continue;
            }
            Matcher m = MIME_FILENAME.matcher(id);
            while (m.find()) {
                String name = Path.of(m.group(1).trim()).getFileName().toString();
                if (!name.isBlank()) {
                    names.add(name);
                }
            }
        }
        return List.copyOf(names);
    }

    public static Path withOriginalFileName(Path src, String originalName) throws IOException {
        if (src == null || !Files.isRegularFile(src)) {
            return src;
        }
        String name = originalName == null ? "" : Path.of(originalName).getFileName().toString().trim();
        if (name.isEmpty() || !SAFE_NAME.matcher(name).matches()) {
            return src;
        }
        if (name.equalsIgnoreCase(src.getFileName().toString())) {
            return src;
        }
        Path dest = src.resolveSibling(name);
        Files.copy(src, dest, StandardCopyOption.REPLACE_EXISTING);
        return dest;
    }

    public static List<String> fileNamesFromDescriptor(byte[] raw) {
        if (raw == null || raw.length < 8) {
            return List.of();
        }
        ByteBuffer buf = ByteBuffer.wrap(raw).order(ByteOrder.LITTLE_ENDIAN);
        int count = buf.getInt();
        if (count <= 0 || count > 32) {
            return List.of();
        }
        int remaining = raw.length - 4;
        if (remaining % count != 0) {
            return List.of();
        }
        int struct = remaining / count;
        if (struct < 80) {
            return List.of();
        }
        List<String> names = new ArrayList<>();
        for (int i = 0; i < count; i++) {
            int start = 4 + i * struct + 72;
            String name = struct >= 500
                    ? utf16z(raw, start, Math.min(raw.length, start + 520))
                    : asciiZ(raw, start, Math.min(raw.length, start + 260));
            if (!name.isBlank()) {
                names.add(name);
            }
        }
        return List.copyOf(names);
    }

    static List<Path> withInferredDatedNames(List<Path> files) {
        return applyOriginalName(files, List.of());
    }

    /** 入庫日（YYMMDD）から {@code RVSHEETyyyyMM.csv} を推定する。 */
    public static String datedRvsheetNameFromCsv(Path src) {
        if (src == null || !Files.isRegularFile(src) || datedRvsheet(src.getFileName().toString())) {
            return null;
        }
        String text;
        try {
            text = new String(Files.readAllBytes(src), java.nio.charset.Charset.forName("windows-31j"));
        } catch (IOException e) {
            return null;
        }
        java.util.Map<String, Integer> counts = new java.util.HashMap<>();
        java.util.regex.Matcher m = NYUKO_YYMMDD.matcher(text);
        int best = 0;
        String bestYm = null;
        while (m.find()) {
            String ym = String.format(Locale.ROOT, "20%s%s", m.group(1), m.group(2));
            int n = counts.merge(ym, 1, Integer::sum);
            if (n > best) {
                best = n;
                bestYm = ym;
            }
        }
        return bestYm == null ? null : "RVSHEET" + bestYm + ".csv";
    }

    private static List<Path> applyOriginalName(List<Path> files, List<String> originals) {
        if (files == null || files.isEmpty()) {
            return List.of();
        }
        List<Path> out = new ArrayList<>();
        for (int i = 0; i < files.size(); i++) {
            Path p = files.get(i);
            String orig = null;
            if (originals != null) {
                if (i < originals.size()) {
                    orig = originals.get(i);
                } else if (files.size() == 1 && originals.size() == 1) {
                    orig = originals.get(0);
                }
            }
            try {
                if (orig != null) {
                    p = withOriginalFileName(p, orig);
                }
                if (p != null && p.getFileName() != null && !datedRvsheet(p.getFileName().toString())) {
                    String inferred = datedRvsheetNameFromCsv(p);
                    if (inferred != null) {
                        p = withOriginalFileName(p, inferred);
                    }
                }
            } catch (IOException ignored) {
            }
            out.add(p);
        }
        return List.copyOf(out);
    }

    private static List<Path> existingJavaFxFiles(Dragboard db) {
        try {
            if (!db.hasFiles() || db.getFiles() == null) {
                return List.of();
            }
            List<Path> out = new ArrayList<>();
            for (File f : db.getFiles()) {
                if (f != null && f.isFile()) {
                    out.add(f.toPath().toAbsolutePath().normalize());
                }
            }
            return out;
        } catch (RuntimeException e) {
            return List.of();
        }
    }

    private static List<Path> existingPathFiles(List<Path> candidates) {
        if (candidates == null || candidates.isEmpty()) {
            return List.of();
        }
        List<Path> out = new ArrayList<>();
        for (Path p : candidates) {
            if (p != null && Files.isRegularFile(p)) {
                out.add(p.toAbsolutePath().normalize());
            }
        }
        return out;
    }

    private static List<Path> pathsFromText(Dragboard db) {
        List<Path> out = new ArrayList<>();
        addPath(out, db.getString());
        if (db.hasUrl()) {
            addPath(out, db.getUrl());
        }
        try {
            Set<DataFormat> types = db.getContentTypes();
            if (types != null) {
                for (DataFormat fmt : types) {
                    Object content = db.getContent(fmt);
                    if (content instanceof String s) {
                        addPath(out, s);
                    }
                }
            }
        } catch (RuntimeException ignored) {
        }
        return out;
    }

    private static List<String> mimeIds(Dragboard db) {
        if (db == null) {
            return List.of();
        }
        List<String> ids = new ArrayList<>();
        try {
            Set<DataFormat> types = db.getContentTypes();
            if (types == null) {
                return List.of();
            }
            for (DataFormat fmt : types) {
                if (fmt.getIdentifiers() != null) {
                    ids.addAll(fmt.getIdentifiers());
                }
            }
        } catch (RuntimeException ignored) {
        }
        return ids;
    }

    /**
     * ファイル本体は短い MIME {@code message/external-body} にある。
     * パラメータ付きの DataFormat へ {@code getContent} しても null になる。
     */
    private static Object externalBodyContent(Dragboard db) {
        Object viaShort = contentOf(db, "message/external-body");
        if (isFilePayload(viaShort)) {
            return viaShort;
        }
        try {
            Set<DataFormat> types = db.getContentTypes();
            if (types == null) {
                return null;
            }
            for (DataFormat fmt : types) {
                if (!isExternalBody(fmt)) {
                    continue;
                }
                Object content = db.getContent(fmt);
                if (isFilePayload(content)) {
                    return content;
                }
            }
        } catch (RuntimeException ignored) {
        }
        return null;
    }

    private static Object contentOf(Dragboard db, String mime) {
        try {
            DataFormat df = DataFormat.lookupMimeType(mime);
            if (df == null) {
                df = new DataFormat(mime);
            }
            return db.getContent(df);
        } catch (RuntimeException e) {
            return null;
        }
    }

    private static boolean isExternalBody(DataFormat fmt) {
        if (fmt == null || fmt.getIdentifiers() == null) {
            return false;
        }
        for (String id : fmt.getIdentifiers()) {
            if (id != null && id.toLowerCase(Locale.ROOT).contains("message/external-body")) {
                return true;
            }
        }
        return false;
    }

    private static boolean isFilePayload(Object content) {
        byte[] raw = payloadBytes(content);
        return raw != null && fileNamesFromDescriptor(raw).isEmpty();
    }

    static byte[] payloadBytes(Object content) {
        if (content instanceof byte[] raw) {
            return raw.length == 0 ? null : raw;
        }
        if (content instanceof ByteBuffer buf) {
            ByteBuffer slice = buf.slice();
            if (!slice.hasRemaining()) {
                return null;
            }
            byte[] raw = new byte[slice.remaining()];
            slice.get(raw);
            return raw;
        }
        return null;
    }

    private static List<String> originalNamesFromDragboard(Dragboard db) {
        if (db == null) {
            return List.of();
        }
        List<String> names = new ArrayList<>();
        try {
            Set<DataFormat> types = db.getContentTypes();
            if (types == null) {
                return List.of();
            }
            for (DataFormat fmt : types) {
                if (fmt.getIdentifiers() != null) {
                    names.addAll(fileNamesFromContentTypeIds(fmt.getIdentifiers()));
                }
                Object content = db.getContent(fmt);
                if (content instanceof byte[] raw) {
                    names.addAll(fileNamesFromDescriptor(raw));
                } else if (content instanceof ByteBuffer buf) {
                    byte[] raw = new byte[buf.remaining()];
                    buf.slice().get(raw);
                    names.addAll(fileNamesFromDescriptor(raw));
                }
            }
        } catch (RuntimeException ignored) {
        }
        return names;
    }

    private static void addPath(List<Path> out, String raw) {
        if (raw == null || raw.isBlank()) {
            return;
        }
        String t = raw.trim().replace("file:///", "").replace("file://", "");
        try {
            Path p = Path.of(t);
            if (Files.isRegularFile(p)) {
                out.add(p.toAbsolutePath().normalize());
            }
        } catch (RuntimeException ignored) {
        }
    }

    private static boolean datedRvsheet(String name) {
        return name != null && name.toUpperCase(Locale.ROOT).matches("RVSHEET\\d{6}\\.CSV");
    }

    private static String utf16z(byte[] raw, int start, int end) {
        int i = start;
        while (i + 1 < end && (raw[i] != 0 || raw[i + 1] != 0)) {
            i += 2;
        }
        return new String(raw, start, i - start, StandardCharsets.UTF_16LE).trim();
    }

    private static String asciiZ(byte[] raw, int start, int end) {
        int i = start;
        while (i < end && raw[i] != 0) {
            i++;
        }
        return new String(raw, start, i - start, StandardCharsets.US_ASCII).trim();
    }
}
