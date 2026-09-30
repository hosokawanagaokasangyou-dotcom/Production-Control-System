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
import java.nio.file.StandardOpenOption;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.fasterxml.jackson.databind.ObjectMapper;

import javafx.scene.input.DataFormat;
import javafx.scene.input.Dragboard;

import jp.co.pm.ai.desktop.debug.AgentDebugLog;

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
    private static final ObjectMapper DEBUG_JSON = new ObjectMapper();
    private static final ThreadLocal<Map<String, Object>> DROP_TRACE = new ThreadLocal<>();

    private KouchinOutlookDropSupport() {}

    public static List<Path> resolveDroppedFiles(Dragboard db) {
        if (db == null) {
            debugDrop("E", "KouchinOutlookDropSupport.resolveDroppedFiles", "null dragboard", Map.of(
                    "thread", Thread.currentThread().getName()));
            return List.of();
        }
        Map<String, Object> trace = new LinkedHashMap<>();
        DROP_TRACE.set(trace);
        try {
            List<String> originals = originalNamesFromDragboard(db);
            boolean hasFiles = false;
            int javafxCount = -1;
            String hasFilesError = "";
            try {
                hasFiles = db.hasFiles();
                javafxCount = db.getFiles() == null ? 0 : db.getFiles().size();
            } catch (RuntimeException ex) {
                hasFilesError = ex.getClass().getSimpleName();
            }
            List<Path> files = existingJavaFxFiles(db);
            String branch = files.isEmpty() ? "" : "javafx";
            if (files.isEmpty()) {
                files = existingPathFiles(pathsFromText(db));
                if (!files.isEmpty()) {
                    branch = "text";
                }
            }
            if (files.isEmpty()) {
                files = materializeExternalBody(mimeIds(db), externalBodyContent(db), systemTempDir());
                if (!files.isEmpty()) {
                    branch = "external-body";
                }
            }
            if (files.isEmpty()) {
                Object clip = clipboardExternalBody();
                notePayload("clipboard", clip);
                files = materializeExternalBody(mimeIds(db), clip, systemTempDir());
                if (!files.isEmpty()) {
                    branch = "clipboard";
                }
            }
            if (files.isEmpty()) {
                OutlookRunningAttachment.Result saved = OutlookRunningAttachment.saveMatch(originals, systemTempDir());
                trace.put("outlookDetail", saved.detail());
                if (saved.path() != null) {
                    files = List.of(saved.path());
                    branch = "outlook-com";
                }
            }
            if (files.isEmpty()) {
                files = recentTempCsvFiles(systemTempDir(), Instant.now());
                if (!files.isEmpty()) {
                    branch = "recent-temp";
                }
            }
            if (files.isEmpty()) {
                files = tempFilesNamed(systemTempDir(), originals);
                if (!files.isEmpty()) {
                    branch = "named-temp";
                }
            }
            if (files.isEmpty()) {
                Path cached = newestRecentRvsheetUnder(outlookContentCacheDir(), Instant.now());
                if (cached != null) {
                    files = List.of(cached);
                    branch = "outlook-cache";
                }
            }
            if (files.isEmpty() && !originals.isEmpty()) {
                Path newest = newestTempRvsheetCsv(systemTempDir());
                if (newest != null) {
                    files = List.of(newest);
                    branch = "newest-temp";
                }
            }
            List<Path> named = applyOriginalName(files, originals);
            trace.put("thread", Thread.currentThread().getName());
            trace.put("hasFiles", hasFiles);
            trace.put("javafxCount", javafxCount);
            trace.put("hasFilesError", hasFilesError);
            trace.put("mimeCount", mimeIds(db).size());
            trace.put("mimes", clip(mimeIds(db), 6));
            trace.put("originalCount", originals.size());
            trace.put("originals", clip(originals, 6));
            trace.put("branch", branch.isEmpty() ? "empty" : branch);
            trace.put("resultCount", named.size());
            trace.put("resultNames", debugNames(named));
            trace.put("tempIsDir", Files.isDirectory(systemTempDir()));
            Path cache = outlookContentCacheDir();
            trace.put("cacheIsDir", cache != null && Files.isDirectory(cache));
            trace.put("contents", contentSummary(db));
            debugDrop("A", "KouchinOutlookDropSupport.resolveDroppedFiles", "resolved", trace);
            return named;
        } finally {
            DROP_TRACE.remove();
        }
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
        boolean descriptor = raw != null && !fileNamesFromDescriptor(raw).isEmpty();
        Map<String, Object> trace = DROP_TRACE.get();
        if (trace != null) {
            trace.put("materializeLen", raw == null ? -1 : raw.length);
            trace.put("materializeDescriptor", descriptor);
            trace.put("materializeHead", headHex(raw));
        }
        if (raw == null || tempDir == null || descriptor) {
            return List.of();
        }
        List<String> names = fileNamesFromContentTypeIds(mimeIds);
        String name = sanitizeDropFileName(names.isEmpty() ? null : names.get(0));
        try {
            Path dir = tempDir.resolve("pm-ai-outlook-drop");
            Files.createDirectories(dir);
            Path dest = dir.resolve(name);
            Files.write(dest, raw);
            if (trace != null) {
                trace.put("materializeName", name);
            }
            return List.of(dest.toAbsolutePath().normalize());
        } catch (IOException e) {
            if (trace != null) {
                trace.put("materializeIo", e.getClass().getSimpleName());
            }
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
    private static Object clipboardExternalBody() {
        try {
            DataFormat df = DataFormat.lookupMimeType("message/external-body");
            if (df == null) {
                df = new DataFormat("message/external-body");
            }
            return javafx.scene.input.Clipboard.getSystemClipboard().getContent(df);
        } catch (RuntimeException ex) {
            Map<String, Object> trace = DROP_TRACE.get();
            if (trace != null) {
                trace.put("clipboardError", ex.getClass().getSimpleName());
            }
            return null;
        }
    }

    private static String contentSummary(Dragboard db) {
        StringBuilder sb = new StringBuilder();
        try {
            Set<DataFormat> types = db.getContentTypes();
            if (types == null) {
                return "";
            }
            for (DataFormat fmt : types) {
                String id = "";
                if (fmt.getIdentifiers() != null && !fmt.getIdentifiers().isEmpty()) {
                    id = fmt.getIdentifiers().iterator().next();
                }
                Object content = null;
                String err = "";
                try {
                    content = db.getContent(fmt);
                } catch (RuntimeException ex) {
                    err = ex.getClass().getSimpleName();
                }
                byte[] raw = payloadBytes(content);
                if (sb.length() > 0) {
                    sb.append(';');
                }
                sb.append(id.length() > 48 ? id.substring(0, 48) : id);
                sb.append('=');
                sb.append(content == null ? "null" : content.getClass().getSimpleName());
                sb.append(':');
                sb.append(raw == null ? -1 : raw.length);
                if (!err.isEmpty()) {
                    sb.append('!').append(err);
                }
                if (sb.length() > 400) {
                    break;
                }
            }
        } catch (RuntimeException ex) {
            return "ex:" + ex.getClass().getSimpleName();
        }
        return sb.toString();
    }

    private static Object externalBodyContent(Dragboard db) {
        Object viaShort = contentOf(db, "message/external-body");
        notePayload("short", viaShort);
        if (isFilePayload(viaShort)) {
            return viaShort;
        }
        try {
            Set<DataFormat> types = db.getContentTypes();
            if (types == null) {
                return null;
            }
            int formats = 0;
            for (DataFormat fmt : types) {
                if (!isExternalBody(fmt)) {
                    continue;
                }
                formats++;
                Object content = db.getContent(fmt);
                notePayload("external", content);
                if (isFilePayload(content)) {
                    Map<String, Object> trace = DROP_TRACE.get();
                    if (trace != null) {
                        trace.put("externalFormats", formats);
                    }
                    return content;
                }
            }
            Map<String, Object> trace = DROP_TRACE.get();
            if (trace != null) {
                trace.put("externalFormats", formats);
            }
        } catch (RuntimeException ex) {
            Map<String, Object> trace = DROP_TRACE.get();
            if (trace != null) {
                trace.put("externalError", ex.getClass().getSimpleName());
            }
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
            Map<String, Object> trace = DROP_TRACE.get();
            if (trace != null) {
                trace.put("contentError", e.getClass().getSimpleName());
            }
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

    static List<String> debugNames(List<Path> files) {
        if (files == null || files.isEmpty()) {
            return List.of();
        }
        List<String> names = new ArrayList<>();
        for (Path p : files) {
            if (p != null && p.getFileName() != null && names.size() < 8) {
                names.add(p.getFileName().toString());
            }
        }
        return List.copyOf(names);
    }

    static void debugDrop(String hypothesisId, String location, String message, Map<String, ?> data) {
        // #region agent log
        try {
            Map<String, Object> payload = new LinkedHashMap<>();
            if (data != null) {
                payload.putAll(data);
            }
            payload.put("runId", "pre-fix");
            AgentDebugLog.appendStructured(Map.of(), "80da7d", hypothesisId, location, message, payload);
            Map<String, Object> line = new LinkedHashMap<>();
            line.put("sessionId", "80da7d");
            line.put("hypothesisId", hypothesisId);
            line.put("location", location);
            line.put("message", message);
            line.put("data", payload);
            line.put("timestamp", System.currentTimeMillis());
            line.put("runId", "pre-fix");
            String json = DEBUG_JSON.writeValueAsString(line);
            Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
            writeDebugLine(cwd.resolve("debug-80da7d.log"), json);
            if (cwd.getParent() != null) {
                writeDebugLine(cwd.getParent().resolve("debug-80da7d.log"), json);
            }
            Path cursor = AgentDebugLog.resolveNdjsonPath(Map.of(), "80da7d");
            if (cursor.getParent() != null && cursor.getParent().getParent() != null) {
                writeDebugLine(cursor.getParent().getParent().resolve("debug-80da7d.log"), json);
            }
        } catch (Throwable ignored) {
        }
        // #endregion
    }

    private static void writeDebugLine(Path file, String json) {
        try {
            if (file.getParent() != null) {
                Files.createDirectories(file.getParent());
            }
            Files.writeString(file, json + System.lineSeparator(), StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE, StandardOpenOption.APPEND);
        } catch (IOException ignored) {
        }
    }

    private static void notePayload(String key, Object content) {
        Map<String, Object> trace = DROP_TRACE.get();
        if (trace == null) {
            return;
        }
        byte[] raw = payloadBytes(content);
        trace.put(key + "Class", content == null ? "null" : content.getClass().getName());
        trace.put(key + "Len", raw == null ? -1 : raw.length);
        trace.put(key + "Descriptor", raw != null && !fileNamesFromDescriptor(raw).isEmpty());
        trace.put(key + "Head", headHex(raw));
    }

    private static String headHex(byte[] raw) {
        if (raw == null || raw.length == 0) {
            return "";
        }
        int n = Math.min(8, raw.length);
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < n; i++) {
            sb.append(String.format(Locale.ROOT, "%02x", raw[i] & 0xff));
        }
        return sb.toString();
    }

    private static List<String> clip(List<String> values, int max) {
        if (values == null || values.isEmpty()) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        for (String v : values) {
            if (v == null || out.size() >= max) {
                continue;
            }
            out.add(v.length() > 180 ? v.substring(0, 180) : v);
        }
        return List.copyOf(out);
    }
}
