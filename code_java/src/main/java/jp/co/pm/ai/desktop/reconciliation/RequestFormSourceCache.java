package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.time.Duration;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 依頼書原本の解析結果とプレビュー PDF を {@code preview_cache/} 以下に保持する。
 *
 * <p>各キャッシュ JSON に {@code schemaVersion} を記録し、{@link #PARSE_SCHEMA_VERSION} /
 * {@link #TPI_PDF_PARSE_SCHEMA_VERSION} / {@link #PREVIEW_SCHEMA_VERSION} と一致しないもの、
 * および {@link #CACHE_MAX_AGE} を超えたものは読込時に破棄し、{@link #pruneStaleDiskCaches} で一括削除する。
 */
final class RequestFormSourceCache {

    private static final ObjectMapper JSON = new ObjectMapper();

    /** 依頼書キャッシュの最大保持期間（1か月＝30日）。 */
    static final Duration CACHE_MAX_AGE = Duration.ofDays(30);
    static final long CACHE_MAX_AGE_MILLIS = CACHE_MAX_AGE.toMillis();

    /**
     * {@link RequestFormOriginalCellLayout} ベースのセル読取に合わせて上げる。
     * 抽出ロジック変更時は必ずインクリメントし、古い parse キャッシュを無効化する。
     */
    static final String PARSE_SCHEMA_VERSION = "request-form-cell-layout-v7";

    /** TPI 依頼書 PDF 用 parse キャッシュ schema（Excel 原本とは別バージョン）。 */
    static final String TPI_PDF_PARSE_SCHEMA_VERSION = "request-form-tpi-pdf-v19";

    /** Excel 原本シート PDF プレビュー用 schema。レンダラ・範囲変更時に上げる。 */
    static final String PREVIEW_SCHEMA_VERSION = "request-form-preview-v2";

    private RequestFormSourceCache() {}

    record SourceFingerprint(long lastModified, long length) {}

    private record ParseCachePayload(
            SourceFingerprint source,
            String schemaVersion,
            long cachedAtMillis,
            List<Map<String, String>> entries) {}

    private record PreviewMeta(
            SourceFingerprint source,
            String range,
            String renderer,
            String schemaVersion,
            long cachedAtMillis) {}

    private record SplitCacheMeta(
            SourceFingerprint source,
            int startPage,
            int endPage,
            String schemaVersion,
            long cachedAtMillis) {}

    /** {@link #clearAllDiskCache(File)} の結果。 */
    record ClearDiskCacheResult(int pdfFilesDeleted, int parseFilesDeleted, int deleteFailures) {
        int totalDeleted() {
            return pdfFilesDeleted + parseFilesDeleted;
        }
    }

    static File cacheRoot(File repoPreviewCacheDir) {
        return repoPreviewCacheDir;
    }

    static File parseDir(File cacheRoot) {
        File dir = new File(cacheRoot, "parse");
        if (!dir.exists()) {
            dir.mkdirs();
        }
        return dir;
    }

    static File pdfDir(File cacheRoot) {
        File dir = new File(cacheRoot, "pdf");
        if (!dir.exists()) {
            dir.mkdirs();
        }
        return dir;
    }

    static File splitDir(File cacheRoot) {
        File dir = new File(cacheRoot, "tpi-split");
        if (!dir.exists()) {
            dir.mkdirs();
        }
        return dir;
    }

    static File splitCacheFile(File cacheRoot, File sourcePdf, String iraiNo) {
        String sourceBase = safeBaseName(sourcePdf != null ? sourcePdf.getName() : "unknown");
        String iraiBase = safeBaseName(iraiNo != null ? iraiNo : "unknown");
        return new File(splitDir(cacheRoot), sourceBase + "__" + iraiBase + ".pdf");
    }

    static File splitMetaFile(File splitPdf) {
        return new File(splitPdf.getAbsolutePath() + ".meta.json");
    }

    static boolean isSplitCacheValid(
            File splitPdf, File sourcePdf, int startPage0, int endPage0) {
        if (splitPdf == null
                || !splitPdf.isFile()
                || splitPdf.length() < 128
                || !looksLikePdf(splitPdf)) {
            return false;
        }
        File metaFile = splitMetaFile(splitPdf);
        if (!metaFile.isFile()) {
            return false;
        }
        try {
            SplitCacheMeta meta = JSON.readValue(metaFile, SplitCacheMeta.class);
            return meta != null
                    && matches(sourcePdf, meta.source())
                    && meta.startPage() == startPage0
                    && meta.endPage() == endPage0
                    && TPI_PDF_PARSE_SCHEMA_VERSION.equals(meta.schemaVersion())
                    && !isCacheExpired(meta.cachedAtMillis());
        } catch (IOException ex) {
            return false;
        }
    }

    static void writeSplitCacheMeta(
            File splitPdf, File sourcePdf, int startPage0, int endPage0) throws IOException {
        Files.createDirectories(splitMetaFile(splitPdf).getParentFile().toPath());
        JSON.writeValue(
                splitMetaFile(splitPdf),
                new SplitCacheMeta(
                        fingerprint(sourcePdf),
                        startPage0,
                        endPage0,
                        TPI_PDF_PARSE_SCHEMA_VERSION,
                        nowMillis()));
    }

    /** @deprecated {@link #pdfCacheFile(File, String, String)} を使用 */
    @Deprecated
    static File pngCacheFile(File cacheRoot, String workbookName, String sheetName) {
        return pdfCacheFile(cacheRoot, workbookName, sheetName);
    }

    static File pdfCacheFile(File cacheRoot, String workbookName, String sheetName) {
        String cacheName =
                (workbookName + "_" + sheetName).replaceAll("[\\\\/:*?\"<>|]", "_") + ".pdf";
        return new File(pdfDir(cacheRoot), cacheName);
    }

    static File previewMetaFile(File previewFile) {
        return new File(previewFile.getAbsolutePath() + ".meta.json");
    }

    /** @deprecated {@link #previewMetaFile(File)} を使用 */
    @Deprecated
    static File pngMetaFile(File pngFile) {
        return previewMetaFile(pngFile);
    }

    static boolean isPreviewCacheValid(File pdfFile, File sourceFile) {
        if (pdfFile == null || !pdfFile.isFile() || pdfFile.length() < 128) {
            return false;
        }
        if (!looksLikePdf(pdfFile)) {
            return false;
        }
        File metaFile = previewMetaFile(pdfFile);
        if (!metaFile.isFile()) {
            return false;
        }
        try {
            PreviewMeta meta = JSON.readValue(metaFile, PreviewMeta.class);
            if (meta == null
                    || !matches(sourceFile, meta.source())
                    || !RequestFormSheetPreviewRenderer.PREVIEW_RANGE_SPEC.equals(meta.range())
                    || !RequestFormSheetPreviewPdfRenderer.rendererSpec().equals(meta.renderer())
                    || !PREVIEW_SCHEMA_VERSION.equals(meta.schemaVersion())
                    || isCacheExpired(meta.cachedAtMillis())) {
                return false;
            }
            return true;
        } catch (IOException ex) {
            return false;
        }
    }

    static void writePreviewMeta(File pdfFile, File sourceFile) throws IOException {
        Files.createDirectories(previewMetaFile(pdfFile).getParentFile().toPath());
        JSON.writeValue(
                previewMetaFile(pdfFile),
                new PreviewMeta(
                        fingerprint(sourceFile),
                        RequestFormSheetPreviewRenderer.PREVIEW_RANGE_SPEC,
                        RequestFormSheetPreviewPdfRenderer.rendererSpec(),
                        PREVIEW_SCHEMA_VERSION,
                        nowMillis()));
    }

    static void deletePreviewCache(File pdfFile) {
        if (pdfFile != null && pdfFile.exists()) {
            pdfFile.delete();
        }
        File meta = pdfFile != null ? previewMetaFile(pdfFile) : null;
        if (meta != null && meta.exists()) {
            meta.delete();
        }
    }

    /**
     * {@code preview_cache/pdf} と {@code preview_cache/parse} 配下のファイルをすべて削除する。
     * ディレクトリ自体は残す。
     */
    static ClearDiskCacheResult clearAllDiskCache(File cacheRoot) {
        if (cacheRoot == null) {
            return new ClearDiskCacheResult(0, 0, 0);
        }
        int[] pdf = new int[2];
        int[] parse = new int[2];
        clearDirectoryFiles(pdfDir(cacheRoot), pdf);
        clearDirectoryFiles(parseDir(cacheRoot), parse);
        clearDirectoryFiles(splitDir(cacheRoot), parse);
        return new ClearDiskCacheResult(pdf[0], parse[0], pdf[1] + parse[1]);
    }

    /**
     * スキーマ不一致の parse キャッシュと、tpi-split の孤立メタを削除する。
     *
     * @return 削除したファイル数
     */
    static int pruneStaleDiskCaches(File cacheRoot) {
        if (cacheRoot == null) {
            return 0;
        }
        return pruneStaleParseCacheFiles(cacheRoot)
                + pruneStalePreviewCacheFiles(cacheRoot)
                + pruneStaleTpiSplitCacheFiles(cacheRoot)
                + pruneOrphanTpiSplitArtifacts(cacheRoot);
    }

    static int pruneStaleParseCacheFiles(File cacheRoot) {
        File dir = parseDir(cacheRoot);
        File[] children = dir.listFiles();
        if (children == null) {
            return 0;
        }
        int deleted = 0;
        for (File child : children) {
            if (child == null || !child.isFile() || !child.getName().endsWith(".json")) {
                continue;
            }
            if (isStaleParseCacheFile(child) || isFileOlderThanMaxAge(child)) {
                if (deleteFileQuietly(child)) {
                    deleted++;
                }
            }
        }
        return deleted;
    }

    static int pruneStalePreviewCacheFiles(File cacheRoot) {
        File dir = pdfDir(cacheRoot);
        File[] children = dir.listFiles();
        if (children == null) {
            return 0;
        }
        int deleted = 0;
        for (File child : children) {
            if (child == null || !child.isFile()) {
                continue;
            }
            String name = child.getName();
            if (!name.endsWith(".meta.json")) {
                continue;
            }
            File pdf =
                    new File(child.getParentFile(), name.substring(0, name.length() - ".meta.json".length()));
            if (isStalePreviewMetaFile(child) || isFileOlderThanMaxAge(pdf)) {
                if (deleteFileQuietly(pdf)) {
                    deleted++;
                }
                if (deleteFileQuietly(child)) {
                    deleted++;
                }
            } else if (name.endsWith(".pdf") && isFileOlderThanMaxAge(child)) {
                File meta = previewMetaFile(child);
                if (deleteFileQuietly(child)) {
                    deleted++;
                }
                if (deleteFileQuietly(meta)) {
                    deleted++;
                }
            }
        }
        return deleted;
    }

    static int pruneStaleTpiSplitCacheFiles(File cacheRoot) {
        File dir = splitDir(cacheRoot);
        File[] children = dir.listFiles();
        if (children == null) {
            return 0;
        }
        int deleted = 0;
        for (File child : children) {
            if (child == null || !child.isFile()) {
                continue;
            }
            String name = child.getName();
            if (name.endsWith(".meta.json")) {
                if (isStaleSplitMetaFile(child)) {
                    File pdf =
                            new File(
                                    child.getParentFile(),
                                    name.substring(0, name.length() - ".meta.json".length()));
                    if (deleteFileQuietly(pdf)) {
                        deleted++;
                    }
                    if (deleteFileQuietly(child)) {
                        deleted++;
                    }
                }
            } else if (name.endsWith(".pdf")) {
                File meta = splitMetaFile(child);
                if (!meta.isFile() || isStaleSplitMetaFile(meta) || isFileOlderThanMaxAge(child)) {
                    if (deleteFileQuietly(child)) {
                        deleted++;
                    }
                    if (deleteFileQuietly(meta)) {
                        deleted++;
                    }
                }
            }
        }
        return deleted;
    }

    /** parse キャッシュと、TPI 束ね PDF に紐づく split 成果物を削除する。 */
    static void invalidateParseCacheForSource(File cacheRoot, File sourceFile) {
        deleteFileQuietly(parseCacheFile(cacheRoot, sourceFile));
        if (sourceFile != null
                && sourceFile.getName().toLowerCase(java.util.Locale.ROOT).endsWith(".pdf")) {
            deleteTpiSplitArtifactsForSourcePdf(cacheRoot, sourceFile);
        }
    }

    static void deleteTpiSplitArtifactsForSourcePdf(File cacheRoot, File sourcePdf) {
        if (cacheRoot == null || sourcePdf == null) {
            return;
        }
        String prefix = safeBaseName(sourcePdf.getName()) + "__";
        File dir = splitDir(cacheRoot);
        File[] children = dir.listFiles();
        if (children == null) {
            return;
        }
        for (File child : children) {
            if (child == null || !child.isFile()) {
                continue;
            }
            if (child.getName().startsWith(prefix)) {
                deleteFileQuietly(child);
            }
        }
    }

    static int pruneOrphanTpiSplitArtifacts(File cacheRoot) {
        File dir = splitDir(cacheRoot);
        File[] children = dir.listFiles();
        if (children == null) {
            return 0;
        }
        int deleted = 0;
        for (File child : children) {
            if (child == null || !child.isFile()) {
                continue;
            }
            String name = child.getName();
            if (name.endsWith(".meta.json")) {
                File pdf = new File(child.getParentFile(), name.substring(0, name.length() - ".meta.json".length()));
                if (!pdf.isFile()) {
                    if (deleteFileQuietly(child)) {
                        deleted++;
                    }
                }
            }
        }
        return deleted;
    }

    private static void clearDirectoryFiles(File dir, int[] deletedAndFailures) {
        if (dir == null || !dir.isDirectory()) {
            return;
        }
        File[] children = dir.listFiles();
        if (children == null) {
            return;
        }
        for (File child : children) {
            if (child.isFile()) {
                if (child.delete()) {
                    deletedAndFailures[0]++;
                } else {
                    deletedAndFailures[1]++;
                }
            }
        }
    }

    private static boolean looksLikePdf(File file) {
        try (FileInputStream in = new FileInputStream(file)) {
            byte[] header = in.readNBytes(5);
            return header.length >= 5
                    && header[0] == '%'
                    && header[1] == 'P'
                    && header[2] == 'D'
                    && header[3] == 'F'
                    && header[4] == '-';
        } catch (IOException ex) {
            return false;
        }
    }

    static SourceFingerprint fingerprint(File source) {
        return new SourceFingerprint(source.lastModified(), source.length());
    }

    static boolean matches(File source, SourceFingerprint fingerprint) {
        if (source == null || !source.isFile() || fingerprint == null) {
            return false;
        }
        return source.lastModified() == fingerprint.lastModified()
                && source.length() == fingerprint.length();
    }

    static File parseCacheFile(File cacheRoot, File sourceFile) {
        return new File(parseDir(cacheRoot), safeBaseName(sourceFile.getName()) + ".json");
    }

    static String parseSchemaVersionFor(File sourceFile) {
        if (sourceFile != null
                && sourceFile.getName().toLowerCase(java.util.Locale.ROOT).endsWith(".pdf")) {
            return TPI_PDF_PARSE_SCHEMA_VERSION;
        }
        return PARSE_SCHEMA_VERSION;
    }

    static Optional<List<Map<String, String>>> loadParseEntries(File cacheRoot, File sourceFile) {
        File cacheFile = parseCacheFile(cacheRoot, sourceFile);
        if (!cacheFile.isFile()) {
            return Optional.empty();
        }
        try {
            ParseCachePayload payload = JSON.readValue(cacheFile, ParseCachePayload.class);
            if (payload == null
                    || payload.entries() == null
                    || !parseSchemaVersionFor(sourceFile).equals(payload.schemaVersion())
                    || !matches(sourceFile, payload.source())
                    || isCacheExpired(payload.cachedAtMillis())) {
                invalidateParseCacheForSource(cacheRoot, sourceFile);
                return Optional.empty();
            }
            List<Map<String, String>> copy = new ArrayList<>(payload.entries().size());
            for (Map<String, String> entry : payload.entries()) {
                copy.add(entry != null ? new LinkedHashMap<>(entry) : new LinkedHashMap<>());
            }
            return Optional.of(copy);
        } catch (IOException ex) {
            invalidateParseCacheForSource(cacheRoot, sourceFile);
            return Optional.empty();
        }
    }

    static void saveParseEntries(
            File cacheRoot, File sourceFile, List<Map<String, String>> entries) throws IOException {
        File cacheFile = parseCacheFile(cacheRoot, sourceFile);
        Files.createDirectories(cacheFile.getParentFile().toPath());
        List<Map<String, String>> stored = new ArrayList<>();
        if (entries != null) {
            for (Map<String, String> entry : entries) {
                stored.add(entry != null ? new LinkedHashMap<>(entry) : new LinkedHashMap<>());
            }
        }
        ParseCachePayload payload =
                new ParseCachePayload(
                        fingerprint(sourceFile),
                        parseSchemaVersionFor(sourceFile),
                        nowMillis(),
                        stored);
        JSON.writerWithDefaultPrettyPrinter().writeValue(cacheFile, payload);
    }

    private static boolean isStalePreviewMetaFile(File metaFile) {
        try {
            PreviewMeta meta = JSON.readValue(metaFile, PreviewMeta.class);
            if (meta == null) {
                return true;
            }
            return !PREVIEW_SCHEMA_VERSION.equals(meta.schemaVersion())
                    || !RequestFormSheetPreviewRenderer.PREVIEW_RANGE_SPEC.equals(meta.range())
                    || !RequestFormSheetPreviewPdfRenderer.rendererSpec().equals(meta.renderer())
                    || isCacheExpired(meta.cachedAtMillis());
        } catch (IOException ex) {
            return true;
        }
    }

    private static boolean isStaleSplitMetaFile(File metaFile) {
        try {
            SplitCacheMeta meta = JSON.readValue(metaFile, SplitCacheMeta.class);
            if (meta == null) {
                return true;
            }
            return !TPI_PDF_PARSE_SCHEMA_VERSION.equals(meta.schemaVersion())
                    || isCacheExpired(meta.cachedAtMillis());
        } catch (IOException ex) {
            return true;
        }
    }

    static boolean isCacheExpired(long cachedAtMillis) {
        if (cachedAtMillis <= 0L) {
            return true;
        }
        return nowMillis() - cachedAtMillis > CACHE_MAX_AGE_MILLIS;
    }

    static boolean isFileOlderThanMaxAge(File file) {
        if (file == null || !file.isFile()) {
            return false;
        }
        return nowMillis() - file.lastModified() > CACHE_MAX_AGE_MILLIS;
    }

    private static long nowMillis() {
        return System.currentTimeMillis();
    }

    private static boolean isStaleParseCacheFileByPayload(ParseCachePayload payload) {
        if (payload == null || payload.schemaVersion() == null) {
            return true;
        }
        if (isCacheExpired(payload.cachedAtMillis())) {
            return true;
        }
        String schema = payload.schemaVersion();
        if (schema.startsWith("request-form-tpi-pdf-")) {
            return !TPI_PDF_PARSE_SCHEMA_VERSION.equals(schema);
        }
        if (schema.startsWith("request-form-cell-layout-")) {
            return !PARSE_SCHEMA_VERSION.equals(schema);
        }
        return true;
    }

    private static boolean isStaleParseCacheFile(File cacheFile) {
        try {
            ParseCachePayload payload = JSON.readValue(cacheFile, ParseCachePayload.class);
            return isStaleParseCacheFileByPayload(payload);
        } catch (IOException ex) {
            return true;
        }
    }

    private static boolean deleteFileQuietly(File file) {
        if (file == null || !file.isFile()) {
            return false;
        }
        try {
            return file.delete();
        } catch (Exception ex) {
            return false;
        }
    }

    private static String safeBaseName(String name) {
        if (name == null || name.isBlank()) {
            return "unknown";
        }
        int dot = name.lastIndexOf('.');
        String base = dot > 0 ? name.substring(0, dot) : name;
        return base.replaceAll("[\\\\/:*?\"<>|]", "_");
    }
}
