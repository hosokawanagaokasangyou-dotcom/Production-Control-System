package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 依頼書原本の解析結果とプレビュー PDF を {@code preview_cache/} 以下に保持する。
 * 原本 Excel の更新日時・サイズと {@link #PARSE_SCHEMA_VERSION} が一致するときのみキャッシュを再利用する。
 */
final class RequestFormSourceCache {

    private static final ObjectMapper JSON = new ObjectMapper();

    /**
     * {@link RequestFormOriginalCellLayout} ベースのセル読取に合わせて上げる。
     * 抽出ロジック変更時は必ずインクリメントし、古い parse キャッシュを無効化する。
     */
    static final String PARSE_SCHEMA_VERSION = "request-form-cell-layout-v3";

    private RequestFormSourceCache() {}

    record SourceFingerprint(long lastModified, long length) {}

    private record ParseCachePayload(
            SourceFingerprint source, String schemaVersion, List<Map<String, String>> entries) {}

    private record PreviewMeta(SourceFingerprint source, String range, String renderer) {}

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
                    || !RequestFormSheetPreviewRenderer.PREVIEW_RENDERER_SPEC.equals(
                            meta.renderer())) {
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
                        RequestFormSheetPreviewRenderer.PREVIEW_RENDERER_SPEC));
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

    static Optional<List<Map<String, String>>> loadParseEntries(File cacheRoot, File sourceFile) {
        File cacheFile = parseCacheFile(cacheRoot, sourceFile);
        if (!cacheFile.isFile()) {
            return Optional.empty();
        }
        try {
            ParseCachePayload payload = JSON.readValue(cacheFile, ParseCachePayload.class);
            if (payload == null
                    || payload.entries() == null
                    || !PARSE_SCHEMA_VERSION.equals(payload.schemaVersion())
                    || !matches(sourceFile, payload.source())) {
                return Optional.empty();
            }
            List<Map<String, String>> copy = new ArrayList<>(payload.entries().size());
            for (Map<String, String> entry : payload.entries()) {
                copy.add(entry != null ? new LinkedHashMap<>(entry) : new LinkedHashMap<>());
            }
            return Optional.of(copy);
        } catch (IOException ex) {
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
                new ParseCachePayload(fingerprint(sourceFile), PARSE_SCHEMA_VERSION, stored);
        JSON.writerWithDefaultPrettyPrinter().writeValue(cacheFile, payload);
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
