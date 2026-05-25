package jp.co.pm.ai.desktop.reconciliation;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javax.imageio.ImageIO;

import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 依頼書原本の解析結果とプレビュー PNG を {@code preview_cache/} 以下に保持する。
 * 原本 Excel の更新日時・サイズが一致するときのみキャッシュを再利用する。
 */
final class RequestFormSourceCache {

    private static final ObjectMapper JSON = new ObjectMapper();

    private RequestFormSourceCache() {}

    record SourceFingerprint(long lastModified, long length) {}

    private record ParseCachePayload(SourceFingerprint source, List<Map<String, String>> entries) {}

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

    static File pngDir(File cacheRoot) {
        File dir = new File(cacheRoot, "png");
        if (!dir.exists()) {
            dir.mkdirs();
        }
        return dir;
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
        ParseCachePayload payload = new ParseCachePayload(fingerprint(sourceFile), stored);
        JSON.writerWithDefaultPrettyPrinter().writeValue(cacheFile, payload);
    }

    static File pngCacheFile(File cacheRoot, String workbookName, String sheetName) {
        String cacheName =
                (workbookName + "_" + sheetName).replaceAll("[\\\\/:*?\"<>|]", "_") + ".png";
        return new File(pngDir(cacheRoot), cacheName);
    }

    static File pngMetaFile(File pngFile) {
        return new File(pngFile.getAbsolutePath() + ".meta.json");
    }

    static boolean isPreviewCacheValid(File pngFile, File sourceFile) {
        if (pngFile == null || !pngFile.isFile() || pngFile.length() < 128) {
            return false;
        }
        File metaFile = pngMetaFile(pngFile);
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
            BufferedImage img = ImageIO.read(pngFile);
            return img != null && img.getWidth() >= 16 && img.getHeight() >= 16;
        } catch (IOException ex) {
            return false;
        }
    }

    static void writePreviewMeta(File pngFile, File sourceFile) throws IOException {
        Files.createDirectories(pngMetaFile(pngFile).getParentFile().toPath());
        JSON.writeValue(
                pngMetaFile(pngFile),
                new PreviewMeta(
                        fingerprint(sourceFile),
                        RequestFormSheetPreviewRenderer.PREVIEW_RANGE_SPEC,
                        RequestFormSheetPreviewRenderer.PREVIEW_RENDERER_SPEC));
    }

    static void deletePreviewCache(File pngFile) {
        if (pngFile != null && pngFile.exists()) {
            pngFile.delete();
        }
        File meta = pngFile != null ? pngMetaFile(pngFile) : null;
        if (meta != null && meta.exists()) {
            meta.delete();
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
