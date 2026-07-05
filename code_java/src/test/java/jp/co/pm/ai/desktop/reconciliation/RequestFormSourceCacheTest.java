package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RequestFormSourceCacheTest {

    @Test
    void parseCache_invalidatesWhenSchemaVersionChanges(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("依頼書B.xlsm").toFile();
        Files.writeString(source.toPath(), "dummy");

        File cacheFile = RequestFormSourceCache.parseCacheFile(cacheRoot, source);
        Files.createDirectories(cacheFile.getParentFile().toPath());
        Files.writeString(
                cacheFile.toPath(),
                "{\"source\":{\"lastModified\":"
                        + source.lastModified()
                        + ",\"length\":"
                        + source.length()
                        + "},\"entries\":[{\"依頼Ｎｏ\":\"E6-2\",\"加工賃\":\"18+18+13\"}]}");

        assertTrue(RequestFormSourceCache.loadParseEntries(cacheRoot, source).isEmpty());
        assertFalse(cacheFile.isFile(), "stale parse cache should be deleted");
    }

    @Test
    void parseCache_reusesEntriesWhileSourceUnchanged(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("依頼書A.xlsm").toFile();
        Files.writeString(source.toPath(), "dummy");

        List<Map<String, String>> entries =
                List.of(Map.of("依頼Ｎｏ", "E5-1", "原本シート名", "E5-1"));
        RequestFormSourceCache.saveParseEntries(cacheRoot, source, entries);

        Optional<List<Map<String, String>>> loaded =
                RequestFormSourceCache.loadParseEntries(cacheRoot, source);
        assertTrue(loaded.isPresent());
        assertEquals("E5-1", loaded.get().get(0).get("依頼Ｎｏ"));

        Files.writeString(source.toPath(), "updated");
        assertTrue(RequestFormSourceCache.loadParseEntries(cacheRoot, source).isEmpty());
        assertFalse(RequestFormSourceCache.parseCacheFile(cacheRoot, source).isFile());
    }

    @Test
    void previewMeta_invalidatesWhenSourceChanges(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("book.xlsm").toFile();
        Files.writeString(source.toPath(), "v1");

        File pdf = RequestFormSourceCache.pdfCacheFile(cacheRoot, "book.xlsm", "E5-1");
        Files.createDirectories(pdf.getParentFile().toPath());
        Files.write(pdf.toPath(), "%PDF-1.4\n% preview\n".getBytes(java.nio.charset.StandardCharsets.US_ASCII));
        RequestFormSourceCache.writePreviewMeta(pdf, source);

        Files.writeString(source.toPath(), "v2");
        assertFalse(RequestFormSourceCache.isPreviewCacheValid(pdf, source));
    }

    @Test
    void previewMeta_invalidatesWhenRangeSpecChanges(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("book.xlsm").toFile();
        Files.writeString(source.toPath(), "v1");

        File pdf = RequestFormSourceCache.pdfCacheFile(cacheRoot, "book.xlsm", "E5-1");
        Files.createDirectories(pdf.getParentFile().toPath());
        Files.write(pdf.toPath(), "%PDF-1.4\n% preview\n".getBytes(java.nio.charset.StandardCharsets.US_ASCII));
        Files.writeString(
                RequestFormSourceCache.previewMetaFile(pdf).toPath(),
                "{\"source\":{\"lastModified\":"
                        + source.lastModified()
                        + ",\"length\":"
                        + source.length()
                        + "},\"range\":\"A1:Z99\"}");

        assertFalse(RequestFormSourceCache.isPreviewCacheValid(pdf, source));
    }

    @Test
    void previewMeta_invalidatesWhenRendererSpecChanges(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("book.xlsm").toFile();
        Files.writeString(source.toPath(), "v1");

        File pdf = RequestFormSourceCache.pdfCacheFile(cacheRoot, "book.xlsm", "E5-1");
        Files.createDirectories(pdf.getParentFile().toPath());
        Files.write(pdf.toPath(), "%PDF-1.4\n% preview\n".getBytes(java.nio.charset.StandardCharsets.US_ASCII));
        Files.writeString(
                RequestFormSourceCache.previewMetaFile(pdf).toPath(),
                "{\"source\":{\"lastModified\":"
                        + source.lastModified()
                        + ",\"length\":"
                        + source.length()
                        + "},\"range\":\"A1:AO29\",\"renderer\":\"legacy\"}");

        assertFalse(RequestFormSourceCache.isPreviewCacheValid(pdf, source));
    }

    @Test
    void clearAllDiskCache_removesPdfAndParseFiles(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("book.xlsm").toFile();
        Files.writeString(source.toPath(), "v1");

        File pdf = RequestFormSourceCache.pdfCacheFile(cacheRoot, "book.xlsm", "E5-1");
        Files.createDirectories(pdf.getParentFile().toPath());
        Files.write(pdf.toPath(), "%PDF-1.4\n".getBytes(java.nio.charset.StandardCharsets.US_ASCII));
        RequestFormSourceCache.writePreviewMeta(pdf, source);
        RequestFormSourceCache.saveParseEntries(cacheRoot, source, List.of(Map.of("依頼Ｎｏ", "E5-1")));

        RequestFormSourceCache.ClearDiskCacheResult result =
                RequestFormSourceCache.clearAllDiskCache(cacheRoot);
        assertEquals(2, result.pdfFilesDeleted());
        assertEquals(1, result.parseFilesDeleted());
        assertEquals(0, result.deleteFailures());
        assertFalse(pdf.isFile());
        assertFalse(RequestFormSourceCache.parseCacheFile(cacheRoot, source).isFile());
    }

    @Test
    void pruneStaleParseCacheFiles_removesOutdatedTpiSchema(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File stale =
                new File(RequestFormSourceCache.parseDir(cacheRoot), "GB__GB60604.json");
        Files.createDirectories(stale.getParentFile().toPath());
        Files.writeString(
                stale.toPath(),
                "{\"schemaVersion\":\"request-form-tpi-pdf-v8\",\"entries\":[]}");
        File current =
                new File(RequestFormSourceCache.parseDir(cacheRoot), "GB.json");
        Files.writeString(
                current.toPath(),
                "{\"schemaVersion\":\""
                        + RequestFormSourceCache.TPI_PDF_PARSE_SCHEMA_VERSION
                        + "\",\"cachedAtMillis\":"
                        + System.currentTimeMillis()
                        + ",\"entries\":[]}");

        int deleted = RequestFormSourceCache.pruneStaleParseCacheFiles(cacheRoot);

        assertEquals(1, deleted);
        assertFalse(stale.isFile());
        assertTrue(current.isFile());
    }

    @Test
    void previewMeta_invalidatesWhenSchemaVersionMissing(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("book.xlsm").toFile();
        Files.writeString(source.toPath(), "v1");

        File pdf = RequestFormSourceCache.pdfCacheFile(cacheRoot, "book.xlsm", "E5-1");
        Files.createDirectories(pdf.getParentFile().toPath());
        Files.write(pdf.toPath(), "%PDF-1.4\n% preview\n".getBytes(java.nio.charset.StandardCharsets.US_ASCII));
        Files.writeString(
                RequestFormSourceCache.previewMetaFile(pdf).toPath(),
                "{\"source\":{\"lastModified\":"
                        + source.lastModified()
                        + ",\"length\":"
                        + source.length()
                        + "},\"range\":\""
                        + RequestFormSheetPreviewRenderer.PREVIEW_RANGE_SPEC
                        + "\",\"renderer\":\""
                        + RequestFormSheetPreviewPdfRenderer.rendererSpec()
                        + "\"}");

        assertFalse(RequestFormSourceCache.isPreviewCacheValid(pdf, source));
        assertEquals(2, RequestFormSourceCache.pruneStalePreviewCacheFiles(cacheRoot));
        assertFalse(pdf.isFile());
    }

    @Test
    void loadParseEntries_invalidatesTpiSplitWhenSchemaStale(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("GB.pdf").toFile();
        Files.writeString(source.toPath(), "dummy");

        File splitPdf =
                RequestFormSourceCache.splitCacheFile(cacheRoot, source, "GB60606");
        Files.createDirectories(splitPdf.getParentFile().toPath());
        Files.write(splitPdf.toPath(), "%PDF-1.4\n".getBytes(java.nio.charset.StandardCharsets.US_ASCII));
        Files.writeString(
                RequestFormSourceCache.splitMetaFile(splitPdf).toPath(),
                "{\"source\":{\"lastModified\":"
                        + source.lastModified()
                        + ",\"length\":"
                        + source.length()
                        + "},\"startPage\":0,\"endPage\":0,\"schemaVersion\":\"request-form-tpi-pdf-v8\"}");

        File cacheFile = RequestFormSourceCache.parseCacheFile(cacheRoot, source);
        Files.createDirectories(cacheFile.getParentFile().toPath());
        Files.writeString(
                cacheFile.toPath(),
                "{\"source\":{\"lastModified\":"
                        + source.lastModified()
                        + ",\"length\":"
                        + source.length()
                        + "},\"schemaVersion\":\"request-form-tpi-pdf-v8\",\"entries\":[]}");

        assertTrue(RequestFormSourceCache.loadParseEntries(cacheRoot, source).isEmpty());
        assertFalse(splitPdf.isFile());
        assertFalse(RequestFormSourceCache.splitMetaFile(splitPdf).isFile());
    }

    @Test
    void pruneStaleTpiSplitCacheFiles_removesOutdatedSchema(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File splitDir = RequestFormSourceCache.splitDir(cacheRoot);
        File splitPdf = new File(splitDir, "GB__GB60606.pdf");
        Files.createDirectories(splitDir.toPath());
        Files.write(splitPdf.toPath(), "%PDF-1.4\n".getBytes(java.nio.charset.StandardCharsets.US_ASCII));
        Files.writeString(
                RequestFormSourceCache.splitMetaFile(splitPdf).toPath(),
                "{\"schemaVersion\":\"request-form-tpi-pdf-v8\",\"startPage\":0,\"endPage\":0}");

        int deleted = RequestFormSourceCache.pruneStaleTpiSplitCacheFiles(cacheRoot);

        assertEquals(2, deleted);
        assertFalse(splitPdf.isFile());
    }

    @Test
    void parseCache_invalidatesWhenOlderThanMaxAge(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("依頼書C.xlsm").toFile();
        Files.writeString(source.toPath(), "dummy");

        List<Map<String, String>> entries = List.of(Map.of("依頼Ｎｏ", "E7-1"));
        RequestFormSourceCache.saveParseEntries(cacheRoot, source, entries);

        File cacheFile = RequestFormSourceCache.parseCacheFile(cacheRoot, source);
        long expiredAt =
                System.currentTimeMillis()
                        - RequestFormSourceCache.CACHE_MAX_AGE_MILLIS
                        - 1L;
        String json = Files.readString(cacheFile.toPath());
        json = json.replaceFirst("\"cachedAtMillis\"\\s*:\\s*\\d+", "\"cachedAtMillis\":" + expiredAt);
        Files.writeString(cacheFile.toPath(), json);

        assertTrue(RequestFormSourceCache.loadParseEntries(cacheRoot, source).isEmpty());
        assertFalse(cacheFile.isFile());
    }

    @Test
    void pruneStaleParseCacheFiles_removesExpiredByFileAge(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File cacheFile =
                new File(RequestFormSourceCache.parseDir(cacheRoot), "old-book.json");
        Files.createDirectories(cacheFile.getParentFile().toPath());
        Files.writeString(
                cacheFile.toPath(),
                "{\"schemaVersion\":\""
                        + RequestFormSourceCache.PARSE_SCHEMA_VERSION
                        + "\",\"cachedAtMillis\":"
                        + System.currentTimeMillis()
                        + ",\"entries\":[]}");
        long oldTime =
                System.currentTimeMillis()
                        - RequestFormSourceCache.CACHE_MAX_AGE_MILLIS
                        - 86_400_000L;
        assertTrue(cacheFile.setLastModified(oldTime));

        int deleted = RequestFormSourceCache.pruneStaleParseCacheFiles(cacheRoot);

        assertEquals(1, deleted);
        assertFalse(cacheFile.isFile());
    }

    @Test
    void pruneOrphanTpiSplitArtifacts_removesMetaWithoutPdf(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File splitDir = RequestFormSourceCache.splitDir(cacheRoot);
        File meta = new File(splitDir, "GB__GB60604.pdf.meta.json");
        Files.createDirectories(splitDir.toPath());
        Files.writeString(meta.toPath(), "{}");

        int deleted = RequestFormSourceCache.pruneOrphanTpiSplitArtifacts(cacheRoot);

        assertEquals(1, deleted);
        assertFalse(meta.isFile());
    }
}
