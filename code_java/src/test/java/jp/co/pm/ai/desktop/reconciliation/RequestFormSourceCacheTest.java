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
}
