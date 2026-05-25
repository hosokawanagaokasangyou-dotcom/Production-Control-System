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

        File png = RequestFormSourceCache.pngCacheFile(cacheRoot, "book.xlsm", "E5-1");
        Files.createDirectories(png.getParentFile().toPath());
        Files.write(png.toPath(), new byte[] {(byte) 0x89, 0x50, 0x4E, 0x47, 0, 0, 0, 0});
        RequestFormSourceCache.writePreviewMeta(png, source);

        Files.writeString(source.toPath(), "v2");
        assertFalse(RequestFormSourceCache.isPreviewCacheValid(png, source));
    }

    @Test
    void previewMeta_invalidatesWhenRangeSpecChanges(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("book.xlsm").toFile();
        Files.writeString(source.toPath(), "v1");

        File png = RequestFormSourceCache.pngCacheFile(cacheRoot, "book.xlsm", "E5-1");
        Files.createDirectories(png.getParentFile().toPath());
        Files.write(png.toPath(), new byte[] {(byte) 0x89, 0x50, 0x4E, 0x47, 0, 0, 0, 0});
        Files.writeString(
                RequestFormSourceCache.pngMetaFile(png).toPath(),
                "{\"source\":{\"lastModified\":"
                        + source.lastModified()
                        + ",\"length\":"
                        + source.length()
                        + "},\"range\":\"A1:Z99\"}");

        assertFalse(RequestFormSourceCache.isPreviewCacheValid(png, source));
    }

    @Test
    void previewMeta_invalidatesWhenRendererSpecChanges(@TempDir Path tmp) throws Exception {
        File cacheRoot = tmp.resolve("preview_cache").toFile();
        File source = tmp.resolve("book.xlsm").toFile();
        Files.writeString(source.toPath(), "v1");

        File png = RequestFormSourceCache.pngCacheFile(cacheRoot, "book.xlsm", "E5-1");
        Files.createDirectories(png.getParentFile().toPath());
        Files.write(png.toPath(), new byte[] {(byte) 0x89, 0x50, 0x4E, 0x47, 0, 0, 0, 0});
        Files.writeString(
                RequestFormSourceCache.pngMetaFile(png).toPath(),
                "{\"source\":{\"lastModified\":"
                        + source.lastModified()
                        + ",\"length\":"
                        + source.length()
                        + "},\"range\":\"A1:AN28\",\"renderer\":\"legacy\"}");

        assertFalse(RequestFormSourceCache.isPreviewCacheValid(png, source));
    }
}
