package jp.co.pm.ai.desktop.reconciliation;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;

class JuchuHeaderAliasRegistryHeaderRowTest {

    @TempDir Path tempDir;

    @Test
    void headerRow_defaultsToRow3() {
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        assertEquals(3, registry.headerRowOneBasedFor("C:\\test\\juchu.xlsm"));
        assertEquals(2, registry.headerRowIndexFor("C:\\test\\juchu.xlsm"));
    }

    @Test
    void headerRow_perFileOverrideAndPersistence() throws Exception {
        Path store = tempDir.resolve("aliases.properties");
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry(store);
        String path = "C:\\test\\juchu.xlsm";

        registry.setHeaderRowOneBasedFor(path, 5);
        registry.saveToDisk();

        JuchuHeaderAliasRegistry reloaded = new JuchuHeaderAliasRegistry(store);
        reloaded.reloadFromDisk();
        assertEquals(5, reloaded.headerRowOneBasedFor(path));
        assertEquals(4, reloaded.headerRowIndexFor(path));
    }

    @Test
    void headerRow_sameAsFactoryDefaultIsNotStored() throws Exception {
        Path store = tempDir.resolve("aliases.properties");
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry(store);
        String path = "C:\\test\\juchu.xlsm";

        registry.setFactoryDefaultHeaderRowOneBased(4);
        registry.setHeaderRowOneBasedFor(path, 4);
        registry.saveToDisk();

        JuchuHeaderAliasRegistry reloaded = new JuchuHeaderAliasRegistry(store);
        reloaded.reloadFromDisk();
        assertEquals(4, reloaded.headerRowOneBasedFor(path));
        assertEquals(4, reloaded.factoryDefaultHeaderRowOneBased());
    }

    @Test
    void resolveHeaderRow_usesRegistry() {
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry();
        String path = "C:\\test\\juchu.xlsm";
        registry.setHeaderRowOneBasedFor(path, 7);

        assertEquals(7, JuchuSheetColumnLayout.resolveHeaderRowOneBased(registry, path));
        assertEquals(6, JuchuSheetColumnLayout.resolveHeaderRowIndex(registry, path));
        assertEquals(7, JuchuSheetColumnLayout.resolveFirstDataRowIndex(registry, path));
    }
}
