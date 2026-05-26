package jp.co.pm.ai.desktop.io;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertTrue;

class PoiWorkbookOpenerTest {

    @Test
    void configureZipSecureLimits_raisesMaxFileCount() {
        PoiWorkbookOpener.configureZipSecureLimits();
        assertDoesNotThrow(() -> ZipSecureFile.setMaxFileCount(ZipSecureFile.getMaxFileCount()));
        assertTrue(ZipSecureFile.getMaxFileCount() >= 10_000);
    }
}
