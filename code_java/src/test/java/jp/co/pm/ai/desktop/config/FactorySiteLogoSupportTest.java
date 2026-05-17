package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.Test;

class FactorySiteLogoSupportTest {

    @Test
    void bundledKonanAndKokubuLogosLoad() {
        assertTrue(FactorySiteLogoSupport.resolveImage(FactorySite.KONAN, Map.of()).isPresent());
        assertTrue(FactorySiteLogoSupport.resolveImage(FactorySite.KOKUBU, Map.of()).isPresent());
    }

    @Test
    void bundledResourcePathsExistOnClasspath() {
        for (FactorySite site : FactorySite.values()) {
            String path = FactorySiteLogoSupport.bundledResourcePath(site);
            assertNotNull(
                    FactorySiteLogoSupport.class.getResource(path),
                    "missing bundled logo: " + path);
        }
    }
}
