package jp.co.pm.ai.desktop.resources;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;
import java.nio.charset.StandardCharsets;

import org.junit.jupiter.api.Test;

/**
 * CI / headless: ensures packaged JSON schema resources are present (no JavaFX).
 */
class SchemaClasspathTest {

    @Test
    void ipcLineSchemaOnClasspath() throws Exception {
        try (InputStream in = classLoaderStream("schema/ipc-line.schema.json")) {
            assertNotNull(in, "schema/ipc-line.schema.json");
            String head = new String(in.readAllBytes(), StandardCharsets.UTF_8).trim();
            assertTrue(head.startsWith("{"), "ipc-line.schema.json should look like JSON");
        }
    }

    @Test
    void columnsSchemaOnClasspath() throws Exception {
        try (InputStream in = classLoaderStream("schema/columns.schema.json")) {
            assertNotNull(in, "schema/columns.schema.json");
        }
    }

    private static InputStream classLoaderStream(String path) {
        return SchemaClasspathTest.class.getClassLoader().getResourceAsStream(path);
    }
}
