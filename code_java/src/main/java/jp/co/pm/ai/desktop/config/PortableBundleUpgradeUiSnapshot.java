package jp.co.pm.ai.desktop.config;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.databind.ObjectMapper;

/** ポータルバージョンアップ前のメインシェルタブ UI スナップショット。 */
public final class PortableBundleUpgradeUiSnapshot {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Path STORE_PATH =
            AppPaths.resolveDesktopAppHomeDir().resolve("pre-upgrade-shell-ui.json");

    private PortableBundleUpgradeUiSnapshot() {}

    public static void capture(DesktopSessionState state) {
        if (state == null) {
            return;
        }
        try {
            Files.createDirectories(STORE_PATH.getParent());
            JSON.writeValue(
                    STORE_PATH.toFile(),
                    DesktopSessionStateStore.toJsonObject(state.extractShellTabUiFields()));
        } catch (Exception ignored) {
        }
    }

    public static DesktopSessionState loadIfPresent() {
        try {
            if (!Files.isRegularFile(STORE_PATH)) {
                return null;
            }
            var root = JSON.readTree(STORE_PATH.toFile());
            return DesktopSessionStateStore.parseSessionFragment(root);
        } catch (Exception ignored) {
            return null;
        }
    }

    public static void clear() {
        try {
            Files.deleteIfExists(STORE_PATH);
        } catch (Exception ignored) {
        }
    }

    public static Path storePathForTests() {
        return STORE_PATH;
    }
}
