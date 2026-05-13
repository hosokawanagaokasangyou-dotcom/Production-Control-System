package jp.co.pm.ai.desktop.config;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * グローバル設定（init_setting への書き出し・「デフォルトに戻す」時の参照）で使う工場。
 *
 * <p>値は {@code ~/.pm-ai-desktop/global-init-setting-target-factory.txt} に {@link FactorySite#name()}（{@code KONAN} /
 * {@code KOKUBU}）で保存する。環境変数の工場プリセット適用時も同ファイルを同期し、意図しない工場の既定を読まないようにする。
 */
public final class GlobalInitSettingTarget {

    private static final Path STORE =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "global-init-setting-target-factory.txt");

    private GlobalInitSettingTarget() {}

    public static Path storePathForTests() {
        return STORE;
    }

    /** 未設定・不正時は {@link FactorySite#KONAN}。 */
    public static FactorySite load() {
        try {
            if (!Files.isRegularFile(STORE)) {
                return FactorySite.KONAN;
            }
            String raw = Files.readString(STORE, StandardCharsets.UTF_8).trim();
            if (raw.isEmpty()) {
                return FactorySite.KONAN;
            }
            return FactorySite.valueOf(raw);
        } catch (Exception ignored) {
            return FactorySite.KONAN;
        }
    }

    public static void save(FactorySite site) {
        if (site == null) {
            return;
        }
        try {
            Files.createDirectories(STORE.getParent());
            Files.writeString(STORE, site.name(), StandardCharsets.UTF_8);
        } catch (Exception ignored) {
        }
    }
}
