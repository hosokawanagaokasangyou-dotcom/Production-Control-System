package jp.co.pm.ai.desktop.config;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;

/**
 * グローバル設定（init_setting への書き出し・「デフォルトに戻す」時の参照）で使う工場。
 *
 * <p>値は {@code ~/.pm-ai-desktop/global-init-setting-target-factory.txt} に {@link FactorySite#name()}（{@code KONAN} /
 * {@code KOKUBU}）で保存する。環境変数の工場プリセット適用時も同ファイルを同期し、意図しない工場の既定を読まないようにする。
 */
public final class GlobalInitSettingTarget {

    private GlobalInitSettingTarget() {}

    private static Path storePath() {
        return AppPaths.resolveDesktopAppHomeDir().resolve("global-init-setting-target-factory.txt");
    }

    public static Path storePathForTests() {
        return storePath();
    }

    /** 未設定・不正時は {@link FactorySite#KONAN}。 */
    public static FactorySite load() {
        try {
            if (!Files.isRegularFile(storePath())) {
                return FactorySite.KONAN;
            }
            String raw = Files.readString(storePath(), StandardCharsets.UTF_8).trim();
            if (raw.isEmpty()) {
                return FactorySite.KONAN;
            }
            return FactorySite.valueOf(raw);
        } catch (Exception ignored) {
            return FactorySite.KONAN;
        }
    }

    /**
     * 環境変数タブの工場別 UNC から推定した工場を優先し、{@link #load()} の永続ファイルと同期する。
     *
     * <p>推定できない／同点のときは {@link #load()}（global-init-setting-target-factory.txt）を使う。
     *
     * <p>工場切替中（{@link #setSuppressUiEnvInferencePersist(boolean)}）は推定による {@link #save} と
     * 推定戻り値を抑止し、明示保存された {@link #load()} のみを返す。
     */
    public static FactorySite loadEffective(Map<String, String> ui) {
        if (suppressUiEnvInferencePersist.get()) {
            return load();
        }
        Optional<FactorySite> inferred = FactorySite.inferFromUiEnv(ui);
        if (inferred.isPresent()) {
            FactorySite site = inferred.get();
            if (load() != site) {
                save(site);
            }
            return site;
        }
        return load();
    }

    /**
     * 推定のみ行い永続ファイルへ {@link #save} しない（起動スプラッシュ等の表示専用）。
     */
    public static FactorySite peekEffective(Map<String, String> ui) {
        if (suppressUiEnvInferencePersist.get()) {
            return load();
        }
        return FactorySite.inferFromUiEnv(ui).orElseGet(GlobalInitSettingTarget::load);
    }

    /** 工場切替中に {@link #loadEffective} が旧工場 UNC から永続工場を巻き戻さないようにする。 */
    public static void setSuppressUiEnvInferencePersist(boolean suppress) {
        suppressUiEnvInferencePersist.set(suppress);
    }

    private static final ThreadLocal<Boolean> suppressUiEnvInferencePersist =
            ThreadLocal.withInitial(() -> Boolean.FALSE);

    public static void save(FactorySite site) {
        if (site == null) {
            return;
        }
        try {
            Files.createDirectories(storePath().getParent());
            Files.writeString(storePath(), site.name(), StandardCharsets.UTF_8);
        } catch (Exception ignored) {
        }
    }
}
