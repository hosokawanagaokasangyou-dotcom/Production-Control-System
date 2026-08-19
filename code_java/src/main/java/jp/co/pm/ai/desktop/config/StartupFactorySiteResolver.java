package jp.co.pm.ai.desktop.config;

import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;

/**
 * 起動時の利用工場。スプラッシュと本番で同じ正本を使う。
 *
 * <p>正本は {@link LastLaunchedFactorySiteStore}（前回起動時に書き出した JSON）。ファイルが無い初回は湖南工場。
 */
public final class StartupFactorySiteResolver {

    private StartupFactorySiteResolver() {}

    /** スプラッシュ・本番共通の採用工場。この呼び出しでは永続ファイルを変更しない。 */
    public static FactorySite resolve() {
        return LastLaunchedFactorySiteStore.load();
    }

    /**
     * ポータブル自動バージョンアップで使う工場。前回起動工場を環境タブ UNC 推定より優先する。
     *
     * <p>起動直後は環境タブに別工場 UNC が残ることがあり、推定で湖南へ巻き戻さない。
     */
    public static FactorySite resolveForPortableUpgrade(
            Optional<FactorySite> followUp, Map<String, String> ui, Optional<Path> canonical) {
        if (followUp != null && followUp.isPresent()) {
            FactorySite fromFollowUp = followUp.get();
            if (fromFollowUp != null && fromFollowUp != FactorySite.RDP_LAUNCHER) {
                return fromFollowUp;
            }
        }
        FactorySite launched = resolve();
        if (launched != null && launched != FactorySite.RDP_LAUNCHER) {
            return launched;
        }
        Optional<FactorySite> fromEnv = FactorySite.inferFromUiEnv(ui != null ? ui : Map.of());
        if (fromEnv.isPresent() && fromEnv.get() != FactorySite.RDP_LAUNCHER) {
            return fromEnv.get();
        }
        FactorySite stored = GlobalInitSettingTarget.load();
        if (stored != null && stored != FactorySite.RDP_LAUNCHER) {
            return stored;
        }
        if (canonical != null && canonical.isPresent()) {
            Optional<FactorySite> fromCanonical =
                    FactorySite.inferFromPortableBundleInitSetting(canonical.get());
            if (fromCanonical.isEmpty()) {
                fromCanonical =
                        FactorySite.inferFromPortableBundleSourceValue(canonical.get().toString());
            }
            if (fromCanonical.isPresent()) {
                return fromCanonical.get();
            }
        }
        return FactorySite.KONAN;
    }

    /** スプラッシュの工場バッジ・テーマ用。{@link #resolve()} と同じ。 */
    public static FactorySite resolveForSplash() {
        return resolve();
    }

    /**
     * ディスク上の永続工場が採用工場と異なるとき、起動時に工場切替（ワークスペース復元）が必要。
     */
    public static boolean requiresStartupSwitch(FactorySite persisted, FactorySite adopted) {
        if (adopted == null || adopted == FactorySite.RDP_LAUNCHER) {
            return false;
        }
        if (persisted == null || persisted == FactorySite.RDP_LAUNCHER) {
            return true;
        }
        return persisted != adopted;
    }
}
