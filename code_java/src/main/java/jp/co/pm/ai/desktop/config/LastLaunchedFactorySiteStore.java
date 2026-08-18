package jp.co.pm.ai.desktop.config;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 前回この PC で起動していた利用工場（{@code ~/.pm-ai-desktop/last-launched-factory-site.json}）。
 *
 * <p>起動スプラッシュの工場バッジ判定の正本。ファイルが無い・不正なときは湖南工場。
 */
public final class LastLaunchedFactorySiteStore {

    public static final String FILE_NAME = "last-launched-factory-site.json";

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private LastLaunchedFactorySiteStore() {}

    public static Path storePath() {
        return AppPaths.resolveLastLaunchedFactorySiteStorePath();
    }

    public static Path storePathForTests() {
        return storePath();
    }

    /** テスト用にファイルを削除する。 */
    public static void resetForTests() {
        try {
            Files.deleteIfExists(storePath());
        } catch (Exception ignored) {
        }
    }

    /**
     * 保存済みの工場を返す。ファイル無し・破損・配台対象外（RDP 等）は {@link FactorySite#KONAN}。
     */
    public static FactorySite load() {
        Path path = storePath();
        if (!Files.isRegularFile(path)) {
            return FactorySite.KONAN;
        }
        try {
            Payload payload = MAPPER.readValue(Files.readString(path, StandardCharsets.UTF_8), Payload.class);
            return parseProductionSite(payload != null ? payload.factorySite() : null);
        } catch (Exception ignored) {
            return FactorySite.KONAN;
        }
    }

    /** 配台対象工場のみ保存する。{@code null} / RDP ランチャーは無視する。 */
    public static void save(FactorySite site) {
        if (site == null || !FactorySite.dispatchProductionSites().contains(site)) {
            return;
        }
        try {
            Path path = storePath();
            Files.createDirectories(path.getParent());
            MAPPER.writerWithDefaultPrettyPrinter().writeValue(path.toFile(), new Payload(site.name()));
        } catch (Exception ignored) {
        }
    }

    private static FactorySite parseProductionSite(String raw) {
        if (raw == null || raw.isBlank()) {
            return FactorySite.KONAN;
        }
        try {
            FactorySite site = FactorySite.valueOf(raw.trim());
            if (FactorySite.dispatchProductionSites().contains(site)) {
                return site;
            }
        } catch (IllegalArgumentException ignored) {
        }
        return FactorySite.KONAN;
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public record Payload(String factorySite) {}
}
