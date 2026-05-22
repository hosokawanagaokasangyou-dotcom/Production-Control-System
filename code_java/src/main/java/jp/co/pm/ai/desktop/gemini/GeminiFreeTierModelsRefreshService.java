package jp.co.pm.ai.desktop.gemini;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.Supplier;

import jp.co.pm.ai.desktop.benchmark.GeminiModelsListRestClient;
import jp.co.pm.ai.desktop.benchmark.GeminiModelsListRestClient.ListedModel;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.GeminiDispatchModelTryOrderDefaults;
import jp.co.pm.ai.desktop.crypto.GeminiCredentialsV2Crypto;

/**
 * Gemini 無料枠向け Flash-Lite モデル一覧の日次バックグラウンド更新（{@code models.list}）。
 */
public final class GeminiFreeTierModelsRefreshService {

    private static final Duration REFRESH_INTERVAL = Duration.ofDays(1);
    private static final Duration HTTP_TIMEOUT = Duration.ofSeconds(90);
    private static final String LIST_SOURCE = "generativelanguage.googleapis.com/v1beta/models";

    private static final DateTimeFormatter STAMP =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")
                    .withZone(ZoneId.systemDefault());

    private final Supplier<Map<String, String>> uiEnvSupplier;
    private final Listener listener;
    private final ScheduledExecutorService scheduler;
    private final AtomicBoolean refreshInFlight = new AtomicBoolean(false);
    private final AtomicBoolean started = new AtomicBoolean(false);

    public GeminiFreeTierModelsRefreshService(
            Supplier<Map<String, String>> uiEnvSupplier, Listener listener) {
        this.uiEnvSupplier = Objects.requireNonNull(uiEnvSupplier, "uiEnvSupplier");
        this.listener = Objects.requireNonNull(listener, "listener");
        this.scheduler =
                Executors.newSingleThreadScheduledExecutor(
                        r -> {
                            Thread t = new Thread(r, "pm-ai-gemini-free-tier-models");
                            t.setDaemon(true);
                            return t;
                        });
    }

    public interface Listener {
        /** バックグラウンドスレッドから呼ばれる。UI 更新は実装側で {@code Platform.runLater} すること。 */
        void onRefreshFinished(RefreshResult result);
    }

    public record RefreshResult(
            boolean success,
            boolean manual,
            List<String> modelIds,
            String message,
            long refreshedAtEpochMillis,
            Path cachePath) {

        public RefreshResult {
            modelIds = modelIds != null ? List.copyOf(modelIds) : List.of();
        }
    }

    /** 起動後1回だけ呼ぶ。24時間ごとに更新し、キャッシュが古ければ初回も即時更新する。 */
    public void start() {
        if (!started.compareAndSet(false, true)) {
            return;
        }
        long periodSec = Math.max(60L, REFRESH_INTERVAL.getSeconds());
        scheduler.scheduleAtFixedRate(
                () -> refreshIfDue(false), periodSec, periodSec, TimeUnit.SECONDS);
        scheduler.execute(() -> refreshIfDue(false));
    }

    public void shutdown() {
        scheduler.shutdown();
    }

    /** 手動「強制更新」用。 */
    public void refreshNow(boolean manual) {
        scheduler.execute(() -> runRefresh(manual));
    }

    private void refreshIfDue(boolean manual) {
        Map<String, String> ui = uiEnvSupplier.get();
        Path cachePath = GeminiFreeTierModelsCache.resolvePath(ui);
        OptionalSnapshot due =
                GeminiFreeTierModelsCache.read(cachePath)
                        .map(s -> new OptionalSnapshot(s, GeminiFreeTierModelsCache.isStale(s, REFRESH_INTERVAL)))
                        .orElse(new OptionalSnapshot(null, true));
        if (!manual && !due.stale()) {
            return;
        }
        runRefresh(manual);
    }

    private record OptionalSnapshot(GeminiFreeTierModelsCache.Snapshot snapshot, boolean stale) {}

    private void runRefresh(boolean manual) {
        if (!refreshInFlight.compareAndSet(false, true)) {
            listener.onRefreshFinished(
                    new RefreshResult(
                            false,
                            manual,
                            List.of(),
                            "別の更新が実行中です。",
                            0L,
                            null));
            return;
        }
        try {
            Map<String, String> ui = uiEnvSupplier.get();
            Path cachePath = GeminiFreeTierModelsCache.resolvePath(ui);
            RefreshResult result = doRefresh(ui, cachePath, manual);
            listener.onRefreshFinished(result);
        } finally {
            refreshInFlight.set(false);
        }
    }

    private RefreshResult doRefresh(Map<String, String> ui, Path cachePath, boolean manual) {
        String apiKey;
        try {
            apiKey = loadApiKey(ui);
        } catch (Exception ex) {
            return failure(cachePath, manual, "API キー取得失敗: " + ex.getMessage());
        }
        if (apiKey == null || apiKey.isBlank()) {
            return failure(cachePath, manual, "Gemini 認証ファイルが未設定または復号できません。");
        }
        List<String> modelIds;
        try {
            List<ListedModel> listed =
                    GeminiModelsListRestClient.listAllModels(apiKey, HTTP_TIMEOUT);
            modelIds = GeminiFreeTierModelSelector.selectFlashLiteGenerateContentModels(listed);
        } catch (IOException | InterruptedException ex) {
            if (ex instanceof InterruptedException) {
                Thread.currentThread().interrupt();
            }
            return failure(cachePath, manual, "models.list 失敗: " + ex.getMessage());
        }
        if (modelIds.isEmpty()) {
            modelIds = List.copyOf(GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_FALLBACK_TRY_ORDER);
        }
        long now = System.currentTimeMillis();
        GeminiFreeTierModelsCache.Snapshot snap =
                new GeminiFreeTierModelsCache.Snapshot(now, modelIds, null, LIST_SOURCE);
        try {
            GeminiFreeTierModelsCache.write(cachePath, snap);
        } catch (IOException ex) {
            return failure(cachePath, manual, "キャッシュ書き込み失敗: " + ex.getMessage());
        }
        String msg =
                (manual ? "手動更新" : "日次更新")
                        + ": Flash-Lite "
                        + modelIds.size()
                        + " 件（"
                        + STAMP.format(Instant.ofEpochMilli(now))
                        + "）→ "
                        + cachePath.getFileName();
        return new RefreshResult(true, manual, modelIds, msg, now, cachePath);
    }

    private static RefreshResult failure(Path cachePath, boolean manual, String message) {
        long now = System.currentTimeMillis();
        if (cachePath != null) {
            try {
                GeminiFreeTierModelsCache.Snapshot prev = GeminiFreeTierModelsCache.read(cachePath).orElse(null);
                List<String> keep =
                        prev != null && prev.hasModels()
                                ? prev.modelIds()
                                : List.copyOf(
                                        GeminiDispatchModelTryOrderDefaults
                                                .PLANNING_CORE_FALLBACK_TRY_ORDER);
                GeminiFreeTierModelsCache.write(
                        cachePath,
                        new GeminiFreeTierModelsCache.Snapshot(now, keep, message, LIST_SOURCE));
            } catch (IOException ignored) {
                // keep previous file
            }
        }
        return new RefreshResult(false, manual, List.of(), message, now, cachePath);
    }

    private static String loadApiKey(Map<String, String> ui) throws Exception {
        Path credPath = resolveGeminiCredentialsPath(ui);
        if (!Files.isRegularFile(credPath)) {
            return null;
        }
        String json = Files.readString(credPath);
        return GeminiCredentialsV2Crypto.decryptGeminiApiKeyFromJsonString(
                json, GeminiCredentialsV2Crypto.DEFAULT_PASSPHRASE);
    }

    private static Path resolveGeminiCredentialsPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String raw = u.get(AppPaths.KEY_GEMINI_CREDENTIALS_JSON);
        if (raw != null && !raw.isBlank()) {
            return Path.of(raw.strip()).toAbsolutePath().normalize();
        }
        return AppPaths.resolveRepoRoot(u)
                .resolve("code")
                .resolve("gemini_credentials.encrypted.json")
                .toAbsolutePath()
                .normalize();
    }
}
