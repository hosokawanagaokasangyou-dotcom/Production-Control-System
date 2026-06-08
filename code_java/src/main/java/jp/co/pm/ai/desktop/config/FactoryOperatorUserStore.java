package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.security.GeneralSecurityException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.security.SecureRandom;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HexFormat;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.crypto.AladdinOperatorCredentialsCrypto;
import jp.co.pm.ai.desktop.io.FactoryOperatorUserBackupStore;
import jp.co.pm.ai.desktop.io.OperatorAladdinCredentialsLauncherJson;

/**
 * 工場別の配台システム操作者名（起動時選択・作成者表示用）と PIN（4～10 桁数字）。
 *
 * <p>操作者一覧・PIN の永続化: {@link AppPaths#factoryOperatorUsersStorePath}（サマリ Excel と同一フォルダの
 * {@link AppPaths#FACTORY_OPERATOR_USERS_BIN}、バイナリ形式）。旧 {@code ~/.pm-ai-desktop/factory-operator-users.json}
 * から初回読込時に移行する。
 *
 * <p>最後に選択した操作者名（{@code lastSelected}）は PC ローカルの
 * {@link AppPaths#localFactoryOperatorLastSelectedPath} のみ。共有 bin には書かない。
 */
public final class FactoryOperatorUserStore {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final SecureRandom SECURE_RANDOM = new SecureRandom();
    private static final HexFormat HEX = HexFormat.of();

    private static final byte[] BINARY_MAGIC = {'P', 'M', 'O', 'U'};
    private static final int BINARY_FORMAT_VERSION = 1;

    private static final Path LEGACY_JSON_STORE =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "factory-operator-users.json");

    private static volatile Path configuredStorePath;
    private static volatile Path configuredNetworkStorePath;
    private static volatile boolean storeConfigured;
    private static volatile boolean usingLocalStoreFallback;

    public static final int SCHEMA_VERSION = 6;

    private static final int LEGACY_SCHEMA_BEFORE_ALADDIN = 5;

    private static final java.util.Set<String> SCHEMA_UPGRADE_BACKUP_PATHS =
            java.util.Collections.synchronizedSet(new java.util.LinkedHashSet<>());
    public static final int MAX_NAMES_PER_FACTORY = 50;
    public static final int MAX_NAME_LENGTH = 40;
    public static final int MIN_PIN_LENGTH = 4;
    public static final int MAX_PIN_LENGTH = 10;

    /** @deprecated {@link #MIN_PIN_LENGTH} を使用 */
    @Deprecated
    public static final int PIN_LENGTH = MIN_PIN_LENGTH;
    public static final int MAX_CONSECUTIVE_PIN_FAILURES = 20;

    /** ユーザー管理者タブを開くための管理者ユーザー名（平文）。 */
    public static final String ADMIN_TAB_USERNAME = "Administrator";

    /** ユーザー管理者タブを開くための管理者パスワード（平文）。 */
    public static final String ADMIN_TAB_PASSWORD = "nagaoka123";

    /**
     * ログイン専用のゲスト操作者名（ユーザー一覧には含めない）。PIN 不要・サマリ Excel 生成不可。
     */
    public static final String GUEST_OPERATOR_NAME = "ゲスト";

    public static final List<String> DEFAULT_NAMES =
            List.of("砂田", "古家", "図司", "細川");

    private static volatile String sessionOperatorName = "";

    public enum PinVerificationResult {
        SUCCESS,
        NO_PIN_REQUIRED,
        WRONG_PIN,
        LOCKED,
        INVALID_PIN
    }

    /** 操作者別アラジン RPA ログイン資格情報。 */
    public record AladdinCredentials(String loginId, String password) {}

    public record FactoryOperatorUsers(
            List<String> names,
            String lastSelected,
            Map<String, String> pinHashes,
            Map<String, Integer> pinFailedAttempts,
            Set<String> pinMustChange,
            Map<String, String> pinPlaintextAdmin,
            Map<String, String> aladdinLoginIds,
            Map<String, String> aladdinPasswordCiphertext) {

        public FactoryOperatorUsers {
            names = names != null ? List.copyOf(names) : List.of();
            lastSelected = lastSelected != null ? lastSelected.strip() : "";
            pinHashes = pinHashes != null ? Map.copyOf(pinHashes) : Map.of();
            pinFailedAttempts = pinFailedAttempts != null ? Map.copyOf(pinFailedAttempts) : Map.of();
            pinMustChange = pinMustChange != null ? Set.copyOf(pinMustChange) : Set.of();
            pinPlaintextAdmin = pinPlaintextAdmin != null ? Map.copyOf(pinPlaintextAdmin) : Map.of();
            aladdinLoginIds = aladdinLoginIds != null ? Map.copyOf(aladdinLoginIds) : Map.of();
            aladdinPasswordCiphertext =
                    aladdinPasswordCiphertext != null ? Map.copyOf(aladdinPasswordCiphertext) : Map.of();
        }

        public FactoryOperatorUsers(List<String> names, String lastSelected) {
            this(names, lastSelected, Map.of(), Map.of(), Set.of(), Map.of(), Map.of(), Map.of());
        }

        public FactoryOperatorUsers(List<String> names, String lastSelected, Map<String, String> pinHashes) {
            this(names, lastSelected, pinHashes, Map.of(), Set.of(), Map.of(), Map.of(), Map.of());
        }

        public FactoryOperatorUsers(
                List<String> names,
                String lastSelected,
                Map<String, String> pinHashes,
                Map<String, Integer> pinFailedAttempts) {
            this(names, lastSelected, pinHashes, pinFailedAttempts, Set.of(), Map.of(), Map.of(), Map.of());
        }

        public FactoryOperatorUsers(
                List<String> names,
                String lastSelected,
                Map<String, String> pinHashes,
                Map<String, Integer> pinFailedAttempts,
                Set<String> pinMustChange) {
            this(names, lastSelected, pinHashes, pinFailedAttempts, pinMustChange, Map.of(), Map.of(), Map.of());
        }

        public FactoryOperatorUsers(
                List<String> names,
                String lastSelected,
                Map<String, String> pinHashes,
                Map<String, Integer> pinFailedAttempts,
                Set<String> pinMustChange,
                Map<String, String> pinPlaintextAdmin) {
            this(
                    names,
                    lastSelected,
                    pinHashes,
                    pinFailedAttempts,
                    pinMustChange,
                    pinPlaintextAdmin,
                    Map.of(),
                    Map.of());
        }
    }

    private FactoryOperatorUserStore() {}

    /** {@link AppPaths#summaryAiDispatchXlsxPath} と同じフォルダへストアパスを解決する。 */
    public static synchronized void configureFromUi(Map<String, String> ui) {
        configureFromUi(ui, null);
    }

    /**
     * 利用工場に合わせた bin パスへストアを解決する。
     *
     * <p>サマリ Excel 環境変数が別工場を指すときも {@code site} の DATA フォルダ側を使う。
     */
    public static synchronized void configureFromUi(Map<String, String> ui, FactorySite site) {
        Map<String, String> u = ui != null ? ui : Map.of();
        FactorySite effective = site != null ? site : GlobalInitSettingTarget.load();
        Path network = AppPaths.factoryOperatorUsersStorePath(u, effective);
        Path local = AppPaths.localFactoryOperatorUsersStorePath(effective);
        Path next = resolveWritableStorePath(network, local);
        if (storeConfigured
                && next.equals(configuredStorePath)
                && Objects.equals(network, configuredNetworkStorePath)) {
            return;
        }
        configuredStorePath = next;
        configuredNetworkStorePath = network;
        usingLocalStoreFallback = next.equals(local);
        storeConfigured = true;
        if (usingLocalStoreFallback) {
            seedLocalStoreFromNetworkIfNeeded(network, local);
        }
    }

    /** ネットワーク共有ではなくローカル退避を使っているとき true。 */
    public static boolean usingLocalStoreFallback() {
        return usingLocalStoreFallback;
    }

    /** 工場別 UNC 上の正本パス（アクセス可否は問わない）。 */
    public static Path networkStorePath() {
        return configuredNetworkStorePath;
    }

    public static Path storePath() {
        String test = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        if (test != null && !test.isBlank()) {
            return Path.of(test).toAbsolutePath().normalize();
        }
        if (storeConfigured && configuredStorePath != null) {
            return configuredStorePath;
        }
        return AppPaths.factoryOperatorUsersStorePath(Map.of());
    }

    /** 現在セッションで選択中の操作者名（起動時に設定）。 */
    public static String sessionOperatorName() {
        return sessionOperatorName != null ? sessionOperatorName.strip() : "";
    }

    public static void clearSessionOperatorName() {
        sessionOperatorName = "";
    }

    public static boolean isGuestOperator(String name) {
        return GUEST_OPERATOR_NAME.equals(normalizeName(name));
    }

    public static boolean isGuestSession() {
        return isGuestOperator(sessionOperatorName());
    }

    /** 現在セッションがサマリ Excel の生成・上書きを行えるか。 */
    public static boolean sessionMayGenerateSummaryExcel() {
        return !isGuestSession();
    }

    /** 依頼書入力の転記・一時保存・設定変更を行えるか。 */
    public static boolean sessionMayMutateRequestFormInput() {
        return !isGuestSession();
    }

    /** 起動時ログイン選択肢（登録操作者＋ゲスト）。 */
    public static List<String> loginChoicesForFactory(FactorySite site) throws IOException {
        List<String> choices = new ArrayList<>(namesForFactory(site));
        if (!choices.contains(GUEST_OPERATOR_NAME)) {
            choices.add(GUEST_OPERATOR_NAME);
        }
        return choices;
    }

    public static List<String> namesForFactory(FactorySite site) throws IOException {
        return new ArrayList<>(loadFactory(site).names());
    }

    public static String lastSelectedForFactory(FactorySite site) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        migrateLastSelectedFromSharedStoreIfNeeded(factory);
        return loadLastSelectedLocal(factory);
    }

    /**
     * PC ローカルに保存した最終操作者でセッションを復元する。
     *
     * <p>一覧に無い・PIN ロック・初回 PIN 変更待ちのときは false。PIN 設定済みでも同一 PC では再入力しない。
     */
    public static boolean tryRestoreSessionFromLocalLastSelected(FactorySite site) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String last = lastSelectedForFactory(factory);
        if (last.isEmpty()) {
            return false;
        }
        if (!loginChoicesForFactory(factory).contains(last)) {
            return false;
        }
        if (!isGuestOperator(last)) {
            if (isPinLocked(factory, last)) {
                return false;
            }
            if (mustChangePin(factory, last)) {
                return false;
            }
            if (!loadFactory(factory).names().contains(last)) {
                return false;
            }
        }
        selectSessionOperator(factory, last);
        return true;
    }

    public static boolean hasPin(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty() || isGuestOperator(normalized)) {
            return false;
        }
        String hash = loadFactory(factory).pinHashes().get(normalized);
        return hash != null && !hash.isBlank();
    }

    /** 初回ログイン後の PIN 変更が未完了（管理者発行／新規追加のランダム PIN）。 */
    public static boolean mustChangePin(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            return false;
        }
        return loadFactory(factory).pinMustChange().contains(normalized);
    }

    /** 管理者向け: 記録されている PIN 平文（発行・変更後のみ。旧データは空）。 */
    public static Optional<String> adminViewablePin(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            return Optional.empty();
        }
        String pin = loadFactory(factory).pinPlaintextAdmin().get(normalized);
        if (pin == null || pin.isBlank()) {
            return Optional.empty();
        }
        return Optional.of(pin);
    }

    public static boolean isPinLocked(FactorySite site, String name) throws IOException {
        return pinFailedAttemptCount(site, name) >= MAX_CONSECUTIVE_PIN_FAILURES;
    }

    public static int pinFailedAttemptCount(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            return 0;
        }
        Integer count = loadFactory(factory).pinFailedAttempts().get(normalized);
        return count != null ? Math.max(0, count) : 0;
    }

    public static int remainingPinAttempts(FactorySite site, String name) throws IOException {
        int failed = pinFailedAttemptCount(site, name);
        if (failed >= MAX_CONSECUTIVE_PIN_FAILURES) {
            return 0;
        }
        return MAX_CONSECUTIVE_PIN_FAILURES - failed;
    }

    /** 副作用なしの PIN 照合（失敗回数は増やさない）。 */
    public static boolean verifyPin(FactorySite site, String name, String pin) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        String pinNorm = normalizePin(pin);
        if (normalized.isEmpty() || pinNorm == null) {
            return false;
        }
        String expected = loadFactory(factory).pinHashes().get(normalized);
        if (expected == null || expected.isBlank()) {
            return true;
        }
        return MessageDigest.isEqual(
                expected.getBytes(StandardCharsets.UTF_8),
                hashPin(factory, normalized, pinNorm).getBytes(StandardCharsets.UTF_8));
    }

    /**
     * 起動時 PIN 入力向け。連続失敗を記録し、{@link #MAX_CONSECUTIVE_PIN_FAILURES} 回でロックする。
     */
    public static PinVerificationResult verifyPinAttempt(FactorySite site, String name, String pin)
            throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            return PinVerificationResult.INVALID_PIN;
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (!current.names().contains(normalized)) {
            return PinVerificationResult.INVALID_PIN;
        }
        int failures = current.pinFailedAttempts().getOrDefault(normalized, 0);
        if (failures >= MAX_CONSECUTIVE_PIN_FAILURES) {
            return PinVerificationResult.LOCKED;
        }
        String expected = current.pinHashes().get(normalized);
        if (expected == null || expected.isBlank()) {
            return PinVerificationResult.NO_PIN_REQUIRED;
        }
        String pinNorm = normalizePin(pin);
        if (pinNorm == null) {
            return PinVerificationResult.INVALID_PIN;
        }
        if (MessageDigest.isEqual(
                expected.getBytes(StandardCharsets.UTF_8),
                hashPin(factory, normalized, pinNorm).getBytes(StandardCharsets.UTF_8))) {
            clearPinFailures(doc, factory, normalized);
            saveDocument(doc);
            return PinVerificationResult.SUCCESS;
        }
        int nextFailures = failures + 1;
        Map<String, Integer> attempts = new LinkedHashMap<>(current.pinFailedAttempts());
        attempts.put(normalized, nextFailures);
        doc.factories()
                .put(
                        factory,
                        forSharedStore(
                                current,
                                current.names(),
                                current.pinHashes(),
                                attempts,
                                current.pinMustChange(),
                                current.pinPlaintextAdmin()));
        saveDocument(doc);
        return nextFailures >= MAX_CONSECUTIVE_PIN_FAILURES
                ? PinVerificationResult.LOCKED
                : PinVerificationResult.WRONG_PIN;
    }

    /** 管理者が PIN ロックを解除する（PIN 自体は変更しない）。 */
    public static void unlockPin(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("名前が空です。");
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (!current.names().contains(normalized)) {
            throw new IllegalArgumentException("操作者名が一覧にありません: " + normalized);
        }
        clearPinFailures(doc, factory, normalized);
        saveDocument(doc);
    }

    public static String pinLengthRangeDescriptionJa() {
        return MIN_PIN_LENGTH + "～" + MAX_PIN_LENGTH + "桁の数字";
    }

    /** ユーザー管理者タブ解錠: ユーザー名 {@link #ADMIN_TAB_USERNAME} とパスワードを照合する。 */
    public static boolean verifyAdminTabAccess(String username, String password) {
        String user = username != null ? username.strip() : "";
        String pass = password != null ? password : "";
        return ADMIN_TAB_USERNAME.equals(user) && ADMIN_TAB_PASSWORD.equals(pass);
    }

    /**
     * ログイン中の操作者が自分の PIN を変更する（未設定のときは新規設定）。
     *
     * @throws IllegalStateException 自分以外／ロック中
     * @throws IllegalArgumentException PIN 形式不正・現在 PIN 不一致
     */
    public static void changePinByUser(
            FactorySite site, String name, String currentPin, String newPin) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("名前が空です。");
        }
        if (!normalized.equals(sessionOperatorName())) {
            throw new IllegalStateException("自分の PIN のみ変更できます。");
        }
        if (isPinLocked(factory, normalized)) {
            throw new IllegalStateException(
                    "PIN がロックされています。ユーザー管理者タブでロック解除してください。");
        }
        String newPinNorm = normalizePin(newPin);
        if (newPinNorm == null) {
            throw new IllegalArgumentException("新しい PIN は " + pinLengthRangeDescriptionJa() + " です。");
        }
        if (hasPin(factory, normalized)) {
            String currentNorm = normalizePin(currentPin);
            if (currentNorm == null || !verifyPin(factory, normalized, currentNorm)) {
                throw new IllegalArgumentException("現在の PIN が正しくありません。");
            }
        }
        assignPin(factory, normalized, newPinNorm, false);
    }

    /**
     * 初回ログイン時の強制 PIN 変更（ランダム初期 PIN 入力後）。セッション未設定でも呼べる。
     *
     * @throws IllegalStateException 初回変更フラグが無い
     * @throws IllegalArgumentException PIN 形式不正・現在 PIN 不一致
     */
    public static void changePinOnFirstLogin(
            FactorySite site, String name, String currentPin, String newPin) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("名前が空です。");
        }
        if (!mustChangePin(factory, normalized)) {
            throw new IllegalStateException("初回 PIN 変更は不要です。");
        }
        if (isPinLocked(factory, normalized)) {
            throw new IllegalStateException(
                    "PIN がロックされています。ユーザー管理者タブでロック解除してください。");
        }
        String newPinNorm = normalizePin(newPin);
        if (newPinNorm == null) {
            throw new IllegalArgumentException("新しい PIN は " + pinLengthRangeDescriptionJa() + " です。");
        }
        String currentNorm = normalizePin(currentPin);
        if (currentNorm == null || !verifyPin(factory, normalized, currentNorm)) {
            throw new IllegalArgumentException("現在の PIN が正しくありません。");
        }
        assignPin(factory, normalized, newPinNorm, false);
    }

    /**
     * 管理者向け: ランダム 4 桁 PIN を新規発行する（既存 PIN があれば上書き）。
     *
     * @return 発行した PIN（管理者タブからいつでも {@link #adminViewablePin} で閲覧可能）
     */
    public static String issuePin(FactorySite site, String name) throws IOException {
        return assignPin(site, name, generatePin(), true);
    }

    /**
     * 管理者向け: {@link #issuePin} と同じだが意図を明示する別名。
     */
    public static String reissuePin(FactorySite site, String name) throws IOException {
        return issuePin(site, name);
    }

    /**
     * 管理者向け: 指定した PIN を手動で設定する（初回ログイン時の変更は不要）。
     *
     * @return 設定した PIN（正規化後）
     */
    public static String assignPinByAdmin(FactorySite site, String name, String pin) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (isGuestOperator(normalized)) {
            throw new IllegalArgumentException("ゲストユーザーには PIN を設定できません。");
        }
        return assignPin(factory, normalized, pin, false);
    }

    private static String assignPin(
            FactorySite site, String name, String pin, boolean requireChangeOnNextLogin) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        String pinNorm = normalizePin(pin);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("名前が空です。");
        }
        if (pinNorm == null) {
            throw new IllegalArgumentException("PIN は " + pinLengthRangeDescriptionJa() + " です。");
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (!current.names().contains(normalized)) {
            throw new IllegalArgumentException("操作者名が一覧にありません: " + normalized);
        }
        Map<String, String> pins = new LinkedHashMap<>(current.pinHashes());
        pins.put(normalized, hashPin(factory, normalized, pinNorm));
        Map<String, Integer> attempts = new LinkedHashMap<>(current.pinFailedAttempts());
        attempts.remove(normalized);
        Set<String> mustChange = new LinkedHashSet<>(current.pinMustChange());
        if (requireChangeOnNextLogin) {
            mustChange.add(normalized);
        } else {
            mustChange.remove(normalized);
        }
        Map<String, String> plaintextAdmin = new LinkedHashMap<>(current.pinPlaintextAdmin());
        plaintextAdmin.put(normalized, pinNorm);
        doc.factories()
                .put(
                        factory,
                        forSharedStore(current, current.names(), pins, attempts, mustChange, plaintextAdmin));
        saveDocument(doc);
        return pinNorm;
    }

    /**
     * セッション操作者を設定し、工場別の最終選択をローカルへ永続化する。
     *
     * @throws IllegalArgumentException 名前が一覧に無い／空
     */
    public static void selectSessionOperator(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("操作者名が空です。");
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (isGuestOperator(normalized)) {
            sessionOperatorName = normalized;
            saveLastSelectedLocal(factory, normalized);
            return;
        }
        if (!current.names().contains(normalized)) {
            throw new IllegalArgumentException("操作者名が一覧にありません: " + normalized);
        }
        sessionOperatorName = normalized;
        saveLastSelectedLocal(factory, normalized);
    }

    /**
     * 操作者名を追加し、ランダム 4 桁 PIN を発行する。
     *
     * @return 発行した PIN（初回ログイン時に変更必須。管理者のみこのタイミングで平文表示）
     */
    public static String addName(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("名前が空です。");
        }
        if (isGuestOperator(normalized)) {
            throw new IllegalArgumentException("「" + GUEST_OPERATOR_NAME + "」は予約された操作者名です。");
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (current.names().contains(normalized)) {
            throw new IllegalArgumentException("同じ名前が既にあります: " + normalized);
        }
        if (current.names().size() >= MAX_NAMES_PER_FACTORY) {
            throw new IllegalArgumentException("名前は工場あたり最大 " + MAX_NAMES_PER_FACTORY + " 件です。");
        }
        List<String> next = new ArrayList<>(current.names());
        next.add(normalized);
        String pin = generatePin();
        Map<String, String> pins = new LinkedHashMap<>(current.pinHashes());
        pins.put(normalized, hashPin(factory, normalized, pin));
        Set<String> mustChange = new LinkedHashSet<>(current.pinMustChange());
        mustChange.add(normalized);
        Map<String, String> plaintextAdmin = new LinkedHashMap<>(current.pinPlaintextAdmin());
        plaintextAdmin.put(normalized, pin);
        doc.factories()
                .put(
                        factory,
                        forSharedStore(
                                current,
                                next,
                                pins,
                                current.pinFailedAttempts(),
                                mustChange,
                                plaintextAdmin));
        saveDocument(doc);
        return pin;
    }

    public static void removeName(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (current.names().size() <= 1) {
            throw new IllegalStateException("最後の1件は削除できません。");
        }
        if (!current.names().contains(normalized)) {
            return;
        }
        List<String> next = new ArrayList<>(current.names());
        next.remove(normalized);
        if (normalized.equals(loadLastSelectedLocal(factory))) {
            clearLastSelectedLocal(factory);
        }
        Map<String, String> pins = new LinkedHashMap<>(current.pinHashes());
        pins.remove(normalized);
        Map<String, Integer> attempts = new LinkedHashMap<>(current.pinFailedAttempts());
        attempts.remove(normalized);
        Set<String> mustChange = new LinkedHashSet<>(current.pinMustChange());
        mustChange.remove(normalized);
        Map<String, String> plaintextAdmin = new LinkedHashMap<>(current.pinPlaintextAdmin());
        plaintextAdmin.remove(normalized);
        Map<String, String> aladdinIds = new LinkedHashMap<>(current.aladdinLoginIds());
        aladdinIds.remove(normalized);
        Map<String, String> aladdinPasswords = new LinkedHashMap<>(current.aladdinPasswordCiphertext());
        aladdinPasswords.remove(normalized);
        doc.factories()
                .put(
                        factory,
                        forSharedStore(
                                next,
                                pins,
                                attempts,
                                mustChange,
                                plaintextAdmin,
                                aladdinIds,
                                aladdinPasswords));
        if (normalized.equals(sessionOperatorName) && factory == GlobalInitSettingTarget.load()) {
            sessionOperatorName = "";
        }
        saveDocument(doc);
    }

    public static void resetNamesToDefaults(FactorySite site) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        String localLast = loadLastSelectedLocal(factory);
        if (!localLast.isEmpty()
                && !DEFAULT_NAMES.contains(localLast)
                && !isGuestOperator(localLast)) {
            clearLastSelectedLocal(factory);
        }
        Map<String, String> pins = new LinkedHashMap<>();
        Map<String, Integer> attempts = new LinkedHashMap<>();
        Set<String> mustChange = new LinkedHashSet<>();
        Map<String, String> plaintextAdmin = new LinkedHashMap<>();
        Map<String, String> aladdinIds = new LinkedHashMap<>();
        Map<String, String> aladdinPasswords = new LinkedHashMap<>();
        for (String n : DEFAULT_NAMES) {
            String h = current.pinHashes().get(n);
            if (h != null && !h.isBlank()) {
                pins.put(n, h);
            }
            Integer failed = current.pinFailedAttempts().get(n);
            if (failed != null && failed > 0) {
                attempts.put(n, failed);
            }
            if (current.pinMustChange().contains(n)) {
                mustChange.add(n);
            }
            String plain = current.pinPlaintextAdmin().get(n);
            if (plain != null && !plain.isBlank()) {
                plaintextAdmin.put(n, plain);
            }
            String aladdinId = current.aladdinLoginIds().get(n);
            if (aladdinId != null && !aladdinId.isBlank()) {
                aladdinIds.put(n, aladdinId);
            }
            String aladdinCipher = current.aladdinPasswordCiphertext().get(n);
            if (aladdinCipher != null && !aladdinCipher.isBlank()) {
                aladdinPasswords.put(n, aladdinCipher);
            }
        }
        doc.factories()
                .put(
                        factory,
                        forSharedStore(
                                DEFAULT_NAMES,
                                pins,
                                attempts,
                                mustChange,
                                plaintextAdmin,
                                aladdinIds,
                                aladdinPasswords));
        if (!DEFAULT_NAMES.contains(sessionOperatorName) && factory == GlobalInitSettingTarget.load()) {
            sessionOperatorName = "";
        }
        saveDocument(doc);
    }

    /** 管理者一覧表示用: PIN 平文または未記録の案内。 */
    public static String adminPinDisplayLabel(FactorySite site, String name) throws IOException {
        if (!hasPin(site, name)) {
            return "—";
        }
        return adminViewablePin(site, name).orElse("（再発行で確認）");
    }

    /** 一覧表示用: PIN / ロック状態。 */
    public static String pinStatusLabel(FactorySite site, String name) throws IOException {
        if (isPinLocked(site, name)) {
            return "ロック";
        }
        if (mustChangePin(site, name)) {
            return "初回変更待";
        }
        return hasPin(site, name) ? "設定済" : "未設定";
    }

    /** 当該操作者のアラジン ログイン ID（未設定なら空）。 */
    public static String aladdinLoginIdFor(FactorySite site, String name) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            return "";
        }
        String id = loadFactory(factory).aladdinLoginIds().get(normalized);
        return id != null ? id.strip() : "";
    }

    /** アラジン ID と復号可能パスワードの両方が設定済みか。 */
    public static boolean hasAladdinCredentials(FactorySite site, String name) throws IOException {
        return aladdinCredentialsFor(site, name).isPresent();
    }

    /** 当該操作者のアラジン資格情報（ID・パスワード両方あるときのみ）。 */
    public static Optional<AladdinCredentials> aladdinCredentialsFor(FactorySite site, String name)
            throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            return Optional.empty();
        }
        FactoryOperatorUsers current = loadFactory(factory);
        String loginId = current.aladdinLoginIds().get(normalized);
        String ciphertext = current.aladdinPasswordCiphertext().get(normalized);
        if (loginId == null || loginId.isBlank() || ciphertext == null || ciphertext.isBlank()) {
            return Optional.empty();
        }
        try {
            JsonNode payload = JSON.readTree(ciphertext);
            String password = AladdinOperatorCredentialsCrypto.decryptFromPayload(payload);
            if (password.isBlank()) {
                return Optional.empty();
            }
            return Optional.of(new AladdinCredentials(loginId.strip(), password));
        } catch (Exception ex) {
            return Optional.empty();
        }
    }

    /**
     * 操作者別アラジン RPA ログイン資格情報を保存する。
     *
     * @throws IllegalArgumentException 操作者名・ID・パスワードが不正
     */
    public static void setAladdinCredentials(
            FactorySite site, String name, String loginId, String password) throws IOException {
        FactorySite factory = site != null ? site : FactorySite.KONAN;
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("操作者名が空です。");
        }
        if (isGuestOperator(normalized)) {
            throw new IllegalArgumentException("ゲストにはアラジン資格情報を設定できません。");
        }
        String id = loginId != null ? loginId.strip() : "";
        if (id.isEmpty()) {
            throw new IllegalArgumentException("アラジン ログイン ID が空です。");
        }
        String pass = password != null ? password : "";
        if (pass.isBlank()) {
            throw new IllegalArgumentException("アラジン パスワードが空です。");
        }
        Document doc = loadDocument();
        FactoryOperatorUsers current = ensureFactory(doc, factory);
        if (!current.names().contains(normalized)) {
            throw new IllegalArgumentException("操作者名が一覧にありません: " + normalized);
        }
        ObjectNode payload;
        try {
            payload = AladdinOperatorCredentialsCrypto.encryptToPayload(pass);
        } catch (GeneralSecurityException ex) {
            throw new IOException("アラジン資格情報の暗号化に失敗しました。", ex);
        }
        Map<String, String> loginIds = new LinkedHashMap<>(current.aladdinLoginIds());
        Map<String, String> passwords = new LinkedHashMap<>(current.aladdinPasswordCiphertext());
        loginIds.put(normalized, id);
        passwords.put(normalized, JSON.writeValueAsString(payload));
        doc.factories()
                .put(
                        factory,
                        forSharedStore(
                                current.names(),
                                current.pinHashes(),
                                current.pinFailedAttempts(),
                                current.pinMustChange(),
                                current.pinPlaintextAdmin(),
                                loginIds,
                                passwords));
        saveDocument(doc);
    }

    private static FactoryOperatorUsers loadFactory(FactorySite site) throws IOException {
        return ensureFactory(loadDocument(), site != null ? site : FactorySite.KONAN);
    }

    private static Document loadDocument() throws IOException {
        Path path = storePath();
        migrateLegacyStoreIfNeeded(path);
        if (!Files.isRegularFile(path)) {
            Path network = configuredNetworkStorePath;
            if (network != null
                    && !network.equals(path)
                    && Files.isRegularFile(network)
                    && Files.isReadable(network)) {
                JsonNode root = readStoreRoot(network);
                if (root != null && root.isObject()) {
                    Document doc = parseDocumentRoot(root);
                    saveDocumentToPath(path, doc);
                    return doc;
                }
            }
            return defaultDocument();
        }
        JsonNode root = readStoreRoot(path);
        if (root == null || !root.isObject()) {
            return defaultDocument();
        }
        return parseDocumentRoot(root);
    }

    private static JsonNode readStoreRoot(Path path) throws IOException {
        byte[] bytes = Files.readAllBytes(path);
        if (isBinaryStore(bytes)) {
            return decodeBinaryPayload(bytes);
        }
        return JSON.readTree(bytes);
    }

    private static Document parseDocumentRoot(JsonNode root) throws IOException {
        int ver = root.path("schemaVersion").asInt(0);
        if (ver <= 0) {
            ver = 1;
        }
        if (ver > SCHEMA_VERSION) {
            throw new IOException(
                    "操作者名設定はより新しいアプリ向けです (schemaVersion=" + ver + ")");
        }
        Map<FactorySite, FactoryOperatorUsers> factories = new LinkedHashMap<>();
        JsonNode factoriesNode = root.get("factories");
        if (factoriesNode != null && factoriesNode.isObject()) {
            for (FactorySite site : FactorySite.values()) {
                JsonNode n = factoriesNode.get(site.name());
                if (n != null && n.isObject()) {
                    factories.put(site, parseFactory(n));
                }
            }
        }
        Document doc = new Document(ver, factories);
        for (FactorySite site : FactorySite.values()) {
            ensureFactory(doc, site);
        }
        return doc;
    }

    private static void migrateLegacyStoreIfNeeded(Path targetBin) throws IOException {
        if (Files.isRegularFile(targetBin)) {
            return;
        }
        Path legacySibling = targetBin.resolveSibling("factory-operator-users.json");
        for (Path legacy : List.of(resolveLegacyJsonStorePath(), legacySibling)) {
            if (!Files.isRegularFile(legacy)) {
                continue;
            }
            JsonNode root = JSON.readTree(legacy.toFile());
            Document doc = parseDocumentRoot(root);
            persistLegacyLastSelectedFromDocument(doc);
            saveDocumentToPath(targetBin, doc);
            return;
        }
    }

    /** 旧 JSON 移行時: 共有ストア保存前に {@code lastSelected} をローカルへ退避する。 */
    private static void persistLegacyLastSelectedFromDocument(Document doc) throws IOException {
        if (doc == null) {
            return;
        }
        for (Map.Entry<FactorySite, FactoryOperatorUsers> e : doc.factories().entrySet()) {
            String legacy = e.getValue().lastSelected();
            if (!legacy.isEmpty() && !Files.isRegularFile(localLastSelectedPath(e.getKey()))) {
                saveLastSelectedLocal(e.getKey(), legacy);
            }
        }
    }

    private static boolean isBinaryStore(byte[] bytes) {
        return bytes.length >= BINARY_MAGIC.length
                && Arrays.compare(
                                Arrays.copyOf(bytes, BINARY_MAGIC.length),
                                BINARY_MAGIC)
                        == 0;
    }

    private static JsonNode decodeBinaryPayload(byte[] fileBytes) throws IOException {
        ByteBuffer buf = ByteBuffer.wrap(fileBytes);
        byte[] magic = new byte[BINARY_MAGIC.length];
        buf.get(magic);
        if (!Arrays.equals(magic, BINARY_MAGIC)) {
            throw new IOException("操作者名設定ファイルの形式が不正です。");
        }
        int formatVersion = buf.getShort() & 0xffff;
        if (formatVersion != BINARY_FORMAT_VERSION) {
            throw new IOException(
                    "操作者名設定のバイナリ形式が未対応です (formatVersion=" + formatVersion + ")");
        }
        int payloadLen = buf.getInt();
        if (payloadLen < 0 || buf.remaining() < payloadLen) {
            throw new IOException("操作者名設定ファイルが壊れています。");
        }
        byte[] payload = new byte[payloadLen];
        buf.get(payload);
        return JSON.readTree(payload);
    }

    private static byte[] encodeBinaryDocument(ObjectNode root) throws IOException {
        byte[] payload = JSON.writeValueAsBytes(root);
        ByteBuffer buf = ByteBuffer.allocate(BINARY_MAGIC.length + 2 + 4 + payload.length);
        buf.put(BINARY_MAGIC);
        buf.putShort((short) BINARY_FORMAT_VERSION);
        buf.putInt(payload.length);
        buf.put(payload);
        return buf.array();
    }

    private static Path resolveLegacyJsonStorePath() {
        String test = System.getProperty("pm.ai.test.factoryOperatorUserLegacyStore");
        if (test != null && !test.isBlank()) {
            return Path.of(test).toAbsolutePath().normalize();
        }
        return LEGACY_JSON_STORE;
    }

    private static ObjectNode documentToObjectNode(Document doc) {
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", SCHEMA_VERSION);
        ObjectNode factories = root.putObject("factories");
        for (Map.Entry<FactorySite, FactoryOperatorUsers> e : doc.factories().entrySet()) {
            ObjectNode fo = factories.putObject(e.getKey().name());
            ArrayNode arr = fo.putArray("names");
            for (String name : e.getValue().names()) {
                arr.add(name);
            }
            fo.put("lastSelected", "");
            ObjectNode pins = fo.putObject("pinHashes");
            for (Map.Entry<String, String> pe : e.getValue().pinHashes().entrySet()) {
                if (e.getValue().names().contains(pe.getKey())) {
                    pins.put(pe.getKey(), pe.getValue());
                }
            }
            ObjectNode attempts = fo.putObject("pinFailedAttempts");
            for (Map.Entry<String, Integer> ae : e.getValue().pinFailedAttempts().entrySet()) {
                if (e.getValue().names().contains(ae.getKey()) && ae.getValue() != null && ae.getValue() > 0) {
                    attempts.put(ae.getKey(), ae.getValue());
                }
            }
            ArrayNode mustChange = fo.putArray("pinMustChange");
            for (String name : e.getValue().pinMustChange()) {
                if (e.getValue().names().contains(name)) {
                    mustChange.add(name);
                }
            }
            ObjectNode plaintextAdmin = fo.putObject("pinPlaintextAdmin");
            for (Map.Entry<String, String> pe : e.getValue().pinPlaintextAdmin().entrySet()) {
                if (e.getValue().names().contains(pe.getKey())) {
                    plaintextAdmin.put(pe.getKey(), pe.getValue());
                }
            }
            ObjectNode aladdinLoginIds = fo.putObject("aladdinLoginIds");
            for (Map.Entry<String, String> ae : e.getValue().aladdinLoginIds().entrySet()) {
                if (e.getValue().names().contains(ae.getKey())) {
                    aladdinLoginIds.put(ae.getKey(), ae.getValue());
                }
            }
            ObjectNode aladdinPasswordCiphertext = fo.putObject("aladdinPasswordCiphertext");
            for (Map.Entry<String, String> ae : e.getValue().aladdinPasswordCiphertext().entrySet()) {
                if (e.getValue().names().contains(ae.getKey())) {
                    aladdinPasswordCiphertext.put(ae.getKey(), ae.getValue());
                }
            }
        }
        return root;
    }

    private static FactoryOperatorUsers parseFactory(JsonNode node) {
        List<String> names = new ArrayList<>();
        JsonNode arr = node.get("names");
        if (arr != null && arr.isArray()) {
            LinkedHashSet<String> dedupe = new LinkedHashSet<>();
            for (JsonNode n : arr) {
                if (n == null || n.isNull()) {
                    continue;
                }
                String v = normalizeName(n.asText(""));
                if (!v.isEmpty()) {
                    dedupe.add(v);
                }
            }
            names.addAll(dedupe);
        }
        if (names.isEmpty()) {
            names.addAll(DEFAULT_NAMES);
        }
        String last = normalizeName(node.path("lastSelected").asText(""));
        if (!last.isEmpty() && !names.contains(last) && !isGuestOperator(last)) {
            last = "";
        }
        Map<String, String> pinHashes = new LinkedHashMap<>();
        JsonNode pinsNode = node.get("pinHashes");
        if (pinsNode != null && pinsNode.isObject()) {
            pinsNode
                    .fields()
                    .forEachRemaining(
                            e -> {
                                String key = normalizeName(e.getKey());
                                if (key.isEmpty() || !names.contains(key)) {
                                    return;
                                }
                                JsonNode v = e.getValue();
                                if (v != null && v.isTextual()) {
                                    String hash = v.asText("").strip();
                                    if (!hash.isEmpty()) {
                                        pinHashes.put(key, hash);
                                    }
                                }
                            });
        }
        Map<String, Integer> pinFailedAttempts = new LinkedHashMap<>();
        JsonNode attemptsNode = node.get("pinFailedAttempts");
        if (attemptsNode != null && attemptsNode.isObject()) {
            attemptsNode
                    .fields()
                    .forEachRemaining(
                            e -> {
                                String key = normalizeName(e.getKey());
                                if (key.isEmpty() || !names.contains(key)) {
                                    return;
                                }
                                JsonNode v = e.getValue();
                                if (v != null && v.isNumber()) {
                                    int count = v.asInt(0);
                                    if (count > 0) {
                                        pinFailedAttempts.put(key, count);
                                    }
                                }
                            });
        }
        Set<String> pinMustChange = new LinkedHashSet<>();
        JsonNode mustChangeNode = node.get("pinMustChange");
        if (mustChangeNode != null && mustChangeNode.isArray()) {
            for (JsonNode n : mustChangeNode) {
                if (n == null || n.isNull()) {
                    continue;
                }
                String key = normalizeName(n.asText(""));
                if (!key.isEmpty() && names.contains(key)) {
                    pinMustChange.add(key);
                }
            }
        }
        Map<String, String> pinPlaintextAdmin = new LinkedHashMap<>();
        JsonNode plaintextNode = node.get("pinPlaintextAdmin");
        if (plaintextNode != null && plaintextNode.isObject()) {
            plaintextNode
                    .fields()
                    .forEachRemaining(
                            e -> {
                                String key = normalizeName(e.getKey());
                                if (key.isEmpty() || !names.contains(key)) {
                                    return;
                                }
                                JsonNode v = e.getValue();
                                if (v != null && v.isTextual()) {
                                    String plain = normalizePin(v.asText(""));
                                    if (plain != null) {
                                        pinPlaintextAdmin.put(key, plain);
                                    }
                                }
                            });
        }
        Map<String, String> aladdinLoginIds = new LinkedHashMap<>();
        JsonNode aladdinIdsNode = node.get("aladdinLoginIds");
        if (aladdinIdsNode != null && aladdinIdsNode.isObject()) {
            aladdinIdsNode
                    .fields()
                    .forEachRemaining(
                            e -> {
                                String key = normalizeName(e.getKey());
                                if (key.isEmpty() || !names.contains(key)) {
                                    return;
                                }
                                JsonNode v = e.getValue();
                                if (v != null && v.isTextual()) {
                                    String id = v.asText("").strip();
                                    if (!id.isEmpty()) {
                                        aladdinLoginIds.put(key, id);
                                    }
                                }
                            });
        }
        Map<String, String> aladdinPasswordCiphertext = new LinkedHashMap<>();
        JsonNode aladdinPwNode = node.get("aladdinPasswordCiphertext");
        if (aladdinPwNode != null && aladdinPwNode.isObject()) {
            aladdinPwNode
                    .fields()
                    .forEachRemaining(
                            e -> {
                                String key = normalizeName(e.getKey());
                                if (key.isEmpty() || !names.contains(key)) {
                                    return;
                                }
                                JsonNode v = e.getValue();
                                if (v != null && v.isTextual()) {
                                    String cipher = v.asText("").strip();
                                    if (!cipher.isEmpty()) {
                                        aladdinPasswordCiphertext.put(key, cipher);
                                    }
                                }
                            });
        }
        return new FactoryOperatorUsers(
                names,
                last,
                pinHashes,
                pinFailedAttempts,
                pinMustChange,
                pinPlaintextAdmin,
                aladdinLoginIds,
                aladdinPasswordCiphertext);
    }

    private static void saveDocument(Document doc) throws IOException {
        Path path = storePath();
        try {
            saveDocumentToPath(path, doc);
        } catch (IOException primary) {
            Path local =
                    configuredNetworkStorePath != null
                            ? AppPaths.localFactoryOperatorUsersStorePath(
                                    factorySiteForConfiguredStore())
                            : null;
            if (local != null && !local.equals(path)) {
                configuredStorePath = local;
                usingLocalStoreFallback = true;
                saveDocumentToPath(local, doc);
                return;
            }
            throw primary;
        }
    }

    private static FactorySite factorySiteForConfiguredStore() {
        Path network = configuredNetworkStorePath;
        if (network != null) {
            Optional<FactorySite> inferred = FactorySite.inferFromPortableBundleSourceValue(network.toString());
            if (inferred.isPresent()) {
                return inferred.get();
            }
        }
        return GlobalInitSettingTarget.load();
    }

    private static Path resolveWritableStorePath(Path network, Path local) {
        if (isOperatorStorePathWritable(network)) {
            return network;
        }
        return local;
    }

    /**
     * 共有 DATA フォルダへ bin を置けるか（親の作成・書込・既存ファイルの更新）。
     */
    static boolean isOperatorStorePathWritable(Path path) {
        if (path == null) {
            return false;
        }
        try {
            Path parent = path.getParent();
            if (parent == null) {
                return false;
            }
            if (!Files.isDirectory(parent)) {
                if (!Files.exists(parent)) {
                    Files.createDirectories(parent);
                }
                if (!Files.isDirectory(parent)) {
                    return false;
                }
            }
            if (!Files.isWritable(parent)) {
                return false;
            }
            if (Files.exists(path) && !Files.isWritable(path)) {
                return false;
            }
            return true;
        } catch (IOException | SecurityException ignored) {
            return false;
        }
    }

    private static void seedLocalStoreFromNetworkIfNeeded(Path network, Path local) {
        if (!Files.isRegularFile(network) || !Files.isReadable(network) || Files.isRegularFile(local)) {
            return;
        }
        try {
            if (local.getParent() != null) {
                Files.createDirectories(local.getParent());
            }
            Files.copy(network, local);
        } catch (IOException ignored) {
        }
    }

    private static void saveDocumentToPath(Path path, Document doc) throws IOException {
        if (path.getParent() != null) {
            Files.createDirectories(path.getParent());
        }
        maybeAutoBackupBeforeSchemaUpgrade(path);
        byte[] encoded = encodeBinaryDocument(documentToObjectNode(doc));
        Files.write(
                path,
                encoded,
                StandardOpenOption.CREATE,
                StandardOpenOption.TRUNCATE_EXISTING,
                StandardOpenOption.WRITE);
        syncLauncherCredentialsJson(path, doc);
    }

    private static void maybeAutoBackupBeforeSchemaUpgrade(Path path) {
        if (!Files.isRegularFile(path)) {
            return;
        }
        String pathKey = path.toAbsolutePath().normalize().toString();
        if (SCHEMA_UPGRADE_BACKUP_PATHS.contains(pathKey)) {
            return;
        }
        try {
            JsonNode root = readStoreRoot(path);
            if (root == null || !root.isObject()) {
                return;
            }
            int ver = root.path("schemaVersion").asInt(0);
            if (ver <= 0) {
                ver = 1;
            }
            if (ver >= SCHEMA_VERSION) {
                return;
            }
            FactoryOperatorUserBackupStore.createAutomaticSchemaUpgradeBackup(
                    Map.of(), ver, "アップデート前自動バックアップ schema-" + ver);
            SCHEMA_UPGRADE_BACKUP_PATHS.add(pathKey);
        } catch (IOException ignored) {
            // バックアップ失敗でも書込は続行
        }
    }

    private static void syncLauncherCredentialsJson(Path storePath, Document doc) {
        Path parent = storePath.getParent();
        if (parent == null || doc == null) {
            return;
        }
        Path jsonPath = parent.resolve(OperatorAladdinCredentialsLauncherJson.FILE_NAME);
        try {
            Map<FactorySite, Map<String, OperatorAladdinCredentialsLauncherJson.OperatorEntry>>
                    byFactory = new LinkedHashMap<>();
            for (Map.Entry<FactorySite, FactoryOperatorUsers> e : doc.factories().entrySet()) {
                Map<String, OperatorAladdinCredentialsLauncherJson.OperatorEntry> operators =
                        new LinkedHashMap<>();
                FactoryOperatorUsers factoryUsers = e.getValue();
                for (String name : factoryUsers.names()) {
                    String loginId = factoryUsers.aladdinLoginIds().get(name);
                    String cipher = factoryUsers.aladdinPasswordCiphertext().get(name);
                    if (loginId == null
                            || loginId.isBlank()
                            || cipher == null
                            || cipher.isBlank()) {
                        continue;
                    }
                    JsonNode payload = JSON.readTree(cipher);
                    if (payload != null && payload.isObject()) {
                        operators.put(
                                name,
                                new OperatorAladdinCredentialsLauncherJson.OperatorEntry(
                                        loginId, (ObjectNode) payload));
                    }
                }
                byFactory.put(e.getKey(), operators);
            }
            OperatorAladdinCredentialsLauncherJson.writeAllFactories(jsonPath, byFactory);
        } catch (Exception ignored) {
            // 副産物 JSON の失敗は bin 保存を妨げない
        }
    }

    private static Document defaultDocument() {
        Map<FactorySite, FactoryOperatorUsers> factories = new LinkedHashMap<>();
        for (FactorySite site : FactorySite.values()) {
            factories.put(site, new FactoryOperatorUsers(DEFAULT_NAMES, ""));
        }
        return new Document(SCHEMA_VERSION, factories);
    }

    private static FactoryOperatorUsers ensureFactory(Document doc, FactorySite site) {
        FactoryOperatorUsers current = doc.factories().get(site);
        if (current != null && !current.names().isEmpty()) {
            return current;
        }
        FactoryOperatorUsers created = new FactoryOperatorUsers(DEFAULT_NAMES, "");
        doc.factories().put(site, created);
        return created;
    }

    private static String normalizeName(String raw) {
        if (raw == null) {
            return "";
        }
        String t = raw.strip();
        if (t.length() > MAX_NAME_LENGTH) {
            t = t.substring(0, MAX_NAME_LENGTH);
        }
        return t;
    }

    public static String normalizePin(String raw) {
        if (raw == null) {
            return null;
        }
        String t = raw.strip();
        if (t.length() < MIN_PIN_LENGTH || t.length() > MAX_PIN_LENGTH || !t.chars().allMatch(Character::isDigit)) {
            return null;
        }
        return t;
    }

    /** 永続化ファイルが無いとき、現行内容をディスクへ書き出す。 */
    public static void ensureStoreFileOnDisk() throws IOException {
        Path path = storePath();
        if (!Files.isRegularFile(path)) {
            saveDocument(loadDocument());
        }
    }

    private static String generatePin() {
        int n = SECURE_RANDOM.nextInt(10_000);
        return String.format("%04d", n);
    }

    private static void clearPinFailures(Document doc, FactorySite factory, String normalized) {
        FactoryOperatorUsers current = doc.factories().get(factory);
        if (current == null || !current.pinFailedAttempts().containsKey(normalized)) {
            return;
        }
        Map<String, Integer> attempts = new LinkedHashMap<>(current.pinFailedAttempts());
        attempts.remove(normalized);
        doc.factories()
                .put(
                        factory,
                        forSharedStore(
                                current,
                                current.names(),
                                current.pinHashes(),
                                attempts,
                                current.pinMustChange(),
                                current.pinPlaintextAdmin()));
    }

    private static FactoryOperatorUsers forSharedStore(
            FactoryOperatorUsers base,
            List<String> names,
            Map<String, String> pinHashes,
            Map<String, Integer> pinFailedAttempts,
            Set<String> pinMustChange,
            Map<String, String> pinPlaintextAdmin) {
        return forSharedStore(
                names,
                pinHashes,
                pinFailedAttempts,
                pinMustChange,
                pinPlaintextAdmin,
                base != null ? base.aladdinLoginIds() : Map.of(),
                base != null ? base.aladdinPasswordCiphertext() : Map.of());
    }

    private static FactoryOperatorUsers forSharedStore(
            List<String> names,
            Map<String, String> pinHashes,
            Map<String, Integer> pinFailedAttempts,
            Set<String> pinMustChange,
            Map<String, String> pinPlaintextAdmin,
            Map<String, String> aladdinLoginIds,
            Map<String, String> aladdinPasswordCiphertext) {
        return new FactoryOperatorUsers(
                names,
                "",
                pinHashes,
                pinFailedAttempts,
                pinMustChange,
                pinPlaintextAdmin,
                filterAladdinMap(aladdinLoginIds, names),
                filterAladdinMap(aladdinPasswordCiphertext, names));
    }

    private static Map<String, String> filterAladdinMap(
            Map<String, String> source, List<String> names) {
        Map<String, String> out = new LinkedHashMap<>();
        if (source == null || names == null) {
            return out;
        }
        for (String name : names) {
            String value = source.get(name);
            if (value != null && !value.isBlank()) {
                out.put(name, value);
            }
        }
        return out;
    }

    private static Path localLastSelectedPath(FactorySite site) {
        String test = System.getProperty("pm.ai.test.factoryOperatorLastSelectedDir");
        if (test != null && !test.isBlank()) {
            FactorySite effective = site != null ? site : FactorySite.KONAN;
            String suffix = effective.name().toLowerCase(Locale.ROOT);
            return Path.of(test)
                    .resolve("last-factory-operator-" + suffix + ".txt")
                    .toAbsolutePath()
                    .normalize();
        }
        return AppPaths.localFactoryOperatorLastSelectedPath(site);
    }

    private static String loadLastSelectedLocal(FactorySite site) throws IOException {
        Path path = localLastSelectedPath(site);
        if (!Files.isRegularFile(path)) {
            return "";
        }
        String text = Files.readString(path, StandardCharsets.UTF_8).strip();
        return normalizeName(text);
    }

    private static void saveLastSelectedLocal(FactorySite site, String name) throws IOException {
        String normalized = normalizeName(name);
        if (normalized.isEmpty()) {
            clearLastSelectedLocal(site);
            return;
        }
        Path path = localLastSelectedPath(site);
        if (path.getParent() != null) {
            Files.createDirectories(path.getParent());
        }
        Files.writeString(
                path,
                normalized + System.lineSeparator(),
                StandardCharsets.UTF_8,
                StandardOpenOption.CREATE,
                StandardOpenOption.TRUNCATE_EXISTING);
    }

    private static void clearLastSelectedLocal(FactorySite site) throws IOException {
        Files.deleteIfExists(localLastSelectedPath(site));
    }

    /**
     * 旧版で共有 bin に残っていた {@code lastSelected} を、初回のみローカルへ移す。
     */
    private static void migrateLastSelectedFromSharedStoreIfNeeded(FactorySite factory) throws IOException {
        if (Files.isRegularFile(localLastSelectedPath(factory))) {
            return;
        }
        FactoryOperatorUsers shared = ensureFactory(loadDocument(), factory);
        String legacy = shared.lastSelected();
        if (legacy.isEmpty()) {
            return;
        }
        saveLastSelectedLocal(factory, legacy);
        if (!legacy.isEmpty()) {
            Document doc = loadDocument();
            FactoryOperatorUsers current = ensureFactory(doc, factory);
            doc.factories()
                    .put(
                            factory,
                            forSharedStore(
                                    current,
                                    current.names(),
                                    current.pinHashes(),
                                    current.pinFailedAttempts(),
                                    current.pinMustChange(),
                                    current.pinPlaintextAdmin()));
            saveDocument(doc);
        }
    }

    private static String hashPin(FactorySite factory, String name, String pin) {
        String payload =
                factory.name() + "|" + normalizeName(name) + "|" + Objects.requireNonNull(pin);
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] digest = md.digest(payload.getBytes(StandardCharsets.UTF_8));
            return HEX.formatHex(digest);
        } catch (NoSuchAlgorithmException ex) {
            throw new IllegalStateException(ex);
        }
    }

    private record Document(int schemaVersion, Map<FactorySite, FactoryOperatorUsers> factories) {

        Document {
            factories = factories != null ? new LinkedHashMap<>(factories) : new LinkedHashMap<>();
        }
    }

    /** テスト用: ストアを既定状態へ戻す。 */
    public static void resetStoreForTests() throws IOException {
        sessionOperatorName = "";
        configuredStorePath = null;
        configuredNetworkStorePath = null;
        usingLocalStoreFallback = false;
        storeConfigured = false;
        SCHEMA_UPGRADE_BACKUP_PATHS.clear();
        Path path = storePath();
        Files.deleteIfExists(path);
        for (FactorySite site : FactorySite.values()) {
            Files.deleteIfExists(localLastSelectedPath(site));
        }
    }

    /** テスト用: ファイルを直接書き込む。 */
    public static void writeRawJsonForTests(String json) throws IOException {
        Path path = storePath();
        if (path.getParent() != null) {
            Files.createDirectories(path.getParent());
        }
        Files.writeString(path, Objects.requireNonNull(json), StandardCharsets.UTF_8);
    }
}
