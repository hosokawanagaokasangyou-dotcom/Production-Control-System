package jp.co.pm.ai.kouchin.trend;

import java.util.Map;

import jp.co.pm.ai.kouchin.verify.KouchinPaths;

/**
 * 月次トレンド実行パラメータ。
 *
 * @param months 直近月数（0以下なら 6）
 * @param factory {@code kokubu} / {@code konan} / {@code both}
 * @param paths 検証と同じパス束。湖南②は {@code konanShisanDir}
 * @param ui 環境タブ値。国分年度フォルダと第3コピー先を読む
 */
public record TrendRequest(int months, String factory, KouchinPaths paths, Map<String, String> ui) {

    public static final int DEFAULT_MONTHS = 6;
    public static final String FACTORY_KOKUBU = "kokubu";
    public static final String FACTORY_KONAN = "konan";
    public static final String FACTORY_BOTH = "both";

    public TrendRequest {
        if (months <= 0) {
            months = DEFAULT_MONTHS;
        }
        factory = normalizeFactory(factory);
        if (paths == null) {
            paths = ui == null || ui.isEmpty() ? KouchinPaths.defaults() : KouchinPaths.fromEnv(ui);
        }
        ui = ui == null ? Map.of() : Map.copyOf(ui);
    }

    public boolean wantKokubu() {
        return FACTORY_BOTH.equals(factory) || FACTORY_KOKUBU.equals(factory);
    }

    public boolean wantKonan() {
        return FACTORY_BOTH.equals(factory) || FACTORY_KONAN.equals(factory);
    }

    static String normalizeFactory(String factory) {
        if (factory == null || factory.isBlank()) {
            return FACTORY_BOTH;
        }
        String f = factory.trim();
        if (FACTORY_KOKUBU.equalsIgnoreCase(f) || FACTORY_KONAN.equalsIgnoreCase(f) || FACTORY_BOTH.equalsIgnoreCase(f)) {
            return f.toLowerCase();
        }
        throw new IllegalArgumentException("工場は kokubu / konan / both で指定してください: " + factory);
    }
}
