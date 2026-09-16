package jp.co.pm.ai.kouchin.verify;

/**
 * 検証の実行入口。
 * 1工場分の実行（{@link #run}）と、両工場をまとめて実行しメール下書きまで作る
 * {@link #runBoth} を提供する。
 */
public final class VerifyService {

    /** 一致とみなす許容差（円）。1円差も不一致として検出する。 */
    public static final double DEFAULT_TOLERANCE = 0.5;

    private VerifyService() {
    }

    /** 既定の許容差で1工場分を実行する。 */
    public static VerifyResult run(FactoryId factory, KouchinPaths paths) {
        return run(factory, paths, DEFAULT_TOLERANCE);
    }

    /**
     * 1工場分の検証を実行し、メール用の数字を直近結果としてキャッシュに保存する。
     *
     * @param factory 対象工場
     * @param paths   各データソースのフォルダ
     * @param tol     一致とみなす許容差（円）
     */
    public static VerifyResult run(FactoryId factory, KouchinPaths paths, double tol) {
        FileDiscovery.invalidateListingCache();
        VerifyResult result = VerifyEngine.run(FactoryProfile.of(factory), paths, tol);
        LastResultCache.saveQuietly(factory, result.mail());
        return result;
    }

    /** 既定の許容差で両工場を実行する。 */
    public static BothResult runBoth(KouchinPaths paths) {
        return runBoth(paths, DEFAULT_TOLERANCE);
    }

    /**
     * 両工場を実行する。片方が失敗しても続行し、失敗した工場のメール行は
     * {@link UnifiedMailBuilder#NOT_VERIFIED} になる。
     */
    public static BothResult runBoth(KouchinPaths paths, double tol) {
        VerifyResult kokubu = null;
        VerifyResult konan = null;
        String kokubuError = null;
        String konanError = null;

        try {
            kokubu = run(FactoryId.KOKUBU, paths, tol);
        } catch (RuntimeException e) {
            kokubuError = e.getMessage() == null ? e.toString() : e.getMessage();
        }
        try {
            konan = run(FactoryId.KONAN, paths, tol);
        } catch (RuntimeException e) {
            konanError = e.getMessage() == null ? e.toString() : e.getMessage();
        }

        String mail = UnifiedMailBuilder.buildText(
                kokubu == null ? null : kokubu.mail(),
                konan == null ? null : konan.mail());
        return new BothResult(kokubu, konan, kokubuError, konanError, mail);
    }
}
