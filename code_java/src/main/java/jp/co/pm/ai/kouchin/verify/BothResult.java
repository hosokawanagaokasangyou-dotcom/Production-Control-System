package jp.co.pm.ai.kouchin.verify;

import java.util.Optional;

/**
 * 両工場の検証結果と、2工場をまとめたメール下書き。
 * 片方が失敗した場合、その工場は {@code null} でエラーメッセージが入る。
 *
 * @param kokubu       国分工場の結果（失敗時 null）
 * @param konan        湖南工場の結果（失敗時 null）
 * @param kokubuError  国分工場の失敗理由（成功時 null）
 * @param konanError   湖南工場の失敗理由（成功時 null）
 * @param unifiedMail  2工場分をまとめたメール下書き
 */
public record BothResult(
        VerifyResult kokubu,
        VerifyResult konan,
        String kokubuError,
        String konanError,
        String unifiedMail) {

    public Optional<VerifyResult> of(FactoryId factory) {
        return Optional.ofNullable(factory == FactoryId.KOKUBU ? kokubu : konan);
    }

    public Optional<String> errorOf(FactoryId factory) {
        return Optional.ofNullable(factory == FactoryId.KOKUBU ? kokubuError : konanError);
    }

    /** 両工場とも成功したか。 */
    public boolean isComplete() {
        return kokubu != null && konan != null;
    }
}
