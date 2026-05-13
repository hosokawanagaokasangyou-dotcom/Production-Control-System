package jp.co.pm.ai.desktop.runtime;

import java.util.Locale;

/** 起動時に確定した {@code prism.order} を参照するヘルパ（JavaFX ツールキット初期化後も参照可）。 */
public final class PrismPipelineHints {

    private PrismPipelineHints() {}

    /**
     * 先頭が {@code es2} / {@code d3d} / {@code metal} のとき、ハードウェアラスタライザが優先される。
     *
     * <p>この構成では Effect（例: 担当バッジの DropShadow）が Prism のマスク経路で NPE になり得るため、
     * 見た目を簡素化する分岐に使う。
     */
    public static boolean hardwareRasterizerFirst() {
        String o = System.getProperty("prism.order", "").trim().toLowerCase(Locale.ROOT);
        if (o.isEmpty()) {
            return false;
        }
        return o.startsWith("es2") || o.startsWith("d3d") || o.startsWith("metal");
    }
}
