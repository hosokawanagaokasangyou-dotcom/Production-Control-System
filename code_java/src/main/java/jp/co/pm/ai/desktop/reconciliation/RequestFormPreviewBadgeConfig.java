package jp.co.pm.ai.desktop.reconciliation;

import jp.co.pm.ai.desktop.config.PersonBadgeStyle;

/** 依頼書プレビュー上部の「原本更新」バッジ表示設定。 */
public record RequestFormPreviewBadgeConfig(String label, PersonBadgeStyle style) {

    public RequestFormPreviewBadgeConfig {
        label = label != null && !label.isBlank() ? label.strip() : "更新";
        style = style != null ? style : PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault();
    }

    public static RequestFormPreviewBadgeConfig defaults() {
        return new RequestFormPreviewBadgeConfig(
                "更新", PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault());
    }
}
