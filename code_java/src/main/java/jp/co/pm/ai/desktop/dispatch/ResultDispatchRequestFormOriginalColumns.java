package jp.co.pm.ai.desktop.dispatch;

import java.text.Normalizer;
import java.util.Locale;
import java.util.Set;

/**
 * 結果_配台表（納期管理ビュー配台結果タブ）で、依頼書原本→受注転記経由の静的属性列見出し。
 *
 * <p>加工計画DATA（アラジン）主体の数量・完了系や、配台試行・日時・メンバー等の算出列は含めない。
 */
public final class ResultDispatchRequestFormOriginalColumns {

    public static final String HEADER_STYLE_CLASS = "pm-request-form-original-column-header";

    private static final Set<String> NORMALIZED_TITLES =
            Set.of(
                    norm("受注日"),
                    norm("受注NO"),
                    norm("依頼NO"),
                    norm("品名(原反)"),
                    norm("使用原反"),
                    norm("原反数"),
                    norm("品名(製品)"),
                    norm("製品名"),
                    norm("加工内容"),
                    norm("在庫場所"),
                    norm("原反投入日"),
                    norm("指定納期"),
                    norm("回答納期"),
                    norm("原反投入場所"));

    private ResultDispatchRequestFormOriginalColumns() {}

    public static boolean isDerivedFromRequestFormOriginal(String columnTitle) {
        if (columnTitle == null || columnTitle.isBlank()) {
            return false;
        }
        return NORMALIZED_TITLES.contains(norm(columnTitle));
    }

    static String norm(String title) {
        return Normalizer.normalize(title.strip(), Normalizer.Form.NFKC)
                .replace(" ", "")
                .replace("　", "")
                .toUpperCase(Locale.ROOT);
    }
}
