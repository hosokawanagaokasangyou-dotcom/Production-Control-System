package jp.co.pm.ai.kouchin.verify;

/** 対象工場。②のフォーマット・フォルダ・①の入庫場所・③の得意先絞り込みが切り替わる。 */
public enum FactoryId {

    /** 国分工場（②=長岡明細 東レまとめ / 入庫場所 A010） */
    KOKUBU("kokubu"),
    /** 湖南工場（②=加工賃試算 東レ3シート / 入庫場所 A010P） */
    KONAN("konan");

    private final String code;

    FactoryId(String code) {
        this.code = code;
    }

    public String code() {
        return code;
    }

    public static FactoryId fromCode(String code) {
        for (FactoryId id : values()) {
            if (id.code.equalsIgnoreCase(code)) {
                return id;
            }
        }
        throw new IllegalArgumentException("不明な工場コードです: " + code);
    }
}
