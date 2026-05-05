package jp.co.pm.ai.desktop.config;

/**
 * アプリ内の「用途別バッジ」編集対象。列挙を増やすと UI バッジタブの選択肢に現れる。
 */
public enum UiBadgePurpose {
    /** 実行・ログタブの段階1付近に表示（ネットワークソースをキャッシュから読んだとき）。 */
    STAGE1_NETWORK_CACHE("stage1NetworkCache", "段階1・ソースキャッシュ表示"),

    /** 設備ガントの担当者バッジ（詳細編集は専用タブ）。 */
    EQUIPMENT_GANTT_PERSON("equipmentGanttPerson", "設備ガント・担当バッジ（別タブで編集）"),

    /** 将来拡張用（スタイル枠のみ）。 */
    RESERVED("reserved", "（予約・将来のバッジ用）");

    private final String storageKey;
    private final String displayLabel;

    UiBadgePurpose(String storageKey, String displayLabel) {
        this.storageKey = storageKey;
        this.displayLabel = displayLabel;
    }

    /** セッション JSON 等に書く安定キー。 */
    public String storageKey() {
        return storageKey;
    }

    public String displayLabel() {
        return displayLabel;
    }

    public static UiBadgePurpose fromStorageKey(String k) {
        if (k == null || k.isBlank()) {
            return STAGE1_NETWORK_CACHE;
        }
        String t = k.trim();
        for (UiBadgePurpose p : values()) {
            if (p.storageKey.equals(t)) {
                return p;
            }
        }
        return STAGE1_NETWORK_CACHE;
    }
}
