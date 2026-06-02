package jp.co.pm.ai.desktop.reconciliation;

/** 依頼一覧コンボの表示範囲（ラジオ未選択時は {@link #WITH_ORIGINAL}）。 */
enum RecordListFilterMode {
    /** 依頼書原本ファイルがある行（既定）。 */
    WITH_ORIGINAL,
    /** 読込済みの全行。 */
    ALL,
    /** ステータスに「既存」を含む行。 */
    EXISTING_ONLY,
    /** 依頼書ありかつステータスに「新規」を含む行。 */
    NEW_ONLY,
    /** 受注ファイルのみ（原本なし）、入力日降順。 */
    JUCHU_WITHOUT_ORIGINAL
}
