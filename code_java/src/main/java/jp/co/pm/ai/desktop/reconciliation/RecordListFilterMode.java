package jp.co.pm.ai.desktop.reconciliation;

/** 依頼一覧コンボの表示範囲（起動時は {@link #ALL}。ラジオ再クリックで未選択のときは {@link #WITH_ORIGINAL}）。 */
enum RecordListFilterMode {
    /** 読込済みの全行（起動時ラジオ既定）。 */
    ALL,
    /** 依頼書原本ファイルがある行。 */
    WITH_ORIGINAL,
    /** ステータスに「既存」を含む行。 */
    EXISTING_ONLY,
    /** 依頼書ありかつステータスに「新規」を含む行。 */
    NEW_ONLY,
    /** 受注ファイルのみ（原本なし）、入力日降順。 */
    JUCHU_WITHOUT_ORIGINAL
}
