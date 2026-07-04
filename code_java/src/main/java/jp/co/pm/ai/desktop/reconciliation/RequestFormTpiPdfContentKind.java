package jp.co.pm.ai.desktop.reconciliation;

/** TPI 依頼書 PDF の読取方式（テキスト抽出か OCR か）。 */
enum RequestFormTpiPdfContentKind {
    /** PDFBox テキスト抽出で依頼書項目を読める。 */
    TEXT,
    /** 画像スキャン PDF のため OCR が必要。 */
    IMAGE_SCAN
}
