package jp.co.pm.ai.desktop.reconciliation;

import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Tooltip;

/**
 * 依頼書入力まわりの「予約ボタン」（表示名だけ先行配置し、処理は後続実装）。
 */
public enum RequestFormReservedButton {

    /** アップロード用 Excel の行を本番アラジンマスタへ反映（未実装）。 */
    ALADDIN_MASTER_UPSERT(
            "アラジンマスタに追加/更新",
            "アップロード用 Excel の行を本番の後加工商品マスタ.xlsx へ反映します（準備中）。");

    private final String reservedName;
    private final String tooltipText;

    RequestFormReservedButton(String reservedName, String tooltipText) {
        this.reservedName = reservedName;
        this.tooltipText = tooltipText;
    }

    /** 予約名（ボタン表示テキスト）。 */
    public String reservedName() {
        return reservedName;
    }

    public String tooltipText() {
        return tooltipText;
    }

    public Button toButton(Label statusLabel) {
        Button btn = new Button(reservedName);
        btn.getStyleClass().add("btn-request-form-reserved");
        btn.setTooltip(new Tooltip(tooltipText));
        btn.setOnAction(
                e -> {
                    if (statusLabel != null) {
                        statusLabel.setText("予約ボタン「" + reservedName + "」は準備中です。");
                    }
                });
        return btn;
    }
}
