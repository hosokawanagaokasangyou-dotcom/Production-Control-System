package jp.co.pm.ai.desktop.ui;

import java.util.List;

/** 段階1 EC面区分「不明」選択ダイアログの確定結果。 */
public record Stage1EcSideUnknownDialogResult(List<Selection> selections) {

    public record Selection(String iraiNo, String ecSideClass) {}
}
