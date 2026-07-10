package jp.co.pm.ai.desktop.io;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * リモートデスクトップタブの起動プロファイル・クイック起動ボタン向けロジック。
 *
 * <p>カタログ順（{@code profileFields} の並び＝プロファイル番号昇順）の先頭 N 件をボタンに割り当てる。
 */
public final class RdpLaunchProfileQuickLaunch {

    /** クイック起動ボタンの最大数。 */
    public static final int BUTTON_SLOT_COUNT = 8;

    /** ボタン表示ラベルの最大文字数（超過時は末尾を「…」で省略）。 */
    public static final int BUTTON_LABEL_MAX_LENGTH = 36;

    private RdpLaunchProfileQuickLaunch() {}

    /**
     * カタログ順のプロファイル番号リストから、クイック起動用の先頭 {@code limit} 件を返す。
     */
    public static List<Integer> catalogOrderProfileNumbers(List<Integer> catalogOrder, int limit) {
        Objects.requireNonNull(catalogOrder, "catalogOrder");
        if (limit <= 0 || catalogOrder.isEmpty()) {
            return List.of();
        }
        int take = Math.min(limit, catalogOrder.size());
        return List.copyOf(catalogOrder.subList(0, take));
    }

    /** クイック起動ボタン向けに {@link RdpLaunchProfile#displayLabel()} 等の全文を短縮する。 */
    public static String buttonLabel(String fullLabel) {
        return buttonLabel(fullLabel, BUTTON_LABEL_MAX_LENGTH);
    }

    public static String buttonLabel(String fullLabel, int maxLength) {
        if (fullLabel == null) {
            return "";
        }
        String trimmed = fullLabel.trim();
        if (maxLength <= 0 || trimmed.length() <= maxLength) {
            return trimmed;
        }
        if (maxLength <= 1) {
            return "…";
        }
        return trimmed.substring(0, maxLength - 1) + "…";
    }

    /** {@code catalogOrder} の先頭 {@link #BUTTON_SLOT_COUNT} 件（テスト・UI 共通）。 */
    public static List<Integer> quickLaunchProfileNumbers(List<Integer> catalogOrder) {
        return catalogOrderProfileNumbers(catalogOrder, BUTTON_SLOT_COUNT);
    }

    /** スロット index（0 始まり）に対応するプロファイル番号。範囲外なら empty。 */
    public static List<Integer> slotProfileNumbers(List<Integer> catalogOrder) {
        List<Integer> quick = quickLaunchProfileNumbers(catalogOrder);
        List<Integer> slots = new ArrayList<>(BUTTON_SLOT_COUNT);
        for (int i = 0; i < BUTTON_SLOT_COUNT; i++) {
            slots.add(i < quick.size() ? quick.get(i) : null);
        }
        return slots;
    }
}
