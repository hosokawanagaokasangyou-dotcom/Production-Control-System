package jp.co.pm.ai.desktop.config;

import java.util.Map;

import javafx.scene.Node;
import javafx.scene.control.Label;

import jp.co.pm.ai.desktop.MainShellTabId;
import jp.co.pm.ai.desktop.PipelineExecutionTimingKind;

/** 段階3ロジックを残したまま、関連UIの表示可否だけを一元判定する。 */
public final class Stage3UiVisibility {

    private Stage3UiVisibility() {}

    /** 未設定時は非表示。環境変数タブの truthy 値でのみ表示する。 */
    public static boolean isVisible(Map<String, String> ui) {
        return AppPaths.isTruthyUiEnv(
                ui, AppPaths.KEY_PM_AI_STAGE3_UI_VISIBLE, false);
    }

    /** メインシェルの段階3入力タブだけを設定に従って隠す。 */
    public static boolean isMainShellTabVisible(MainShellTabId id, Map<String, String> ui) {
        return id != MainShellTabId.PLAN_INPUT_STAGE3 || isVisible(ui);
    }

    /** 段階3系の計測表示だけを設定に従って隠す。履歴データは変更しない。 */
    public static boolean isTimingKindVisible(
            PipelineExecutionTimingKind kind, Map<String, String> ui) {
        if (kind == null) {
            return true;
        }
        return switch (kind) {
            case STAGE3_0, STAGE3_1, STAGE3_2, STAGE3 -> isVisible(ui);
            default -> true;
        };
    }

    /** JavaFXノードの表示とレイアウト参加を必ず同時に切り替える。 */
    public static void apply(Node node, boolean visible) {
        if (node == null) {
            return;
        }
        node.setVisible(visible);
        node.setManaged(visible);
    }

    /** 段階3バッジだけをOFF時に隠す。段階2/2.1バッジは既存表示を維持する。 */
    public static void applyPlanningStageBadgePolicy(Label badge, Map<String, String> ui) {
        if (badge == null || isVisible(ui)) {
            return;
        }
        String text = badge.getText();
        if (text != null && text.startsWith("段階3")) {
            apply(badge, false);
        }
    }
}
