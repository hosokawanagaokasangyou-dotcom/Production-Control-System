package jp.co.pm.ai.desktop.ui;

import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.Tooltip;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.runtime.FxJvmMemoryStatusBar;

/** メインシェル下部のグローバルステータスバー表示。 */
public final class GlobalAppStatusBar {

    private static final int MESSAGE_MAX = 240;

    private final Label messageLabel;
    private final ProgressIndicator progressIndicator;
    private final ProgressBar progressBar;
    private final Label tabLabel;
    private final Label operatorLabel;
    private final Label factoryLabel;
    private final Label attendanceLabel;
    private final Label memoryLabel;

    public GlobalAppStatusBar(
            Label messageLabel,
            ProgressIndicator progressIndicator,
            ProgressBar progressBar,
            Label tabLabel,
            Label operatorLabel,
            Label factoryLabel,
            Label attendanceLabel,
            Label memoryLabel) {
        this.messageLabel = messageLabel;
        this.progressIndicator = progressIndicator;
        this.progressBar = progressBar;
        this.tabLabel = tabLabel;
        this.operatorLabel = operatorLabel;
        this.factoryLabel = factoryLabel;
        this.attendanceLabel = attendanceLabel;
        this.memoryLabel = memoryLabel;
    }

    public void startMemoryMonitor(Stage stage) {
        FxJvmMemoryStatusBar.start(memoryLabel, stage);
    }

    public void setMessage(String message) {
        if (messageLabel == null) {
            return;
        }
        String text = shorten(message, MESSAGE_MAX);
        messageLabel.setText(text);
        if (message != null && message.length() > MESSAGE_MAX) {
            messageLabel.setTooltip(new Tooltip(message));
        } else {
            messageLabel.setTooltip(null);
        }
    }

    /**
     * 長時間タスクの進捗表示。
     *
     * @param fraction 0.0–1.0。{@code null} で非表示。{@link Double#isNaN} で不定（スピナー）。
     */
    public void setTaskProgress(Double fraction) {
        boolean showIndeterminate = fraction != null && fraction.isNaN();
        boolean showBar = fraction != null && !showIndeterminate;
        if (progressIndicator != null) {
            progressIndicator.setVisible(showIndeterminate);
            progressIndicator.setManaged(showIndeterminate);
        }
        if (progressBar != null) {
            progressBar.setVisible(showBar);
            progressBar.setManaged(showBar);
            if (showBar) {
                double clamped = Math.max(0.0, Math.min(1.0, fraction));
                progressBar.setProgress(clamped);
            }
        }
    }

    public void clearTaskProgress() {
        setTaskProgress(null);
    }

    public void setTabName(String tabName) {
        setMeta(tabLabel, "タブ", tabName);
    }

    public void setOperator(String operator) {
        setMeta(operatorLabel, "操作者", operator);
    }

    public void setFactory(String factoryLabelText) {
        setMeta(factoryLabel, "工場", factoryLabelText);
    }

    public void setAttendanceReady(boolean ready, String detail) {
        if (attendanceLabel == null) {
            return;
        }
        String shortText = ready ? "勤怠: OK" : "勤怠: 未準備";
        attendanceLabel.setText(shortText);
        attendanceLabel.getStyleClass().remove("pm-global-status-warn");
        if (!ready) {
            attendanceLabel.getStyleClass().add("pm-global-status-warn");
        }
        if (detail != null && !detail.isBlank()) {
            attendanceLabel.setTooltip(new Tooltip(detail));
        } else {
            attendanceLabel.setTooltip(null);
        }
    }

    private static void setMeta(Label label, String prefix, String value) {
        if (label == null) {
            return;
        }
        String text = value != null && !value.isBlank() ? value : "—";
        label.setText(prefix + ": " + text);
    }

    private static String shorten(String message, int max) {
        if (message == null) {
            return "";
        }
        if (message.length() <= max) {
            return message;
        }
        return message.substring(0, max - 1) + "…";
    }
}
