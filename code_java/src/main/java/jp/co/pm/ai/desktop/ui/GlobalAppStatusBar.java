package jp.co.pm.ai.desktop.ui;

import javafx.scene.control.Label;
import javafx.scene.control.Tooltip;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.runtime.FxJvmMemoryStatusBar;

/** メインシェル下部のグローバルステータスバー表示。 */
public final class GlobalAppStatusBar {

    private static final int MESSAGE_MAX = 240;

    private final Label messageLabel;
    private final Label tabLabel;
    private final Label operatorLabel;
    private final Label factoryLabel;
    private final Label attendanceLabel;
    private final Label memoryLabel;

    public GlobalAppStatusBar(
            Label messageLabel,
            Label tabLabel,
            Label operatorLabel,
            Label factoryLabel,
            Label attendanceLabel,
            Label memoryLabel) {
        this.messageLabel = messageLabel;
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
        String v = value != null && !value.isBlank() ? value.strip() : "—";
        label.setText(prefix + ": " + v);
    }

    private static String shorten(String raw, int max) {
        if (raw == null || raw.isBlank()) {
            return "";
        }
        String oneLine = raw.replace('\r', ' ').replace('\n', ' ').strip();
        if (oneLine.length() <= max) {
            return oneLine;
        }
        return oneLine.substring(0, max - 1) + "…";
    }
}
