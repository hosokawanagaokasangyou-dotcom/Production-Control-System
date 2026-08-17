package jp.co.pm.ai.desktop.ui;

import javafx.scene.control.Alert;
import javafx.scene.control.TextArea;

/**
 * 長い本文の {@link Alert} が画面高を超えて OK が押せなくなるのを防ぐ。
 */
public final class ThemedAlertContentSupport {

    static final int SCROLL_LINE_THRESHOLD = 8;

    static final int SCROLL_CHAR_THRESHOLD = 400;

    private static final double SCROLL_PREF_WIDTH = 760;

    private static final double SCROLL_PREF_HEIGHT = 320;

    private static final double SCROLL_MAX_HEIGHT = 420;

    private ThemedAlertContentSupport() {}

    public static boolean needsScrollableContent(String message) {
        if (message == null || message.isEmpty()) {
            return false;
        }
        int lines = 1;
        for (int i = 0; i < message.length(); i++) {
            if (message.charAt(i) == '\n') {
                lines++;
            }
        }
        return lines >= SCROLL_LINE_THRESHOLD || message.length() >= SCROLL_CHAR_THRESHOLD;
    }

    /** 長文ならスクロール可能な TextArea、短文なら通常の contentText。 */
    public static void applyContent(Alert alert, String message) {
        if (alert == null) {
            return;
        }
        String text = message != null ? message : "";
        if (!needsScrollableContent(text)) {
            alert.setContentText(text);
            return;
        }
        alert.setContentText(null);
        TextArea area = new TextArea(text);
        area.setEditable(false);
        area.setWrapText(true);
        area.setPrefWidth(SCROLL_PREF_WIDTH);
        area.setPrefHeight(SCROLL_PREF_HEIGHT);
        area.setMaxHeight(SCROLL_MAX_HEIGHT);
        area.setMinHeight(160);
        alert.getDialogPane().setContent(area);
        alert.getDialogPane().setPrefWidth(SCROLL_PREF_WIDTH + 40);
        alert.setResizable(true);
    }
}
