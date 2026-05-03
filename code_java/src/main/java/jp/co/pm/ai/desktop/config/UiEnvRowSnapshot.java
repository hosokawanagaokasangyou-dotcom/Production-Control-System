package jp.co.pm.ai.desktop.config;

/** One row of the \u74b0\u5883\u5909\u6570 tab persisted in {@link DesktopSessionStateStore}. */
public record UiEnvRowSnapshot(String name, String value, String description) {}
