package jp.co.pm.ai.desktop.config;

/** One row of the 環境変数 tab persisted in {@link DesktopSessionStateStore}. */
public record UiEnvRowSnapshot(String name, String value, String description) {}
