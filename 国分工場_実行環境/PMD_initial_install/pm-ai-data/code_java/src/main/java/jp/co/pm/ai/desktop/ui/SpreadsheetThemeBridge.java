package jp.co.pm.ai.desktop.ui;

import java.util.Objects;

import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * Applies theme-aware CSS after ControlsFX {@code spreadsheet.css} (which hard-codes white cell fills).
 */
public final class SpreadsheetThemeBridge {

    private static final String BRIDGE_CSS = "/jp/co/pm/ai/desktop/css/spreadsheet-theme-bridge.css";

    private SpreadsheetThemeBridge() {}

    /** Adds the bridge stylesheet once so {@link SpreadsheetView} cells follow the scene theme. */
    public static void install(SpreadsheetView view) {
        String url =
                Objects.requireNonNull(SpreadsheetThemeBridge.class.getResource(BRIDGE_CSS), BRIDGE_CSS)
                        .toExternalForm();
        if (!view.getStylesheets().contains(url)) {
            view.getStylesheets().add(url);
        }
    }
}
