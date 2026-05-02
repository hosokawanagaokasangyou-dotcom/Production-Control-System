package jp.co.pm.ai.desktop.ui;

import javafx.stage.FileChooser;

import jp.co.pm.ai.desktop.config.AppPaths;

/** Maps env var names to {@link FileChooser.ExtensionFilter} sets for the desktop UI. */
public final class FileChooserForEnvKey {

    private FileChooserForEnvKey() {}

    public static void apply(FileChooser fc, String envKey) {
        fc.getExtensionFilters().clear();
        String k = envKey != null ? envKey.trim() : "";
        if (AppPaths.isJsonFilePathEnvKey(k)) {
            fc.getExtensionFilters()
                    .addAll(
                            new FileChooser.ExtensionFilter("JSON", "*.json"),
                            new FileChooser.ExtensionFilter("All", "*.*"));
        } else if (AppPaths.isCsvFilePathEnvKey(k)) {
            fc.getExtensionFilters()
                    .addAll(
                            new FileChooser.ExtensionFilter("CSV", "*.csv"),
                            new FileChooser.ExtensionFilter("All", "*.*"));
        } else if (AppPaths.isExcelWorkbookPathEnvKey(k)) {
            fc.getExtensionFilters()
                    .addAll(
                            new FileChooser.ExtensionFilter("Excel", "*.xlsm", "*.xlsx"),
                            new FileChooser.ExtensionFilter("All", "*.*"));
        } else {
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("All", "*.*"));
        }
    }
}
