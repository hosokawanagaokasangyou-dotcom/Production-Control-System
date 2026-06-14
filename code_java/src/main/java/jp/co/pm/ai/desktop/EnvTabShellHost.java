package jp.co.pm.ai.desktop;

import javafx.collections.ObservableList;
import javafx.stage.Stage;

/** {@link EnvTabController} が依存するシェル操作。 */
public interface EnvTabShellHost extends DesktopShellHost {

    Stage getPrimaryStage();

    ObservableList<EnvVarRow> getEnvRows();

    void confirmAndResetEnvRowsToDefaults();

    void addMissingReferenceEnvRows();

    default void requestGeminiFreeTierModelsForceRefresh() {}

    default void refreshApiModelBenchmarkDerivedLabels() {}
}
