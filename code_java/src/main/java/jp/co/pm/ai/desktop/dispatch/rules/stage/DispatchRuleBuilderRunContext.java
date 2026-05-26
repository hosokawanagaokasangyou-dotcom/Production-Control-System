package jp.co.pm.ai.desktop.dispatch.rules.stage;

import java.nio.file.Path;
import java.util.function.Consumer;

import javafx.application.Platform;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;

/** Shared run/edit state for builder banner. */
public final class DispatchRuleBuilderRunContext {

    private static final DispatchRuleBuilderRunContext INSTANCE = new DispatchRuleBuilderRunContext();

    private final StringProperty activeStage = new SimpleStringProperty("");
    private final StringProperty snapshotId = new SimpleStringProperty("");
    private final StringProperty snapshotCapturedAt = new SimpleStringProperty("");
    private final BooleanProperty pipelineBusy = new SimpleBooleanProperty(false);
    private final BooleanProperty dirty = new SimpleBooleanProperty(false);
    private Path activeSnapshotPath;

    private Consumer<String> bannerConsumer = s -> {};

    public static DispatchRuleBuilderRunContext get() {
        return INSTANCE;
    }

    public void setBannerConsumer(Consumer<String> consumer) {
        this.bannerConsumer = consumer != null ? consumer : s -> {};
        refreshBanner();
    }

    public void beginRun(String stage, String runId, Path snapshotPath) {
        Platform.runLater(
                () -> {
                    pipelineBusy.set(true);
                    activeStage.set(stage);
                    snapshotId.set(runId);
                    snapshotCapturedAt.set(java.time.Instant.now().toString());
                    activeSnapshotPath = snapshotPath;
                    refreshBanner();
                });
    }

    public void clearActiveRun() {
        Platform.runLater(
                () -> {
                    pipelineBusy.set(false);
                    activeStage.set("");
                    snapshotId.set("");
                    snapshotCapturedAt.set("");
                    activeSnapshotPath = null;
                    refreshBanner();
                });
    }

    public void setDirty(boolean value) {
        Platform.runLater(
                () -> {
                    dirty.set(value);
                    refreshBanner();
                });
    }

    public BooleanProperty dirtyProperty() {
        return dirty;
    }

    public Path activeSnapshotPath() {
        return activeSnapshotPath;
    }

    private void refreshBanner() {
        if (pipelineBusy.get()) {
            bannerConsumer.accept(
                    "⏳ "
                            + activeStage.get()
                            + " 実行中 — 適用中: "
                            + snapshotId.get()
                            + " | 編集・保存は次回実行から | 試走はシミュレーションのみ");
        } else if (dirty.get()) {
            bannerConsumer.accept("未保存 — 次回段階1～3.5 開始時に反映。保存してください。");
        } else {
            bannerConsumer.accept("編集は次回段階1～3.5 実行開始時に適用されます。");
        }
    }
}
