package jp.co.pm.ai.desktop.kouchin;

import java.util.Locale;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.TabPane;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.util.Duration;
import jp.co.pm.ai.desktop.MainShellController;

/**
 * 「後加工工賃」メインタブのホスト。子タブ「検証」「月次トレンド」「参照先」。
 * 検証・トレンド実行中はタブ中央にプログレス付きオーバーレイを出す。
 */
public class KouchinHostTabController {

    @FXML private TabPane innerTabPane;
    @FXML private StackPane busyOverlay;
    @FXML private Region busyBackdrop;
    @FXML private ProgressIndicator busyIndicator;
    @FXML private ProgressBar busyBar;
    @FXML private Label busyLabel;
    @FXML private Label busyElapsedLabel;
    @FXML private Button busyCancelButton;
    @FXML private KouchinVerifyTabController verifyTabController;
    @FXML private KouchinTrendTabController trendTabController;
    @FXML private KouchinSourcesTabController sourcesTabController;

    private MainShellController shell;
    private Timeline busyTick;
    private long busyStartMs;

    @FXML
    private void initialize() {
        if (innerTabPane != null) {
            innerTabPane.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
            innerTabPane
                    .getSelectionModel()
                    .selectedItemProperty()
                    .addListener((obs, o, n) -> notifyActiveChildSelected());
        }
        if (busyBackdrop != null) {
            busyBackdrop.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
            busyBackdrop.setMouseTransparent(true);
        }
        if (busyOverlay != null) {
            busyOverlay.setMaxSize(Double.MAX_VALUE, Double.MAX_VALUE);
            setRunBusy(false, "");
        }
        busyTick =
                new Timeline(
                        new KeyFrame(
                                Duration.millis(400),
                                e -> updateBusyElapsed()));
        busyTick.setCycleCount(Timeline.INDEFINITE);
    }

    public void bindShell(MainShellController shell) {
        this.shell = shell;
        if (verifyTabController != null) {
            verifyTabController.bindShell(shell, this);
        }
        if (trendTabController != null) {
            trendTabController.bindShell(shell, this);
        }
        if (sourcesTabController != null) {
            sourcesTabController.bindShell(shell, this);
        }
    }

    public void onMainShellTabSelected() {
        notifyActiveChildSelected();
    }

    public void onMainShellTabDeselected() {
        if (verifyTabController != null) {
            verifyTabController.onMainShellTabDeselected();
        }
        if (trendTabController != null) {
            trendTabController.onMainShellTabDeselected();
        }
        if (sourcesTabController != null) {
            sourcesTabController.onMainShellTabDeselected();
        }
    }

    public boolean hasUnappliedSourceEdits() {
        return sourcesTabController != null && sourcesTabController.hasUnappliedEdits();
    }

    public void onSourcesApplied() {
        if (verifyTabController != null) {
            verifyTabController.markUnverified();
            verifyTabController.reloadDiscovery();
        }
    }

    public KouchinVerifyTabController verifyTab() {
        return verifyTabController;
    }

    public KouchinTrendTabController trendTab() {
        return trendTabController;
    }

    public KouchinSourcesTabController sourcesTab() {
        return sourcesTabController;
    }

    /**
     * 検証／月次トレンド実行中のタブ中央モーダル。下部ステータスバーと併用する。
     */
    public void setRunBusy(boolean busy, String message) {
        if (busyOverlay == null) {
            return;
        }
        if (busy) {
            String text = message == null || message.isBlank() ? "後加工工賃 検証中…" : message;
            if (busyLabel != null) {
                busyLabel.setText(text);
            }
            if (busyBar != null) {
                busyBar.setProgress(ProgressBar.INDETERMINATE_PROGRESS);
            }
            if (busyCancelButton != null) {
                busyCancelButton.setDisable(false);
            }
            busyStartMs = System.currentTimeMillis();
            updateBusyElapsed();
            if (busyTick != null) {
                busyTick.playFromStart();
            }
            busyOverlay.setVisible(true);
            busyOverlay.setMouseTransparent(false);
        } else {
            if (busyTick != null) {
                busyTick.stop();
            }
            if (busyElapsedLabel != null) {
                busyElapsedLabel.setText("");
            }
            busyOverlay.setVisible(false);
            busyOverlay.setMouseTransparent(true);
        }
        if (innerTabPane != null) {
            innerTabPane.setDisable(busy);
        }
        if (verifyTabController != null) {
            verifyTabController.refreshRunEnabled();
        }
        if (trendTabController != null) {
            trendTabController.refreshRunEnabled();
        }
    }

    @FXML
    private void onBusyCancel() {
        if (shell != null) {
            shell.requestKouchinCancel();
        }
        if (busyCancelButton != null) {
            busyCancelButton.setDisable(true);
        }
        if (busyLabel != null) {
            busyLabel.setText("中断を要求しました…");
        }
    }

    private void updateBusyElapsed() {
        if (busyElapsedLabel == null) {
            return;
        }
        double sec = (System.currentTimeMillis() - busyStartMs) / 1000.0;
        busyElapsedLabel.setText(String.format(Locale.ROOT, "経過 %.1f 秒（処理中）", sec));
    }

    private void notifyActiveChildSelected() {
        if (innerTabPane == null) {
            return;
        }
        int idx = innerTabPane.getSelectionModel().getSelectedIndex();
        if (verifyTabController != null) {
            if (idx == 0) {
                verifyTabController.onMainShellTabSelected();
            } else {
                verifyTabController.onMainShellTabDeselected();
            }
        }
        if (trendTabController != null) {
            if (idx == 1) {
                trendTabController.onMainShellTabSelected();
            } else {
                trendTabController.onMainShellTabDeselected();
            }
        }
        if (sourcesTabController != null) {
            if (idx == 2) {
                sourcesTabController.onMainShellTabSelected();
            } else {
                sourcesTabController.onMainShellTabDeselected();
            }
        }
    }
}
