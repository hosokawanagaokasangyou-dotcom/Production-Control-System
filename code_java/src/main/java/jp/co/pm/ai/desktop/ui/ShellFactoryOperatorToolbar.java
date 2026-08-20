package jp.co.pm.ai.desktop.ui;

import java.util.Map;

import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Tooltip;
import javafx.scene.image.ImageView;
import javafx.scene.layout.StackPane;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.FactorySiteLogoSupport;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;

/** メインシェル最上部ツールバー左端の工場・操作者 UI。 */
public final class ShellFactoryOperatorToolbar {

    private final StackPane factoryLogoHost;
    private final ImageView factoryLogoImageView;
    private final Label factoryLogoCaptionLabel;
    private final ComboBox<FactorySite> factorySiteCombo;
    private final Label operatorUserLabel;
    private final Button changeSessionOperatorButton;
    private final Button changeOperatorPinButton;

    private MainShellController shell;
    private boolean suppressFactorySiteComboEvents;
    private boolean factorySiteComboDisabledByShell;

    public ShellFactoryOperatorToolbar(
            StackPane factoryLogoHost,
            ImageView factoryLogoImageView,
            Label factoryLogoCaptionLabel,
            ComboBox<FactorySite> factorySiteCombo,
            Label operatorUserLabel,
            Button changeSessionOperatorButton,
            Button changeOperatorPinButton) {
        this.factoryLogoHost = factoryLogoHost;
        this.factoryLogoImageView = factoryLogoImageView;
        this.factoryLogoCaptionLabel = factoryLogoCaptionLabel;
        this.factorySiteCombo = factorySiteCombo;
        this.operatorUserLabel = operatorUserLabel;
        this.changeSessionOperatorButton = changeSessionOperatorButton;
        this.changeOperatorPinButton = changeOperatorPinButton;
    }

    public void wire(MainShellController shell) {
        this.shell = shell;
        wireFactorySiteCombo();
        refreshFactorySiteLogo();
        refreshOperatorUserLabel();
    }

    public void refreshFactorySiteLogo() {
        if (factoryLogoHost == null || factoryLogoCaptionLabel == null) {
            return;
        }
        FactorySite site = GlobalInitSettingTarget.load();
        factoryLogoCaptionLabel.setText(site.displayLabelJa());
        factoryLogoHost.getStyleClass().removeIf(c -> c.startsWith("pm-factory-logo-"));
        factoryLogoHost.getStyleClass().add("pm-factory-logo-" + site.name().toLowerCase());
        Map<String, String> ui = shell != null ? shell.snapshotUiEnv() : Map.of();
        if (factoryLogoImageView != null) {
            FactorySiteLogoSupport.applyBrandingOverrideToImageView(factoryLogoImageView, site, ui);
        }
        boolean branding = factoryLogoImageView != null && factoryLogoImageView.isVisible();
        factoryLogoCaptionLabel.setVisible(!branding);
        factoryLogoCaptionLabel.setManaged(!branding);
        Tooltip.install(factoryLogoHost, new Tooltip(site.displayLabelJa() + "（利用工場）"));
    }

    public void refreshFactorySiteComboPresentation() {
        if (factorySiteCombo != null) {
            factorySiteCombo.requestLayout();
            factorySiteCombo.getSelectionModel().select(factorySiteCombo.getValue());
        }
    }

    public void setFactorySiteComboDisabled(boolean disabled) {
        factorySiteComboDisabledByShell = disabled;
        refreshFactorySiteComboDisabledState();
    }

    public void setGuestSessionFactorySwitchOnly(boolean guestOnly) {
        refreshFactorySiteComboDisabledState();
        refreshOperatorUserLabel();
    }

    public void refreshFactorySiteComboFromStore() {
        if (factorySiteCombo == null) {
            return;
        }
        FactorySite disk = GlobalInitSettingTarget.load();
        if (factorySiteCombo.getValue() != disk) {
            suppressFactorySiteComboEvents = true;
            try {
                factorySiteCombo.setValue(disk);
            } finally {
                suppressFactorySiteComboEvents = false;
            }
        }
    }

    public void refreshOperatorUserLabel() {
        if (operatorUserLabel == null) {
            return;
        }
        String op = FactoryOperatorUserStore.sessionOperatorName();
        operatorUserLabel.setText(op.isBlank() ? "操作者: （未選択）" : "操作者: " + op);
        boolean disablePinChange = op.isBlank() || FactoryOperatorUserStore.isGuestSession();
        if (changeOperatorPinButton != null) {
            changeOperatorPinButton.setDisable(disablePinChange);
        }
        if (changeSessionOperatorButton != null) {
            changeSessionOperatorButton.setDisable(false);
        }
    }

    private void wireFactorySiteCombo() {
        if (factorySiteCombo == null) {
            return;
        }
        factorySiteCombo.getItems().setAll(FactorySite.dispatchProductionSites());
        FactorySiteComboPresentation.wire(
                factorySiteCombo, () -> shell != null ? shell.snapshotUiEnv() : Map.of());
        factorySiteCombo
                .valueProperty()
                .addListener(
                        (obs, oldV, newV) -> {
                            if (suppressFactorySiteComboEvents || newV == null || shell == null) {
                                return;
                            }
                            shell.switchActiveFactorySite(newV);
                        });
        refreshFactorySiteComboFromStore();
    }

    private void refreshFactorySiteComboDisabledState() {
        if (factorySiteCombo != null) {
            factorySiteCombo.setDisable(factorySiteComboDisabledByShell);
        }
    }
}
