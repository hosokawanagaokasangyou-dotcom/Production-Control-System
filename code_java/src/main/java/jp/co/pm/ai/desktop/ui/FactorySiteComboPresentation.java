package jp.co.pm.ai.desktop.ui;

import java.util.Map;
import java.util.function.Supplier;

import javafx.scene.control.ComboBox;
import javafx.scene.control.ListCell;
import javafx.util.StringConverter;

import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.FactorySiteOperatorAccess;

/** 工場コンボのラベル・到達／登録状態表示。 */
public final class FactorySiteComboPresentation {

    private FactorySiteComboPresentation() {}

    /** 選択中ボタン表示。共有 UNC 到達確認をしない（切替確定時に FX スレッドを止めない）。 */
    public static String selectedDisplayLabel(FactorySite site) {
        return site == null ? "" : site.displayLabelJa();
    }

    public static String labelFor(FactorySite site, Map<String, String> ui) {
        if (site == null) {
            return "";
        }
        if (!FactorySiteOperatorAccess.isSessionOperatorAllowedForFactory(ui, site)) {
            return site.displayLabelJa() + " 【ユーザー登録無し】";
        }
        return site.displayLabelJa();
    }

    public static void wire(ComboBox<FactorySite> combo, Supplier<Map<String, String>> uiSupplier) {
        if (combo == null) {
            return;
        }
        combo.setConverter(
                new StringConverter<>() {
                    @Override
                    public String toString(FactorySite site) {
                        return selectedDisplayLabel(site);
                    }

                    @Override
                    public FactorySite fromString(String string) {
                        return null;
                    }
                });
        combo.setCellFactory(
                lv ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(FactorySite item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setText(null);
                                    setDisable(false);
                                    getStyleClass().remove("pm-ai-factory-site-unregistered");
                                    return;
                                }
                                Map<String, String> ui = uiSupplier != null ? uiSupplier.get() : Map.of();
                                setText(labelFor(item, ui));
                                boolean allowed =
                                        FactorySiteOperatorAccess.isSessionOperatorAllowedForFactory(
                                                ui, item);
                                setDisable(!allowed);
                                if (!allowed) {
                                    if (!getStyleClass().contains("pm-ai-factory-site-unregistered")) {
                                        getStyleClass().add("pm-ai-factory-site-unregistered");
                                    }
                                } else {
                                    getStyleClass().remove("pm-ai-factory-site-unregistered");
                                }
                            }
                        });
        combo.setButtonCell(
                new ListCell<>() {
                    @Override
                    protected void updateItem(FactorySite item, boolean empty) {
                        super.updateItem(item, empty);
                        setText(empty ? null : selectedDisplayLabel(item));
                        setDisable(false);
                        getStyleClass().remove("pm-ai-factory-site-unregistered");
                    }
                });
    }

    public static boolean isSelectable(FactorySite site, Map<String, String> ui) {
        return site != null && FactorySiteOperatorAccess.isSessionOperatorAllowedForFactory(ui, site);
    }
}
