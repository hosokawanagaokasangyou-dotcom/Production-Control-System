package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.beans.property.ReadOnlyStringWrapper;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Accordion;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Slider;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.HBox;
import javafx.util.Duration;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.SkillsSheetMemberReader;
import jp.co.pm.ai.desktop.io.gantt.PersonNameBadgeText;
import jp.co.pm.ai.desktop.ui.PersonBadgeNodeFactory;

/** 設備ガント・担当バッジのデザイン編集タブ（master skills メンバー表＋セッション保存）。 */
public final class GanttPersonBadgeDesignTabController {

    /** 表の先頭行：マスタに無いバッジ向けの全体既定。 */
    public static final String GLOBAL_EDIT_LABEL = "（既定・マスタ外のバッジ）";

    @FXML
    private Accordion designAccordion;

    @FXML
    private Label masterWorkbookPathLabel;

    @FXML
    private Button reloadSkillsMembersButton;

    @FXML
    private TableView<BadgeDesignTableItem> badgeMemberTable;

    @FXML
    private HBox badgePreviewBox;

    @FXML
    private TextField badgeFontField;

    @FXML
    private Slider badgeFontPctSlider;

    @FXML
    private Label badgeFontPctLabel;

    @FXML
    private TextField badgeFillField;

    @FXML
    private TextField badgeTextField;

    @FXML
    private TextField badgeStrokeField;

    @FXML
    private Slider badgeStrokeSlider;

    @FXML
    private Label badgeStrokeLabel;

    @FXML
    private Slider badgeCornerSlider;

    @FXML
    private Label badgeCornerLabel;

    @FXML
    private CheckBox badgePillCheck;

    @FXML
    private TextField badgeGlowColorField;

    @FXML
    private Slider badgeGlowRadiusSlider;

    @FXML
    private Label badgeGlowRadiusLabel;

    @FXML
    private Slider badgeGlowSpreadSlider;

    @FXML
    private Label badgeGlowSpreadLabel;

    @FXML
    private Button badgeResetDefaultsButton;

    private MainShellController shell;

    private PauseTransition persistDelay;

    private boolean suppress;

    private PersonBadgeStyle globalStyle = PersonBadgeStyle.defaultStyle();

    private final LinkedHashMap<String, PersonBadgeStyle> perMemberStyles = new LinkedHashMap<>();

    /** セッション由来のバッジ文字だけの旧マップ（同一バッジが複数人にぶつかるときのフォールバック）。 */
    private Map<String, PersonBadgeStyle> legacyLabelStylesFromSession = Map.of();

    /** session-state.json にそのまま書き戻す旧キー（読込時スナップショット・変更不要）。 */
    private Map<String, PersonBadgeStyle> frozenLegacyLabelStylesForPersistence = Map.of();

    private final LinkedHashSet<String> observedBadgeLabels = new LinkedHashSet<>();

    private List<String> masterMemberNames = List.of();

    private boolean autoLoadedSkillsOnce;

    static final class BadgeDesignTableItem {
        final boolean globalFallback;
        final String memberKeyNormalized;
        final String memberDisplay;
        final String badgeShort;

        BadgeDesignTableItem(boolean globalFallback, String memberKeyNorm, String memberDisplay, String badgeShort) {
            this.globalFallback = globalFallback;
            this.memberKeyNormalized = memberKeyNorm != null ? memberKeyNorm : "";
            this.memberDisplay = memberDisplay != null ? memberDisplay : "";
            this.badgeShort = badgeShort != null ? badgeShort : "";
        }
    }

    @FXML
    private void initialize() {
        if (designAccordion != null && designAccordion.getPanes().size() > 0) {
            TitledPane p = designAccordion.getPanes().get(0);
            designAccordion.setExpandedPane(p);
        }
        persistDelay = new PauseTransition(Duration.millis(400));
        persistDelay.setOnFinished(e -> saveAndRefresh());

        setupMemberTable();

        attachListeners();
        suppress = true;
        try {
            rebuildTableItems();
            if (badgeMemberTable != null && !badgeMemberTable.getItems().isEmpty()) {
                badgeMemberTable.getSelectionModel().select(0);
            }
            pushStyleToUi(globalStyle);
            syncLabelsFromSliders();
        } finally {
            suppress = false;
        }
        refreshPreview();
    }

    private void setupMemberTable() {
        if (badgeMemberTable == null) {
            return;
        }
        badgeMemberTable.getSelectionModel().setSelectionMode(SelectionMode.SINGLE);
        badgeMemberTable
                .getSelectionModel()
                .selectedItemProperty()
                .addListener(
                        (o, prev, cur) -> {
                            if (suppress) {
                                return;
                            }
                            suppress = true;
                            try {
                                if (prev != null) {
                                    PersonBadgeStyle built = buildStyleFromUiFields();
                                    if (prev.globalFallback) {
                                        globalStyle = built;
                                    } else {
                                        perMemberStyles.put(prev.memberKeyNormalized, built);
                                    }
                                }
                                if (cur == null) {
                                    return;
                                }
                                if (cur.globalFallback) {
                                    pushStyleToUi(globalStyle);
                                } else {
                                    PersonBadgeStyle st =
                                            perMemberStyles.getOrDefault(cur.memberKeyNormalized, globalStyle);
                                    pushStyleToUi(st);
                                }
                                syncLabelsFromSliders();
                            } finally {
                                suppress = false;
                            }
                            refreshPreview();
                            if (badgeMemberTable != null) {
                                badgeMemberTable.refresh();
                            }
                        });

        TableColumn<BadgeDesignTableItem, String> colMember = new TableColumn<>("メンバー（skills）");
        colMember.setPrefWidth(220);
        colMember.setCellValueFactory(
                cd ->
                        new ReadOnlyStringWrapper(
                                cd.getValue() == null ? "" : cd.getValue().memberDisplay));

        TableColumn<BadgeDesignTableItem, String> colBadge = new TableColumn<>("バッジ表示");
        colBadge.setPrefWidth(90);
        colBadge.setCellValueFactory(
                cd ->
                        new ReadOnlyStringWrapper(
                                cd.getValue() == null ? "" : cd.getValue().badgeShort));

        TableColumn<BadgeDesignTableItem, String> colDetect = new TableColumn<>("ガント検出");
        colDetect.setPrefWidth(88);
        colDetect.setCellValueFactory(
                cd -> {
                    BadgeDesignTableItem it = cd.getValue();
                    if (it == null || it.globalFallback) {
                        return new ReadOnlyStringWrapper("—");
                    }
                    boolean ok =
                            !it.badgeShort.isEmpty()
                                    && !"—".equals(it.badgeShort)
                                    && observedBadgeLabels.contains(it.badgeShort);
                    return new ReadOnlyStringWrapper(ok ? "あり" : "");
                });

        TableColumn<BadgeDesignTableItem, String> colStyle = new TableColumn<>("スタイル");
        colStyle.setPrefWidth(100);
        colStyle.setCellValueFactory(
                cd -> {
                    BadgeDesignTableItem it = cd.getValue();
                    if (it == null) {
                        return new ReadOnlyStringWrapper("");
                    }
                    if (it.globalFallback) {
                        return new ReadOnlyStringWrapper("全体既定");
                    }
                    boolean custom = perMemberStyles.containsKey(it.memberKeyNormalized);
                    return new ReadOnlyStringWrapper(custom ? "個別" : "既定継承");
                });

        badgeMemberTable.getColumns().setAll(colMember, colBadge, colDetect, colStyle);
    }

    void bindShell(MainShellController mainShell) {
        this.shell = mainShell;
        autoLoadedSkillsOnce = false;
        Platform.runLater(this::trySilentInitialMasterLoad);
    }

    /** 設備ガントで検出したバッジ表示文字を「ガント検出」列に反映する。 */
    void mergeObservedBadgeLabels(Collection<String> labels) {
        if (labels == null || labels.isEmpty()) {
            return;
        }
        boolean changed = false;
        for (String raw : labels) {
            String k = PersonBadgeStyle.normalizeLabelKey(raw);
            if (!k.isEmpty() && observedBadgeLabels.add(k)) {
                changed = true;
            }
        }
        if (changed && badgeMemberTable != null) {
            badgeMemberTable.refresh();
        }
    }

    void applyPersonBadgeDesignSession(DesktopSessionState s) {
        if (s == null) {
            return;
        }
        suppress = true;
        try {
            globalStyle = s.resolvedPersonBadgeStyle();
            perMemberStyles.clear();
            perMemberStyles.putAll(s.equipmentGanttPersonBadgeStylesByMemberKey());
            legacyLabelStylesFromSession = Map.copyOf(s.equipmentGanttPersonBadgeStylesByLabel());
            frozenLegacyLabelStylesForPersistence = Map.copyOf(s.equipmentGanttPersonBadgeStylesByLabel());
            rebuildTableItems();
            if (badgeMemberTable != null && !badgeMemberTable.getItems().isEmpty()) {
                badgeMemberTable.getSelectionModel().select(0);
            }
            pushStyleToUi(globalStyle);
            syncLabelsFromSliders();
        } finally {
            suppress = false;
        }
        refreshPreview();
        autoLoadedSkillsOnce = false;
        Platform.runLater(this::trySilentInitialMasterLoad);
    }

    void flushBadgeEditsBeforeSnapshot() {
        commitUiToTarget();
    }

    public PersonBadgeStyle resolveStyleForBadgeLabel(String badgeFragment) {
        commitUiToTarget();
        return resolveStyleForBadgeLabelRaw(badgeFragment);
    }

    private PersonBadgeStyle resolveStyleForBadgeLabelRaw(String badgeFragment) {
        String frag = PersonBadgeStyle.normalizeLabelKey(badgeFragment);
        if (frag.isEmpty()) {
            return globalStyle;
        }
        LinkedHashSet<String> matchKeys = new LinkedHashSet<>();
        for (String raw : masterMemberNames) {
            String badge = PersonNameBadgeText.badgeTwoFromRawName(raw);
            if (badge.isEmpty()) {
                continue;
            }
            if (frag.equals(PersonBadgeStyle.normalizeLabelKey(badge))) {
                matchKeys.add(PersonBadgeStyle.normalizeLabelKey(raw));
            }
        }
        if (matchKeys.size() == 1) {
            PersonBadgeStyle st = perMemberStyles.get(matchKeys.iterator().next());
            return st != null ? st : globalStyle;
        }
        if (matchKeys.size() > 1) {
            PersonBadgeStyle amb = legacyLabelStylesFromSession.get(frag);
            return amb != null ? amb : globalStyle;
        }
        PersonBadgeStyle leg = legacyLabelStylesFromSession.get(frag);
        return leg != null ? leg : globalStyle;
    }

    Map<String, PersonBadgeStyle> snapshotPersonBadgeStylesByLabel() {
        return Map.copyOf(frozenLegacyLabelStylesForPersistence);
    }

    Map<String, PersonBadgeStyle> snapshotPersonBadgeStylesByMemberKey() {
        commitUiToTarget();
        return Map.copyOf(perMemberStyles);
    }

    PersonBadgeStyle previewStyleForGantt() {
        return globalStyle;
    }

    String snapshotPersonBadgeFontFamily() {
        return globalStyle.fontFamily() != null ? globalStyle.fontFamily().strip() : "";
    }

    double snapshotPersonBadgeFontPercent() {
        return globalStyle.fontPercent();
    }

    String snapshotPersonBadgeFillHex() {
        return globalStyle.fillHex() != null ? globalStyle.fillHex() : "";
    }

    String snapshotPersonBadgeTextHex() {
        return globalStyle.textHex() != null ? globalStyle.textHex() : "";
    }

    String snapshotPersonBadgeStrokeHex() {
        return globalStyle.strokeHex() != null ? globalStyle.strokeHex() : "";
    }

    double snapshotPersonBadgeStrokeWidth() {
        return globalStyle.strokeWidth();
    }

    double snapshotPersonBadgeCornerRadius() {
        return globalStyle.cornerRadius();
    }

    boolean snapshotPersonBadgePill() {
        return globalStyle.pill();
    }

    String snapshotPersonBadgeGlowColorHex() {
        return globalStyle.glowColorHex() != null ? globalStyle.glowColorHex() : "";
    }

    double snapshotPersonBadgeGlowRadius() {
        return globalStyle.glowRadius();
    }

    double snapshotPersonBadgeGlowSpread() {
        return globalStyle.glowSpread();
    }

    @FXML
    private void onReloadSkillsMembersAction() {
        loadMembersFromMasterBook(true);
    }

    private void trySilentInitialMasterLoad() {
        if (shell == null || autoLoadedSkillsOnce) {
            updateMasterPathHint();
            return;
        }
        loadMembersFromMasterBook(false);
    }

    private void loadMembersFromMasterBook(boolean showAlertOnFailure) {
        if (shell == null) {
            return;
        }
        Path master = shell.resolveMasterWorkbookIfPresent();
        updateMasterPathHint(master);
        if (master == null) {
            if (showAlertOnFailure) {
                alert(
                        Alert.AlertType.WARNING,
                        "マスタブックが見つかりません",
                        "環境変数 PM_AI_MASTER_WORKBOOK / MASTER_WORKBOOK_FILE と、実行タブのブックパスを確認してください。");
            }
            return;
        }
        try {
            masterMemberNames = SkillsSheetMemberReader.readMemberDisplayNames(master);
            autoLoadedSkillsOnce = true;
            suppress = true;
            try {
                rebuildTableItems();
                if (badgeMemberTable != null && !badgeMemberTable.getItems().isEmpty()) {
                    badgeMemberTable.getSelectionModel().select(0);
                    BadgeDesignTableItem it = badgeMemberTable.getSelectionModel().getSelectedItem();
                    if (it != null) {
                        if (it.globalFallback) {
                            pushStyleToUi(globalStyle);
                        } else {
                            pushStyleToUi(perMemberStyles.getOrDefault(it.memberKeyNormalized, globalStyle));
                        }
                        syncLabelsFromSliders();
                    }
                }
            } finally {
                suppress = false;
            }
            refreshPreview();
            if (badgeMemberTable != null) {
                badgeMemberTable.refresh();
            }
        } catch (IOException ex) {
            if (showAlertOnFailure) {
                alert(
                        Alert.AlertType.ERROR,
                        "skills の読込に失敗",
                        ex.getMessage() != null ? ex.getMessage() : ex.toString());
            }
        }
    }

    private void updateMasterPathHint() {
        updateMasterPathHint(shell != null ? shell.resolveMasterWorkbookIfPresent() : null);
    }

    private void updateMasterPathHint(Path master) {
        if (masterWorkbookPathLabel == null) {
            return;
        }
        if (master == null) {
            masterWorkbookPathLabel.setText("マスタ未検出（環境変数・実行タブを確認）");
        } else {
            masterWorkbookPathLabel.setText(master.toString());
        }
    }

    private static void alert(Alert.AlertType type, String title, String msg) {
        Alert a = new Alert(type, msg);
        a.setHeaderText(title);
        a.showAndWait();
    }

    private void rebuildTableItems() {
        if (badgeMemberTable == null) {
            return;
        }
        BadgeDesignTableItem prevSel =
                badgeMemberTable.getSelectionModel().getSelectedItem();
        String prevKey =
                prevSel != null && !prevSel.globalFallback ? prevSel.memberKeyNormalized : null;

        ObservableList<BadgeDesignTableItem> items = FXCollections.observableArrayList();
        items.add(new BadgeDesignTableItem(true, "", GLOBAL_EDIT_LABEL, "—"));
        for (String raw : masterMemberNames) {
            String mk = PersonBadgeStyle.normalizeLabelKey(raw);
            String badge = PersonNameBadgeText.badgeTwoFromRawName(raw);
            items.add(
                    new BadgeDesignTableItem(
                            false,
                            mk,
                            raw,
                            badge.isEmpty() ? "—" : badge));
        }
        badgeMemberTable.setItems(items);

        if (prevKey != null) {
            for (BadgeDesignTableItem it : items) {
                if (!it.globalFallback && prevKey.equals(it.memberKeyNormalized)) {
                    badgeMemberTable.getSelectionModel().select(it);
                    return;
                }
            }
        }
        badgeMemberTable.getSelectionModel().select(0);
    }

    private void commitUiToTarget() {
        if (suppress || badgeMemberTable == null) {
            return;
        }
        BadgeDesignTableItem sel = badgeMemberTable.getSelectionModel().getSelectedItem();
        if (sel == null) {
            return;
        }
        PersonBadgeStyle built = buildStyleFromUiFields();
        if (sel.globalFallback) {
            globalStyle = built;
        } else {
            perMemberStyles.put(sel.memberKeyNormalized, built);
        }
        badgeMemberTable.refresh();
    }

    private void pushStyleToUi(PersonBadgeStyle st) {
        PersonBadgeStyle x = st != null ? st : PersonBadgeStyle.defaultStyle();
        if (badgeFontField != null) {
            badgeFontField.setText(x.fontFamily() != null ? x.fontFamily() : "");
        }
        if (badgeFontPctSlider != null) {
            badgeFontPctSlider.setValue(Math.clamp(x.fontPercent(), 40, 160));
        }
        if (badgeFillField != null) {
            badgeFillField.setText(x.fillHex());
        }
        if (badgeTextField != null) {
            badgeTextField.setText(x.textHex());
        }
        if (badgeStrokeField != null) {
            badgeStrokeField.setText(x.strokeHex());
        }
        if (badgeStrokeSlider != null) {
            badgeStrokeSlider.setValue(Math.clamp(x.strokeWidth(), 0, 6));
        }
        if (badgeCornerSlider != null) {
            badgeCornerSlider.setValue(Math.clamp(x.cornerRadius(), 0, 24));
        }
        if (badgePillCheck != null) {
            badgePillCheck.setSelected(x.pill());
        }
        if (badgeGlowColorField != null) {
            badgeGlowColorField.setText(x.glowColorHex());
        }
        if (badgeGlowRadiusSlider != null) {
            badgeGlowRadiusSlider.setValue(Math.clamp(x.glowRadius(), 0, 40));
        }
        if (badgeGlowSpreadSlider != null) {
            badgeGlowSpreadSlider.setValue(Math.clamp(x.glowSpread(), 0, 1));
        }
    }

    private void attachListeners() {
        Runnable r = this::schedulePersist;
        if (badgeFontField != null) {
            badgeFontField.textProperty().addListener((o, a, b) -> r.run());
        }
        wireSlider(badgeFontPctSlider, badgeFontPctLabel, "%.0f%%", r);
        if (badgeFillField != null) {
            badgeFillField.textProperty().addListener((o, a, b) -> r.run());
        }
        if (badgeTextField != null) {
            badgeTextField.textProperty().addListener((o, a, b) -> r.run());
        }
        if (badgeStrokeField != null) {
            badgeStrokeField.textProperty().addListener((o, a, b) -> r.run());
        }
        wireSlider(badgeStrokeSlider, badgeStrokeLabel, "%.1f", r);
        wireSlider(badgeCornerSlider, badgeCornerLabel, "%.0f", r);
        if (badgePillCheck != null) {
            badgePillCheck.selectedProperty().addListener((o, a, b) -> r.run());
        }
        if (badgeGlowColorField != null) {
            badgeGlowColorField.textProperty().addListener((o, a, b) -> r.run());
        }
        wireSlider(badgeGlowRadiusSlider, badgeGlowRadiusLabel, "%.0f", r);
        wireSlider(badgeGlowSpreadSlider, badgeGlowSpreadLabel, "%.2f", r);
    }

    private static void wireSlider(Slider sl, Label lb, String fmt, Runnable onChange) {
        if (sl == null) {
            return;
        }
        sl.valueProperty()
                .addListener(
                        (o, a, b) -> {
                            if (lb != null) {
                                lb.setText(String.format(fmt, sl.getValue()));
                            }
                            onChange.run();
                        });
    }

    private void syncLabelsFromSliders() {
        if (badgeFontPctSlider != null && badgeFontPctLabel != null) {
            badgeFontPctLabel.setText(String.format("%.0f%%", badgeFontPctSlider.getValue()));
        }
        if (badgeStrokeSlider != null && badgeStrokeLabel != null) {
            badgeStrokeLabel.setText(String.format("%.1f", badgeStrokeSlider.getValue()));
        }
        if (badgeCornerSlider != null && badgeCornerLabel != null) {
            badgeCornerLabel.setText(String.format("%.0f", badgeCornerSlider.getValue()));
        }
        if (badgeGlowRadiusSlider != null && badgeGlowRadiusLabel != null) {
            badgeGlowRadiusLabel.setText(String.format("%.0f", badgeGlowRadiusSlider.getValue()));
        }
        if (badgeGlowSpreadSlider != null && badgeGlowSpreadLabel != null) {
            badgeGlowSpreadLabel.setText(String.format("%.2f", badgeGlowSpreadSlider.getValue()));
        }
    }

    private void schedulePersist() {
        if (suppress) {
            return;
        }
        refreshPreview();
        persistDelay.stop();
        persistDelay.playFromStart();
    }

    private void saveAndRefresh() {
        if (suppress) {
            return;
        }
        if (shell != null) {
            shell.persistDesktopSessionNow();
            shell.refreshEquipmentGanttGraphicForBadgeChange();
        }
    }

    private static String nz(TextField f) {
        return f != null && f.getText() != null ? f.getText().strip() : "";
    }

    private static String nzOr(TextField f, String fallback) {
        String s = nz(f);
        return s.isEmpty() ? fallback : s;
    }

    private PersonBadgeStyle buildStyleFromUiFields() {
        PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
        return new PersonBadgeStyle(
                nz(badgeFontField),
                badgeFontPctSlider != null ? badgeFontPctSlider.getValue() : d.fontPercent(),
                nzOr(badgeFillField, d.fillHex()),
                nzOr(badgeTextField, d.textHex()),
                nzOr(badgeStrokeField, d.strokeHex()),
                badgeStrokeSlider != null ? badgeStrokeSlider.getValue() : d.strokeWidth(),
                badgeCornerSlider != null ? badgeCornerSlider.getValue() : d.cornerRadius(),
                badgePillCheck != null && badgePillCheck.isSelected(),
                nzOr(badgeGlowColorField, d.glowColorHex()),
                badgeGlowRadiusSlider != null ? badgeGlowRadiusSlider.getValue() : d.glowRadius(),
                badgeGlowSpreadSlider != null ? badgeGlowSpreadSlider.getValue() : d.glowSpread());
    }

    @FXML
    private void onBadgeResetDefaultsAction() {
        BadgeDesignTableItem sel =
                badgeMemberTable != null ? badgeMemberTable.getSelectionModel().getSelectedItem() : null;
        suppress = true;
        try {
            PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
            if (sel == null) {
                globalStyle = d;
                pushStyleToUi(globalStyle);
            } else if (sel.globalFallback) {
                globalStyle = d;
                pushStyleToUi(globalStyle);
            } else {
                perMemberStyles.remove(sel.memberKeyNormalized);
                pushStyleToUi(globalStyle);
            }
            syncLabelsFromSliders();
        } finally {
            suppress = false;
        }
        refreshPreview();
        if (badgeMemberTable != null) {
            badgeMemberTable.refresh();
        }
        saveAndRefresh();
    }

    private void refreshPreview() {
        if (badgePreviewBox == null) {
            return;
        }
        if (!suppress) {
            commitUiToTarget();
        }
        badgePreviewBox.getChildren().clear();
        double zoom = 1.0;
        double rowPx = 13;
        List<String> samples = new ArrayList<>();
        for (String raw : masterMemberNames) {
            String b = PersonNameBadgeText.badgeTwoFromRawName(raw);
            if (!b.isEmpty()) {
                samples.add(b);
            }
            if (samples.size() >= 2) {
                break;
            }
        }
        while (samples.size() < 2) {
            if (samples.isEmpty()) {
                samples.add("山田");
            } else if (samples.size() == 1 && !samples.contains("佐藤")) {
                samples.add("佐藤");
            } else {
                samples.add("?");
            }
        }
        for (int i = 0; i < 2; i++) {
            String s = samples.get(i);
            PersonBadgeStyle st = resolveStyleForBadgeLabelRaw(s);
            badgePreviewBox.getChildren().add(PersonBadgeNodeFactory.createBadge(s, st, zoom, rowPx));
        }
    }
}
