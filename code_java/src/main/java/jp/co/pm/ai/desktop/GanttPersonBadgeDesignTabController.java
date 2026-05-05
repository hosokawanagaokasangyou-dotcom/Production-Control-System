package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ThreadLocalRandom;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.beans.property.ReadOnlyObjectWrapper;
import javafx.beans.property.ReadOnlyStringWrapper;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Accordion;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ColorPicker;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.Slider;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TitledPane;
import javafx.scene.layout.HBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
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

    private static final String FONT_COMBO_DEFAULT_LABEL = "(既定)";

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
    private ComboBox<String> badgeFontCombo;

    @FXML
    private Slider badgeFontPctSlider;

    @FXML
    private Label badgeFontPctLabel;

    @FXML
    private ColorPicker badgeFillPicker;

    @FXML
    private ColorPicker badgeTextPicker;

    @FXML
    private ColorPicker badgeStrokePicker;

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
    private ColorPicker badgeGlowColorPicker;

    @FXML
    private Slider badgeGlowRadiusPctSlider;

    @FXML
    private Label badgeGlowRadiusPctLabel;

    @FXML
    private Slider badgeGlowSpreadPctSlider;

    @FXML
    private Label badgeGlowSpreadPctLabel;

    @FXML
    private Button badgeRandomizeButton;

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

        populateFontCombo();

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

        TableColumn<BadgeDesignTableItem, BadgeDesignTableItem> colPreview = new TableColumn<>("プレビュー");
        colPreview.setPrefWidth(112);
        colPreview.setCellValueFactory(cd -> new ReadOnlyObjectWrapper<>(cd.getValue()));
        colPreview.setCellFactory(
                col ->
                        new TableCell<BadgeDesignTableItem, BadgeDesignTableItem>() {
                            @Override
                            protected void updateItem(BadgeDesignTableItem item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setText(null);
                                    setGraphic(null);
                                    return;
                                }
                                setText(null);
                                PersonBadgeStyle st = styleForTableRowPreview(item);
                                String txt = previewBadgeText(item);
                                setGraphic(PersonBadgeNodeFactory.createBadge(txt, st, 1.0, 12.0));
                            }
                        });

        badgeMemberTable.getColumns().setAll(colMember, colBadge, colDetect, colStyle, colPreview);
    }

    /** 表のプレビュー列用：選択行は UI の現在値、それ以外は保存済みの解決スタイル。 */
    private PersonBadgeStyle styleForTableRowPreview(BadgeDesignTableItem item) {
        if (item == null) {
            return PersonBadgeStyle.defaultStyle();
        }
        BadgeDesignTableItem sel =
                badgeMemberTable != null ? badgeMemberTable.getSelectionModel().getSelectedItem() : null;
        if (sel != null && item == sel) {
            return buildStyleFromUiFields();
        }
        PersonBadgeStyle g = globalStyle != null ? globalStyle : PersonBadgeStyle.defaultStyle();
        if (item.globalFallback) {
            return g;
        }
        return perMemberStyles.getOrDefault(item.memberKeyNormalized, g);
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
        PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
        applyFontFamilyToCombo(x.fontFamily());
        if (badgeFontPctSlider != null) {
            badgeFontPctSlider.setValue(Math.clamp(x.fontPercent(), 40, 160));
        }
        if (badgeFillPicker != null) {
            badgeFillPicker.setValue(parseHexToColor(x.fillHex(), Color.web(d.fillHex())));
        }
        if (badgeTextPicker != null) {
            badgeTextPicker.setValue(parseHexToColor(x.textHex(), Color.web(d.textHex())));
        }
        if (badgeStrokePicker != null) {
            badgeStrokePicker.setValue(parseHexToColor(x.strokeHex(), Color.web(d.strokeHex())));
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
        if (badgeGlowColorPicker != null) {
            badgeGlowColorPicker.setValue(parseHexToColor(x.glowColorHex(), Color.web(d.glowColorHex())));
        }
        double baseR = d.glowRadius();
        double baseS = d.glowSpread();
        if (badgeGlowRadiusPctSlider != null) {
            double pct = baseR > 1e-9 ? (x.glowRadius() / baseR) * 100.0 : 0.0;
            badgeGlowRadiusPctSlider.setValue(Math.clamp(pct, 0, 400));
        }
        if (badgeGlowSpreadPctSlider != null) {
            double pct =
                    baseS > 1e-12
                            ? (x.glowSpread() / baseS) * 100.0
                            : (x.glowSpread() <= 1e-12 ? 0.0 : 100.0);
            badgeGlowSpreadPctSlider.setValue(Math.clamp(pct, 0, 400));
        }
    }

    private void attachListeners() {
        Runnable r = this::schedulePersist;
        if (badgeFontCombo != null) {
            badgeFontCombo.valueProperty().addListener((o, a, b) -> r.run());
        }
        wireSlider(badgeFontPctSlider, badgeFontPctLabel, "%.0f%%", r);
        addColorPickerListener(badgeFillPicker, r);
        addColorPickerListener(badgeTextPicker, r);
        addColorPickerListener(badgeStrokePicker, r);
        wireSlider(badgeStrokeSlider, badgeStrokeLabel, "%.1f", r);
        wireSlider(badgeCornerSlider, badgeCornerLabel, "%.0f", r);
        if (badgePillCheck != null) {
            badgePillCheck.selectedProperty().addListener((o, a, b) -> r.run());
        }
        addColorPickerListener(badgeGlowColorPicker, r);
        wireSlider(badgeGlowRadiusPctSlider, badgeGlowRadiusPctLabel, "%.0f%%", r);
        wireSlider(badgeGlowSpreadPctSlider, badgeGlowSpreadPctLabel, "%.0f%%", r);
    }

    private static void addColorPickerListener(ColorPicker cp, Runnable r) {
        if (cp != null) {
            cp.valueProperty().addListener((o, a, b) -> r.run());
        }
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
        if (badgeGlowRadiusPctSlider != null && badgeGlowRadiusPctLabel != null) {
            badgeGlowRadiusPctLabel.setText(String.format("%.0f%%", badgeGlowRadiusPctSlider.getValue()));
        }
        if (badgeGlowSpreadPctSlider != null && badgeGlowSpreadPctLabel != null) {
            badgeGlowSpreadPctLabel.setText(String.format("%.0f%%", badgeGlowSpreadPctSlider.getValue()));
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

    private PersonBadgeStyle buildStyleFromUiFields() {
        PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
        String fontFam = "";
        if (badgeFontCombo != null && badgeFontCombo.getValue() != null) {
            String v = badgeFontCombo.getValue().strip();
            if (!v.isEmpty() && !FONT_COMBO_DEFAULT_LABEL.equals(v)) {
                fontFam = v;
            }
        }
        double baseR = d.glowRadius();
        double baseS = d.glowSpread();
        double rPct = badgeGlowRadiusPctSlider != null ? badgeGlowRadiusPctSlider.getValue() : 100.0;
        double sPct = badgeGlowSpreadPctSlider != null ? badgeGlowSpreadPctSlider.getValue() : 100.0;
        double glowR = baseR * (rPct / 100.0);
        double glowS = Math.min(1.0, Math.max(0.0, baseS * (sPct / 100.0)));
        return new PersonBadgeStyle(
                fontFam,
                badgeFontPctSlider != null ? badgeFontPctSlider.getValue() : d.fontPercent(),
                colorToHex(badgeFillPicker, d.fillHex()),
                colorToHex(badgeTextPicker, d.textHex()),
                colorToHex(badgeStrokePicker, d.strokeHex()),
                badgeStrokeSlider != null ? badgeStrokeSlider.getValue() : d.strokeWidth(),
                badgeCornerSlider != null ? badgeCornerSlider.getValue() : d.cornerRadius(),
                badgePillCheck != null && badgePillCheck.isSelected(),
                colorToHex(badgeGlowColorPicker, d.glowColorHex()),
                glowR,
                glowS);
    }

    private static String colorToHex(ColorPicker cp, String fallbackHex) {
        if (cp == null || cp.getValue() == null) {
            return fallbackHex;
        }
        return colorToHex(cp.getValue());
    }

    private static String colorToHex(Color c) {
        int r = (int) Math.round(c.getRed() * 255.0);
        int g = (int) Math.round(c.getGreen() * 255.0);
        int b = (int) Math.round(c.getBlue() * 255.0);
        return String.format("#%02x%02x%02x", r, g, b);
    }

    private static Color parseHexToColor(String hex, Color fallback) {
        try {
            String h = hex != null ? hex.strip() : "";
            return h.isEmpty() ? fallback : Color.web(h);
        } catch (IllegalArgumentException e) {
            return fallback;
        }
    }

    private void populateFontCombo() {
        if (badgeFontCombo == null) {
            return;
        }
        ObservableList<String> items = FXCollections.observableArrayList(FONT_COMBO_DEFAULT_LABEL);
        List<String> sorted = new ArrayList<>(Font.getFamilies());
        sorted.sort(String.CASE_INSENSITIVE_ORDER);
        items.addAll(sorted);
        badgeFontCombo.setItems(items);
    }

    private void applyFontFamilyToCombo(String fontFamily) {
        if (badgeFontCombo == null) {
            return;
        }
        if (fontFamily == null || fontFamily.isBlank()) {
            badgeFontCombo.getSelectionModel().select(0);
            return;
        }
        ObservableList<String> items = badgeFontCombo.getItems();
        for (int i = 0; i < items.size(); i++) {
            if (fontFamily.equals(items.get(i))) {
                badgeFontCombo.getSelectionModel().select(i);
                return;
            }
        }
        items.add(fontFamily);
        badgeFontCombo.getSelectionModel().select(fontFamily);
    }

    @FXML
    private void onBadgeRandomizeAction() {
        ThreadLocalRandom rnd = ThreadLocalRandom.current();
        suppress = true;
        try {
            Color fill = randomBadgeFillColor(rnd);
            if (badgeFillPicker != null) {
                badgeFillPicker.setValue(fill);
            }
            if (badgeTextPicker != null) {
                badgeTextPicker.setValue(randomBadgeTextColor(rnd, fill));
            }
            if (badgeStrokePicker != null) {
                badgeStrokePicker.setValue(randomBadgeStrokeColor(rnd, fill));
            }
            if (badgeGlowColorPicker != null) {
                badgeGlowColorPicker.setValue(randomGlowColor(rnd, fill));
            }
            if (badgeFontCombo != null) {
                ObservableList<String> items = badgeFontCombo.getItems();
                if (items != null && items.size() > 1) {
                    int n = items.size();
                    int idx = rnd.nextInt(10) == 0 ? 0 : 1 + rnd.nextInt(n - 1);
                    badgeFontCombo.getSelectionModel().select(idx);
                }
            }
            if (badgeFontPctSlider != null) {
                badgeFontPctSlider.setValue(52 + rnd.nextInt(95));
            }
            if (badgeStrokeSlider != null) {
                double w = Math.round(rnd.nextDouble(0, 3.6) * 2.0) / 2.0;
                badgeStrokeSlider.setValue(Math.min(6, w));
            }
            if (badgeCornerSlider != null) {
                badgeCornerSlider.setValue(rnd.nextInt(22));
            }
            if (badgePillCheck != null) {
                badgePillCheck.setSelected(rnd.nextBoolean());
            }
            if (badgeGlowRadiusPctSlider != null) {
                badgeGlowRadiusPctSlider.setValue(45 + rnd.nextInt(186));
            }
            if (badgeGlowSpreadPctSlider != null) {
                badgeGlowSpreadPctSlider.setValue(35 + rnd.nextInt(166));
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

    private static Color randomBadgeFillColor(ThreadLocalRandom rnd) {
        double h = rnd.nextDouble(360);
        double s = 0.38 + rnd.nextDouble() * 0.52;
        double b = 0.38 + rnd.nextDouble() * 0.42;
        return Color.hsb(h, s, b);
    }

    private static Color randomBadgeTextColor(ThreadLocalRandom rnd, Color fill) {
        if (rnd.nextInt(6) == 0) {
            return Color.hsb(rnd.nextDouble(360), 0.12 + rnd.nextDouble() * 0.2, 0.88 + rnd.nextDouble() * 0.1);
        }
        if (fill == null) {
            return Color.WHITE;
        }
        double lum = 0.299 * fill.getRed() + 0.587 * fill.getGreen() + 0.114 * fill.getBlue();
        return lum > 0.52
                ? Color.color(0.08, 0.09, 0.12)
                : Color.color(0.94, 0.95, 0.97);
    }

    private static Color randomBadgeStrokeColor(ThreadLocalRandom rnd, Color fill) {
        if (fill == null) {
            return Color.rgb(rnd.nextInt(70), rnd.nextInt(70), rnd.nextInt(110));
        }
        return fill.darker().deriveColor(0, 1.2, 1.0 - rnd.nextDouble(0.18, 0.42), 1.0);
    }

    private static Color randomGlowColor(ThreadLocalRandom rnd, Color fill) {
        if (fill != null
                && rnd.nextBoolean()
                && !Double.isNaN(fill.getHue())
                && !Double.isNaN(fill.getSaturation())) {
            return Color.hsb(
                    fill.getHue(),
                    Math.min(1.0, 0.35 + rnd.nextDouble() * 0.45),
                    Math.min(1.0, 0.75 + rnd.nextDouble() * 0.22));
        }
        return Color.hsb(rnd.nextDouble(360), 0.4 + rnd.nextDouble() * 0.45, 0.82 + rnd.nextDouble() * 0.15);
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
        BadgeDesignTableItem sel =
                badgeMemberTable != null ? badgeMemberTable.getSelectionModel().getSelectedItem() : null;
        String txt = previewBadgeText(sel);
        PersonBadgeStyle st = buildStyleFromUiFields();
        badgePreviewBox
                .getChildren()
                .add(PersonBadgeNodeFactory.createBadge(txt, st, 1.0, 13.0));
    }

    /** プレビューに載せる文字（選択行のメンバー名由来）。 */
    private static String previewBadgeText(BadgeDesignTableItem sel) {
        if (sel == null) {
            return "?";
        }
        if (sel.globalFallback) {
            return "既定";
        }
        String b = PersonNameBadgeText.badgeTwoFromRawName(sel.memberDisplay);
        if (!b.isEmpty()) {
            return b;
        }
        String t = PersonNameBadgeText.firstNCodePoints(sel.memberDisplay.strip(), 2);
        return t.isEmpty() ? "?" : t;
    }
}
