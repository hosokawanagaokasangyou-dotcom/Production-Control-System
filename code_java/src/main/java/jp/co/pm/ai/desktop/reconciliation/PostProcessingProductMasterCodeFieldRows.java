package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;

import javafx.geometry.Pos;
import javafx.collections.FXCollections;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;

/**
 * 後加工商品マスタ編集フォームのマスタ連携列（コンボ＋名称表示）。
 */
final class PostProcessingProductMasterCodeFieldRows {

    /** 工程情報タブ等：ラベル列・入力列の幅を揃える。 */
    private static final double LABEL_COL_PREF = 148;
    private static final double CODE_COMBO_PREF = 88;
    private static final double CODE_COMBO_MAX = 104;
    private static final double TEXT_FIELD_PREF = 108;
    private static final double TEXT_FIELD_MAX = 132;
    private static final double NAME_HINT_PREF = 196;
    private static final double NAME_HINT_MAX = 268;
    private static final double FIELD_ROW_MAX = 420;

    private PostProcessingProductMasterCodeFieldRows() {}

    static double labelColumnPrefWidth() {
        return LABEL_COL_PREF;
    }

    static double fieldColumnPrefWidth() {
        return FIELD_ROW_MAX;
    }

    static double fieldColumnMaxWidth() {
        return FIELD_ROW_MAX + 24;
    }

    static void addColumnRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            PostProcessingKouteiNaiyoMasterLookup.Snapshot kouteiNaiyo,
            PostProcessingShuruiMasterLookup.Snapshot shurui,
            PostProcessingKeiriBunruiMasterLookup.Snapshot keiriBunrui,
            PostProcessingYotoMasterLookup.Snapshot yoto,
            PostProcessingShohinBunrui4MasterLookup.Snapshot bunrui4,
            PostProcessingZaikoBunruiMasterLookup.Snapshot zaikoBunrui,
            PostProcessingPlanMachineLookup.Snapshot planMachine,
            Map<String, TextField> fieldByColumn,
            Map<String, Label> nameLabelByColumn,
            Map<String, ComboBox<String>> comboByColumn,
            Consumer<String> onModelColumnChanged,
            Runnable shohinCodeDuplicateCheck,
            AtomicReference<Button> shohinDuplicateCheckButtonRef) {
        Label lbl = new Label(colName + ":");
        lbl.setStyle("-fx-font-size: 11px;");
        grid.add(lbl, 0, rowIndex);

        if ("商品コード".equals(colName)) {
            addShohinCodeRow(
                    grid,
                    rowIndex,
                    colName,
                    model,
                    fieldByColumn,
                    onModelColumnChanged,
                    shohinCodeDuplicateCheck,
                    shohinDuplicateCheckButtonRef);
            return;
        }

        if (planMachine != null
                && planMachine.loaded()
                && PostProcessingPlanMachineLookup.isMachineCodeColumn(colName)) {
            addMachineRow(
                    grid,
                    rowIndex,
                    colName,
                    model,
                    planMachine,
                    fieldByColumn,
                    nameLabelByColumn,
                    comboByColumn,
                    onModelColumnChanged);
            return;
        }
        if (kouteiNaiyo != null
                && kouteiNaiyo.loaded()
                && PostProcessingKouteiNaiyoMasterLookup.isKouteiCodeColumn(colName)) {
            addKouteiRow(
                    grid,
                    rowIndex,
                    colName,
                    model,
                    kouteiNaiyo,
                    fieldByColumn,
                    nameLabelByColumn,
                    comboByColumn,
                    onModelColumnChanged);
            return;
        }
        if (kouteiNaiyo != null
                && kouteiNaiyo.loaded()
                && PostProcessingKouteiNaiyoMasterLookup.isNaiyoCodeColumn(colName)) {
            addNaiyoRow(
                    grid,
                    rowIndex,
                    colName,
                    model,
                    kouteiNaiyo,
                    fieldByColumn,
                    nameLabelByColumn,
                    comboByColumn,
                    onModelColumnChanged);
            return;
        }
        if (shurui != null
                && shurui.loaded()
                && PostProcessingShuruiMasterLookup.isShuruiProductColumn(colName)) {
            addShuruiRow(
                    grid,
                    rowIndex,
                    colName,
                    model,
                    shurui,
                    fieldByColumn,
                    nameLabelByColumn,
                    comboByColumn,
                    onModelColumnChanged);
            return;
        }
        if (keiriBunrui != null
                && keiriBunrui.loaded()
                && PostProcessingKeiriBunruiMasterLookup.isKeiriBunruiProductColumn(colName)) {
            addKeiriBunruiRow(
                    grid,
                    rowIndex,
                    colName,
                    model,
                    keiriBunrui,
                    fieldByColumn,
                    nameLabelByColumn,
                    comboByColumn,
                    onModelColumnChanged);
            return;
        }
        if (yoto != null
                && yoto.loaded()
                && PostProcessingYotoMasterLookup.isYotoProductColumn(colName)) {
            addYotoRow(
                    grid,
                    rowIndex,
                    colName,
                    model,
                    yoto,
                    fieldByColumn,
                    nameLabelByColumn,
                    comboByColumn,
                    onModelColumnChanged);
            return;
        }
        if (bunrui4 != null
                && bunrui4.loaded()
                && PostProcessingShohinBunrui4MasterLookup.isBunrui4ProductColumn(colName)) {
            addBunrui4Row(
                    grid,
                    rowIndex,
                    colName,
                    model,
                    bunrui4,
                    fieldByColumn,
                    nameLabelByColumn,
                    comboByColumn,
                    onModelColumnChanged);
            return;
        }
        if (zaikoBunrui != null
                && zaikoBunrui.loaded()
                && PostProcessingZaikoBunruiMasterLookup.isZaikoBunruiProductColumn(colName)) {
            addZaikoBunruiRow(
                    grid,
                    rowIndex,
                    colName,
                    model,
                    zaikoBunrui,
                    fieldByColumn,
                    nameLabelByColumn,
                    comboByColumn,
                    onModelColumnChanged);
            return;
        }
        if (PostProcessingProductMasterKubunChoices.hasChoices(colName)) {
            addKubunRow(
                    grid,
                    rowIndex,
                    colName,
                    model,
                    fieldByColumn,
                    nameLabelByColumn,
                    comboByColumn,
                    onModelColumnChanged);
            return;
        }
        addPlainTextRow(
                grid, rowIndex, colName, model, fieldByColumn, onModelColumnChanged);
    }

    private static void addShohinCodeRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            Map<String, TextField> fieldByColumn,
            Consumer<String> onModelColumnChanged,
            Runnable onDuplicateCheck,
            AtomicReference<Button> duplicateCheckButtonRef) {
        TextField tf = new TextField(model.get(colName));
        tf.setStyle("-fx-font-size: 11px;");
        applyFormTextFieldSize(tf);
        tf.textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            onModelColumnChanged.accept(colName);
                        });
        fieldByColumn.put(colName, tf);

        Button btn = new Button("重複チェック");
        btn.getStyleClass().add("btn-reload");
        btn.setTooltip(
                new Tooltip("参照マスタとアップロード用一覧で商品コードの重複を確認します。"));
        if (onDuplicateCheck != null) {
            btn.setOnAction(e -> onDuplicateCheck.run());
        }
        if (duplicateCheckButtonRef != null) {
            duplicateCheckButtonRef.set(btn);
        }

        HBox fieldRow = new HBox(6, tf, btn);
        fieldRow.setAlignment(Pos.CENTER_LEFT);
        fieldRow.setMaxWidth(FIELD_ROW_MAX);
        GridPane.setHgrow(fieldRow, Priority.NEVER);
        grid.add(fieldRow, 1, rowIndex);
    }

    private static void addPlainTextRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            Map<String, TextField> fieldByColumn,
            Consumer<String> onModelColumnChanged) {
        TextField tf = new TextField(model.get(colName));
        tf.setStyle("-fx-font-size: 11px;");
        applyFormTextFieldSize(tf);
        tf.textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            onModelColumnChanged.accept(colName);
                        });
        fieldByColumn.put(colName, tf);
        GridPane.setHgrow(tf, Priority.NEVER);
        grid.add(tf, 1, rowIndex);
    }

    private static void addKouteiRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            PostProcessingKouteiNaiyoMasterLookup.Snapshot lookup,
            Map<String, TextField> fieldByColumn,
            Map<String, Label> nameLabelByColumn,
            Map<String, ComboBox<String>> comboByColumn,
            Consumer<String> onModelColumnChanged) {
        TextField codeField = new TextField(model.get(colName));
        codeField.setVisible(false);
        codeField.setManaged(false);
        codeField.setMaxWidth(0);
        codeField.setMaxHeight(0);

        Label nameLabel = new Label(formatKouteiHint(lookup, codeField.getText()));
        nameLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: #8ab4d8;");
        nameLabel.setWrapText(true);
        nameLabelByColumn.put(colName, nameLabel);

        ComboBox<String> combo = new ComboBox<>();
        combo.setEditable(true);
        combo.getStyleClass().add("combo-box");
        combo.getEditor().setStyle("-fx-font-size: 11px;");
        configurePickerCombo(combo, lookup.kouteiComboLabels());
        combo.setTooltip(
                new Tooltip(
                        "一覧はコード＋工程名。選択後はコードのみ（工程名は右ラベル）"));
        comboByColumn.put(colName, combo);

        AtomicBoolean suppress = new AtomicBoolean(false);
        Runnable syncDisplayFromCode =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    suppress.set(true);
                    try {
                        String code = codeField.getText();
                        String name =
                                PostProcessingKouteiNaiyoMasterLookup.resolveKouteiName(lookup, code);
                        nameLabel.setText(formatKouteiHint(lookup, code, name));
                        syncCodeOnlyComboEditor(
                                combo,
                                PostProcessingKouteiNaiyoMasterLookup.normalizeKouteiCode(
                                        code != null ? code : ""));
                    } finally {
                        suppress.set(false);
                    }
                };

        codeField
                .textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            syncDisplayFromCode.run();
                            onModelColumnChanged.accept(colName);
                        });

        Runnable commitCombo =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    String resolved =
                            PostProcessingKouteiNaiyoMasterLookup.resolveCodeFromComboInput(
                                    lookup, comboInputText(combo), true, 4);
                    suppress.set(true);
                    try {
                        codeField.setText(resolved);
                    } finally {
                        suppress.set(false);
                    }
                };

        wireCodeComboCommit(combo, commitCombo, suppress);

        syncDisplayFromCode.run();

        HBox fieldBox = buildCodePickerRow(combo, nameLabel);
        fieldByColumn.put(colName, codeField);
        grid.add(fieldBox, 1, rowIndex);
    }

    private static void addShuruiRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            PostProcessingShuruiMasterLookup.Snapshot lookup,
            Map<String, TextField> fieldByColumn,
            Map<String, Label> nameLabelByColumn,
            Map<String, ComboBox<String>> comboByColumn,
            Consumer<String> onModelColumnChanged) {
        TextField codeField = new TextField(model.get(colName));
        codeField.setVisible(false);
        codeField.setManaged(false);
        codeField.setMaxWidth(0);
        codeField.setMaxHeight(0);

        Label nameLabel = new Label(formatShuruiHint(lookup, codeField.getText()));
        nameLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: #8ab4d8;");
        nameLabel.setWrapText(true);
        nameLabelByColumn.put(colName, nameLabel);

        ComboBox<String> combo = new ComboBox<>();
        combo.setEditable(true);
        combo.getStyleClass().add("combo-box");
        combo.getEditor().setStyle("-fx-font-size: 11px;");
        configurePickerCombo(combo, lookup.comboLabels());
        combo.setTooltip(
                new Tooltip("種類マスタ（一覧はコード＋種類名。確定後はコードのみ）"));
        comboByColumn.put(colName, combo);

        AtomicBoolean suppress = new AtomicBoolean(false);
        Runnable syncDisplayFromCode =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    suppress.set(true);
                    try {
                        String code = codeField.getText();
                        String name =
                                PostProcessingShuruiMasterLookup.resolveName(lookup, code);
                        nameLabel.setText(formatShuruiHint(lookup, code, name));
                        syncCodeOnlyComboEditor(
                                combo,
                                PostProcessingShuruiMasterLookup.normalizeCode(
                                        code != null ? code : ""));
                    } finally {
                        suppress.set(false);
                    }
                };

        codeField
                .textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            syncDisplayFromCode.run();
                            onModelColumnChanged.accept(colName);
                        });

        Runnable commitCombo =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    String resolved =
                            PostProcessingShuruiMasterLookup.resolveCodeFromComboInput(
                                    lookup, comboInputText(combo));
                    suppress.set(true);
                    try {
                        codeField.setText(resolved);
                    } finally {
                        suppress.set(false);
                    }
                };

        wireCodeComboCommit(combo, commitCombo, suppress);

        syncDisplayFromCode.run();

        HBox fieldBox = buildCodePickerRow(combo, nameLabel);
        fieldByColumn.put(colName, codeField);
        grid.add(fieldBox, 1, rowIndex);
    }

    private static void addKeiriBunruiRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            PostProcessingKeiriBunruiMasterLookup.Snapshot lookup,
            Map<String, TextField> fieldByColumn,
            Map<String, Label> nameLabelByColumn,
            Map<String, ComboBox<String>> comboByColumn,
            Consumer<String> onModelColumnChanged) {
        TextField codeField = new TextField(model.get(colName));
        codeField.setVisible(false);
        codeField.setManaged(false);
        codeField.setMaxWidth(0);
        codeField.setMaxHeight(0);

        Label nameLabel = new Label(formatKeiriBunruiHint(lookup, codeField.getText()));
        nameLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: #8ab4d8;");
        nameLabel.setWrapText(true);
        nameLabelByColumn.put(colName, nameLabel);

        ComboBox<String> combo = new ComboBox<>();
        combo.setEditable(true);
        combo.getStyleClass().add("combo-box");
        combo.getEditor().setStyle("-fx-font-size: 11px;");
        configurePickerCombo(combo, lookup.comboLabels());
        combo.setTooltip(
                new Tooltip("経理分類マスタ（一覧はコード＋名称。確定後はコードのみ）"));
        comboByColumn.put(colName, combo);

        AtomicBoolean suppress = new AtomicBoolean(false);
        Runnable syncDisplayFromCode =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    suppress.set(true);
                    try {
                        String code = codeField.getText();
                        String name =
                                PostProcessingKeiriBunruiMasterLookup.resolveName(lookup, code);
                        nameLabel.setText(formatKeiriBunruiHint(lookup, code, name));
                        syncCodeOnlyComboEditor(
                                combo,
                                PostProcessingKeiriBunruiMasterLookup.normalizeCode(
                                        code != null ? code : ""));
                    } finally {
                        suppress.set(false);
                    }
                };

        codeField
                .textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            syncDisplayFromCode.run();
                            onModelColumnChanged.accept(colName);
                        });

        Runnable commitCombo =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    String resolved =
                            PostProcessingKeiriBunruiMasterLookup.resolveCodeFromComboInput(
                                    lookup, comboInputText(combo));
                    suppress.set(true);
                    try {
                        codeField.setText(resolved);
                    } finally {
                        suppress.set(false);
                    }
                };

        wireCodeComboCommit(combo, commitCombo, suppress);

        syncDisplayFromCode.run();

        HBox fieldBox = buildCodePickerRow(combo, nameLabel);
        fieldByColumn.put(colName, codeField);
        grid.add(fieldBox, 1, rowIndex);
    }

    private static void addYotoRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            PostProcessingYotoMasterLookup.Snapshot lookup,
            Map<String, TextField> fieldByColumn,
            Map<String, Label> nameLabelByColumn,
            Map<String, ComboBox<String>> comboByColumn,
            Consumer<String> onModelColumnChanged) {
        TextField codeField = new TextField(model.get(colName));
        codeField.setVisible(false);
        codeField.setManaged(false);
        codeField.setMaxWidth(0);
        codeField.setMaxHeight(0);

        Label nameLabel = new Label(formatYotoHint(lookup, codeField.getText()));
        nameLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: #8ab4d8;");
        nameLabel.setWrapText(true);
        nameLabelByColumn.put(colName, nameLabel);

        ComboBox<String> combo = new ComboBox<>();
        combo.setEditable(true);
        combo.getStyleClass().add("combo-box");
        combo.getEditor().setStyle("-fx-font-size: 11px;");
        configurePickerCombo(combo, lookup.comboLabels());
        combo.setTooltip(
                new Tooltip("用途マスタ（一覧はコード＋用途名。確定後はコードのみ）"));
        comboByColumn.put(colName, combo);

        AtomicBoolean suppress = new AtomicBoolean(false);
        Runnable syncDisplayFromCode =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    suppress.set(true);
                    try {
                        String code = codeField.getText();
                        String name = PostProcessingYotoMasterLookup.resolveName(lookup, code);
                        nameLabel.setText(formatYotoHint(lookup, code, name));
                        syncCodeOnlyComboEditor(
                                combo,
                                PostProcessingYotoMasterLookup.normalizeCode(
                                        code != null ? code : ""));
                    } finally {
                        suppress.set(false);
                    }
                };

        codeField
                .textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            syncDisplayFromCode.run();
                            onModelColumnChanged.accept(colName);
                        });

        Runnable commitCombo =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    String resolved =
                            PostProcessingYotoMasterLookup.resolveCodeFromComboInput(
                                    lookup, comboInputText(combo));
                    suppress.set(true);
                    try {
                        codeField.setText(resolved);
                    } finally {
                        suppress.set(false);
                    }
                };

        wireCodeComboCommit(combo, commitCombo, suppress);

        syncDisplayFromCode.run();

        HBox fieldBox = buildCodePickerRow(combo, nameLabel);
        fieldByColumn.put(colName, codeField);
        grid.add(fieldBox, 1, rowIndex);
    }

    private static void addBunrui4Row(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            PostProcessingShohinBunrui4MasterLookup.Snapshot lookup,
            Map<String, TextField> fieldByColumn,
            Map<String, Label> nameLabelByColumn,
            Map<String, ComboBox<String>> comboByColumn,
            Consumer<String> onModelColumnChanged) {
        TextField codeField = new TextField(model.get(colName));
        codeField.setVisible(false);
        codeField.setManaged(false);
        codeField.setMaxWidth(0);
        codeField.setMaxHeight(0);

        Label nameLabel = new Label(formatBunrui4Hint(lookup, codeField.getText()));
        nameLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: #8ab4d8;");
        nameLabel.setWrapText(true);
        nameLabelByColumn.put(colName, nameLabel);

        ComboBox<String> combo = new ComboBox<>();
        combo.setEditable(true);
        combo.getStyleClass().add("combo-box");
        combo.getEditor().setStyle("-fx-font-size: 11px;");
        configurePickerCombo(combo, lookup.comboLabels());
        combo.setTooltip(
                new Tooltip("商品分類4マスタ（一覧はコード＋名称。確定後はコードのみ）"));
        comboByColumn.put(colName, combo);

        AtomicBoolean suppress = new AtomicBoolean(false);
        Runnable syncDisplayFromCode =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    suppress.set(true);
                    try {
                        String code = codeField.getText();
                        String name =
                                PostProcessingShohinBunrui4MasterLookup.resolveName(lookup, code);
                        nameLabel.setText(formatBunrui4Hint(lookup, code, name));
                        syncCodeOnlyComboEditor(
                                combo,
                                PostProcessingShohinBunrui4MasterLookup.normalizeCode(
                                        code != null ? code : ""));
                    } finally {
                        suppress.set(false);
                    }
                };

        codeField
                .textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            syncDisplayFromCode.run();
                            onModelColumnChanged.accept(colName);
                        });

        Runnable commitCombo =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    String resolved =
                            PostProcessingShohinBunrui4MasterLookup.resolveCodeFromComboInput(
                                    lookup, comboInputText(combo));
                    suppress.set(true);
                    try {
                        codeField.setText(resolved);
                    } finally {
                        suppress.set(false);
                    }
                };

        wireCodeComboCommit(combo, commitCombo, suppress);

        syncDisplayFromCode.run();

        HBox fieldBox = buildCodePickerRow(combo, nameLabel);
        fieldByColumn.put(colName, codeField);
        grid.add(fieldBox, 1, rowIndex);
    }

    private static void addZaikoBunruiRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            PostProcessingZaikoBunruiMasterLookup.Snapshot lookup,
            Map<String, TextField> fieldByColumn,
            Map<String, Label> nameLabelByColumn,
            Map<String, ComboBox<String>> comboByColumn,
            Consumer<String> onModelColumnChanged) {
        TextField codeField = new TextField(model.get(colName));
        codeField.setVisible(false);
        codeField.setManaged(false);
        codeField.setMaxWidth(0);
        codeField.setMaxHeight(0);

        Label nameLabel = new Label(formatZaikoBunruiHint(lookup, codeField.getText()));
        nameLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: #8ab4d8;");
        nameLabel.setWrapText(true);
        nameLabelByColumn.put(colName, nameLabel);

        ComboBox<String> combo = new ComboBox<>();
        combo.setEditable(true);
        combo.getStyleClass().add("combo-box");
        combo.getEditor().setStyle("-fx-font-size: 11px;");
        configurePickerCombo(combo, lookup.comboLabels());
        combo.setTooltip(
                new Tooltip(
                        "在庫分類マスタ（一覧は6桁コード＋名称。保存は1〜5のコード）"));
        comboByColumn.put(colName, combo);

        AtomicBoolean suppress = new AtomicBoolean(false);
        Runnable syncDisplayFromCode =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    suppress.set(true);
                    try {
                        String code = codeField.getText();
                        String name =
                                PostProcessingZaikoBunruiMasterLookup.resolveName(lookup, code);
                        nameLabel.setText(formatZaikoBunruiHint(lookup, code, name));
                        syncCodeOnlyComboEditor(
                                combo,
                                PostProcessingZaikoBunruiMasterLookup.toProductColumnValue(
                                        code != null ? code : ""));
                    } finally {
                        suppress.set(false);
                    }
                };

        codeField
                .textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            syncDisplayFromCode.run();
                            onModelColumnChanged.accept(colName);
                        });

        Runnable commitCombo =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    String resolved =
                            PostProcessingZaikoBunruiMasterLookup.resolveCodeFromComboInput(
                                    lookup, comboInputText(combo));
                    suppress.set(true);
                    try {
                        codeField.setText(resolved);
                    } finally {
                        suppress.set(false);
                    }
                };

        wireCodeComboCommit(combo, commitCombo, suppress);

        syncDisplayFromCode.run();

        HBox fieldBox = buildCodePickerRow(combo, nameLabel);
        fieldByColumn.put(colName, codeField);
        grid.add(fieldBox, 1, rowIndex);
    }

    private static void addKubunRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            Map<String, TextField> fieldByColumn,
            Map<String, Label> nameLabelByColumn,
            Map<String, ComboBox<String>> comboByColumn,
            Consumer<String> onModelColumnChanged) {
        TextField codeField = new TextField(model.get(colName));
        codeField.setVisible(false);
        codeField.setManaged(false);
        codeField.setMaxWidth(0);
        codeField.setMaxHeight(0);

        Label nameLabel = new Label(formatKubunHint(colName, codeField.getText()));
        nameLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: #8ab4d8;");
        nameLabelByColumn.put(colName, nameLabel);

        ComboBox<String> combo = new ComboBox<>();
        combo.setEditable(true);
        combo.getStyleClass().add("combo-box");
        combo.getEditor().setStyle("-fx-font-size: 11px;");
        configurePickerCombo(combo, PostProcessingProductMasterKubunChoices.pickerLabels(colName));
        combo.setTooltip(
                new Tooltip("区分コード（一覧はコード:名称。確定後はコードのみ）"));
        comboByColumn.put(colName, combo);

        AtomicBoolean suppress = new AtomicBoolean(false);
        Runnable syncDisplayFromCode =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    suppress.set(true);
                    try {
                        String code = codeField.getText();
                        nameLabel.setText(formatKubunHint(colName, code));
                        syncCodeOnlyComboEditor(
                                combo,
                                PostProcessingProductMasterKubunChoices.normalizeCode(
                                        colName, code != null ? code : ""));
                    } finally {
                        suppress.set(false);
                    }
                };

        codeField
                .textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            syncDisplayFromCode.run();
                            onModelColumnChanged.accept(colName);
                        });

        Runnable commitCombo =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    String resolved =
                            PostProcessingProductMasterKubunChoices.resolveCodeFromPickerInput(
                                    colName, comboInputText(combo));
                    suppress.set(true);
                    try {
                        codeField.setText(resolved);
                    } finally {
                        suppress.set(false);
                    }
                };

        wireCodeComboCommit(combo, commitCombo, suppress);

        syncDisplayFromCode.run();

        HBox fieldBox = buildCodePickerRow(combo, nameLabel);
        fieldByColumn.put(colName, codeField);
        grid.add(fieldBox, 1, rowIndex);
    }

    private static void addMachineRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            PostProcessingPlanMachineLookup.Snapshot lookup,
            Map<String, TextField> fieldByColumn,
            Map<String, Label> nameLabelByColumn,
            Map<String, ComboBox<String>> comboByColumn,
            Consumer<String> onModelColumnChanged) {
        TextField codeField = new TextField(model.get(colName));
        codeField.setVisible(false);
        codeField.setManaged(false);
        codeField.setMaxWidth(0);
        codeField.setMaxHeight(0);

        Label nameLabel = new Label(formatMachineHint(lookup, codeField.getText()));
        nameLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: #8ab4d8;");
        nameLabel.setWrapText(true);
        nameLabelByColumn.put(colName, nameLabel);

        ComboBox<String> combo = new ComboBox<>();
        combo.setEditable(true);
        combo.getStyleClass().add("combo-box");
        combo.getEditor().setStyle("-fx-font-size: 11px;");
        configurePickerCombo(combo, lookup.machineComboLabels());
        combo.setTooltip(
                new Tooltip(
                        "一覧はコード＋機械名。選択後はコードのみ（機械名は右ラベル）"));
        comboByColumn.put(colName, combo);

        AtomicBoolean suppress = new AtomicBoolean(false);
        Runnable syncDisplayFromCode =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    suppress.set(true);
                    try {
                        String code = codeField.getText();
                        String name =
                                PostProcessingPlanMachineLookup.resolveMachineName(lookup, code);
                        nameLabel.setText(formatMachineHint(lookup, code, name));
                        syncCodeOnlyComboEditor(
                                combo,
                                PostProcessingPlanMachineLookup.normalizeMachineCode(
                                        code != null ? code : ""));
                    } finally {
                        suppress.set(false);
                    }
                };

        codeField
                .textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            syncDisplayFromCode.run();
                            onModelColumnChanged.accept(colName);
                        });

        Runnable commitCombo =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    String resolved =
                            PostProcessingPlanMachineLookup.resolveCodeFromComboInput(
                                    lookup, comboInputText(combo));
                    suppress.set(true);
                    try {
                        codeField.setText(resolved);
                    } finally {
                        suppress.set(false);
                    }
                };

        wireCodeComboCommit(combo, commitCombo, suppress);

        syncDisplayFromCode.run();

        HBox fieldBox = buildCodePickerRow(combo, nameLabel);
        fieldByColumn.put(colName, codeField);
        grid.add(fieldBox, 1, rowIndex);
    }

    private static void addNaiyoRow(
            GridPane grid,
            int rowIndex,
            String colName,
            PostProcessingProductMasterEditorModel model,
            PostProcessingKouteiNaiyoMasterLookup.Snapshot lookup,
            Map<String, TextField> fieldByColumn,
            Map<String, Label> nameLabelByColumn,
            Map<String, ComboBox<String>> comboByColumn,
            Consumer<String> onModelColumnChanged) {
        TextField codeField = new TextField(model.get(colName));
        codeField.setVisible(false);
        codeField.setManaged(false);
        codeField.setMaxWidth(0);
        codeField.setMaxHeight(0);

        Label nameLabel = new Label(formatNaiyoHint(lookup, codeField.getText()));
        nameLabel.setStyle("-fx-font-size: 11px; -fx-text-fill: #8ab4d8;");
        nameLabel.setWrapText(true);
        nameLabelByColumn.put(colName, nameLabel);

        ComboBox<String> combo = new ComboBox<>();
        combo.setEditable(true);
        combo.getStyleClass().add("combo-box");
        combo.getEditor().setStyle("-fx-font-size: 11px;");
        configurePickerCombo(combo, lookup.naiyoComboLabels());
        combo.setTooltip(
                new Tooltip(
                        "一覧はコード＋加工内容名。選択後はコードのみ（内容・工程は右ラベル）"));
        comboByColumn.put(colName, combo);

        AtomicBoolean suppress = new AtomicBoolean(false);
        Runnable syncDisplayFromCode =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    suppress.set(true);
                    try {
                        String code = codeField.getText();
                        PostProcessingKouteiNaiyoMasterLookup.NaiyoEntry entry =
                                PostProcessingKouteiNaiyoMasterLookup.resolveNaiyo(lookup, code);
                        nameLabel.setText(formatNaiyoHint(lookup, entry, code));
                        syncCodeOnlyComboEditor(
                                combo,
                                PostProcessingKouteiNaiyoMasterLookup.normalizeNaiyoCode(
                                        code != null ? code : ""));
                    } finally {
                        suppress.set(false);
                    }
                };

        codeField
                .textProperty()
                .addListener(
                        (obs, o, n) -> {
                            model.set(colName, n != null ? n : "");
                            syncDisplayFromCode.run();
                            onModelColumnChanged.accept(colName);
                        });

        Runnable commitCombo =
                () -> {
                    if (suppress.get()) {
                        return;
                    }
                    String resolved =
                            PostProcessingKouteiNaiyoMasterLookup.resolveCodeFromComboInput(
                                    lookup, comboInputText(combo), false, 4);
                    suppress.set(true);
                    try {
                        codeField.setText(resolved);
                        linkKouteiFromNaiyo(
                                colName,
                                resolved,
                                lookup,
                                fieldByColumn,
                                comboByColumn);
                    } finally {
                        suppress.set(false);
                    }
                };

        wireCodeComboCommit(combo, commitCombo, suppress);

        syncDisplayFromCode.run();

        HBox fieldBox = buildCodePickerRow(combo, nameLabel);
        fieldByColumn.put(colName, codeField);
        grid.add(fieldBox, 1, rowIndex);
    }

    private static void linkKouteiFromNaiyo(
            String naiyoColumn,
            String naiyoCode,
            PostProcessingKouteiNaiyoMasterLookup.Snapshot lookup,
            Map<String, TextField> fieldByColumn,
            Map<String, ComboBox<String>> comboByColumn) {
        PostProcessingKouteiNaiyoMasterLookup.stepIndex(naiyoColumn)
                .ifPresent(
                        step -> {
                            PostProcessingKouteiNaiyoMasterLookup.NaiyoEntry entry =
                                    PostProcessingKouteiNaiyoMasterLookup.resolveNaiyo(
                                            lookup, naiyoCode);
                            if (entry == null || entry.kouteiCode().isEmpty()) {
                                return;
                            }
                            String kouteiCol =
                                    PostProcessingKouteiNaiyoMasterLookup.kouteiColumnForStep(
                                            step);
                            TextField kouteiField = fieldByColumn.get(kouteiCol);
                            if (kouteiField != null) {
                                kouteiField.setText(entry.kouteiCode());
                            }
                        });
    }

    private static String formatZaikoBunruiHint(
            PostProcessingZaikoBunruiMasterLookup.Snapshot lookup, String rawCode) {
        return formatZaikoBunruiHint(
                lookup,
                rawCode,
                PostProcessingZaikoBunruiMasterLookup.resolveName(lookup, rawCode));
    }

    private static String formatZaikoBunruiHint(
            PostProcessingZaikoBunruiMasterLookup.Snapshot lookup,
            String rawCode,
            String name) {
        String lookupCode =
                PostProcessingZaikoBunruiMasterLookup.normalizeLookupCode(
                        rawCode != null ? rawCode : "");
        if (lookupCode.isEmpty()) {
            return "（在庫分類名）";
        }
        if (name == null || name.isBlank()) {
            return "（在庫分類マスタに該当なし: " + lookupCode + "）";
        }
        return "在庫分類名: " + name;
    }

    private static String formatBunrui4Hint(
            PostProcessingShohinBunrui4MasterLookup.Snapshot lookup, String rawCode) {
        return formatBunrui4Hint(
                lookup,
                rawCode,
                PostProcessingShohinBunrui4MasterLookup.resolveName(lookup, rawCode));
    }

    private static String formatBunrui4Hint(
            PostProcessingShohinBunrui4MasterLookup.Snapshot lookup,
            String rawCode,
            String name) {
        String code =
                PostProcessingShohinBunrui4MasterLookup.normalizeCode(
                        rawCode != null ? rawCode : "");
        if (code.isEmpty()) {
            return "（商品分類4名）";
        }
        if (name == null || name.isBlank()) {
            return "（商品分類4マスタに該当なし: " + code + "）";
        }
        return "商品分類4名: " + name;
    }

    private static String formatYotoHint(
            PostProcessingYotoMasterLookup.Snapshot lookup, String rawCode) {
        return formatYotoHint(
                lookup, rawCode, PostProcessingYotoMasterLookup.resolveName(lookup, rawCode));
    }

    private static String formatYotoHint(
            PostProcessingYotoMasterLookup.Snapshot lookup, String rawCode, String name) {
        String code =
                PostProcessingYotoMasterLookup.normalizeCode(rawCode != null ? rawCode : "");
        if (code.isEmpty()) {
            return "（用途名）";
        }
        if (name == null || name.isBlank()) {
            return "（用途マスタに該当なし: " + code + "）";
        }
        return "用途名: " + name;
    }

    private static String formatKeiriBunruiHint(
            PostProcessingKeiriBunruiMasterLookup.Snapshot lookup, String rawCode) {
        return formatKeiriBunruiHint(
                lookup,
                rawCode,
                PostProcessingKeiriBunruiMasterLookup.resolveName(lookup, rawCode));
    }

    private static String formatKeiriBunruiHint(
            PostProcessingKeiriBunruiMasterLookup.Snapshot lookup, String rawCode, String name) {
        String code =
                PostProcessingKeiriBunruiMasterLookup.normalizeCode(rawCode != null ? rawCode : "");
        if (code.isEmpty()) {
            return "（経理分類名）";
        }
        if (name == null || name.isBlank()) {
            return "（経理分類マスタに該当なし: " + code + "）";
        }
        return "経理分類名: " + name;
    }

    private static String formatShuruiHint(
            PostProcessingShuruiMasterLookup.Snapshot lookup, String rawCode) {
        return formatShuruiHint(
                lookup,
                rawCode,
                PostProcessingShuruiMasterLookup.resolveName(lookup, rawCode));
    }

    private static String formatShuruiHint(
            PostProcessingShuruiMasterLookup.Snapshot lookup, String rawCode, String name) {
        String code =
                PostProcessingShuruiMasterLookup.normalizeCode(rawCode != null ? rawCode : "");
        if (code.isEmpty()) {
            return "（種類名）";
        }
        if (name == null || name.isBlank()) {
            return "（種類マスタに該当なし: " + code + "）";
        }
        return "種類名: " + name;
    }

    private static String formatKubunHint(String columnName, String rawCode) {
        String code =
                PostProcessingProductMasterKubunChoices.normalizeCode(
                        columnName, rawCode != null ? rawCode : "");
        if (code.isEmpty()) {
            return "（区分名称）";
        }
        String label = PostProcessingProductMasterKubunChoices.resolveLabel(columnName, code);
        if (label.isEmpty()) {
            return "（未定義: " + code + "）";
        }
        return label;
    }

    private static String formatMachineHint(
            PostProcessingPlanMachineLookup.Snapshot lookup, String rawCode) {
        return formatMachineHint(
                lookup,
                rawCode,
                PostProcessingPlanMachineLookup.resolveMachineName(lookup, rawCode));
    }

    private static String formatMachineHint(
            PostProcessingPlanMachineLookup.Snapshot lookup, String rawCode, String name) {
        String code =
                PostProcessingPlanMachineLookup.normalizeMachineCode(
                        rawCode != null ? rawCode : "");
        if (code.isEmpty()) {
            return "（機械名）";
        }
        if (name == null || name.isBlank()) {
            return "（加工計画に該当なし: " + code + "）";
        }
        return "機械名: " + name;
    }

    private static void applyFormTextFieldSize(TextField tf) {
        tf.setPrefWidth(TEXT_FIELD_PREF);
        tf.setMinWidth(64);
        tf.setMaxWidth(TEXT_FIELD_MAX);
    }

    private static void applyCodeComboSize(ComboBox<String> combo) {
        combo.setPrefWidth(CODE_COMBO_PREF);
        combo.setMinWidth(72);
        combo.setMaxWidth(CODE_COMBO_MAX);
    }

    private static void applyNameHintLabel(Label nameLabel) {
        nameLabel.setMinWidth(100);
        nameLabel.setPrefWidth(NAME_HINT_PREF);
        nameLabel.setMaxWidth(NAME_HINT_MAX);
        nameLabel.setWrapText(true);
    }

    private static HBox buildCodePickerRow(ComboBox<String> combo, Label nameLabel) {
        applyCodeComboSize(combo);
        applyNameHintLabel(nameLabel);
        HBox fieldBox = new HBox(8, combo, nameLabel);
        fieldBox.setAlignment(Pos.CENTER_LEFT);
        fieldBox.setMaxWidth(FIELD_ROW_MAX);
        GridPane.setHgrow(fieldBox, Priority.NEVER);
        return fieldBox;
    }

    /** ドロップダウンはコード＋名称、確定後の入力欄はコードのみ。 */
    private static void configurePickerCombo(ComboBox<String> combo, List<String> pickerLabels) {
        List<String> items =
                pickerLabels != null ? new ArrayList<>(pickerLabels) : new ArrayList<>();
        combo.setItems(FXCollections.observableArrayList(items));
        combo.setCellFactory(
                lv ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                setText(empty || item == null ? null : item);
                            }
                        });
        int rows = items.isEmpty() ? 8 : Math.min(14, Math.max(8, items.size()));
        combo.setVisibleRowCount(rows);
    }

    private static String comboInputText(ComboBox<String> combo) {
        String editor = combo.getEditor().getText();
        if (editor != null && !editor.isBlank()) {
            return editor.trim();
        }
        String value = combo.getValue();
        return value != null ? value.trim() : "";
    }

    private static void syncCodeOnlyComboEditor(ComboBox<String> combo, String normalizedCode) {
        if (normalizedCode == null || normalizedCode.isEmpty()) {
            combo.getEditor().clear();
            combo.setValue(null);
            return;
        }
        combo.getEditor().setText(normalizedCode);
        if (combo.getItems().contains(normalizedCode)) {
            combo.setValue(normalizedCode);
        } else {
            combo.setValue(null);
        }
    }

    private static void wireCodeComboCommit(
            ComboBox<String> combo, Runnable commitCombo, AtomicBoolean suppress) {
        combo.setOnAction(e -> commitCombo.run());
        combo.valueProperty()
                .addListener(
                        (obs, oldVal, newVal) -> {
                            if (suppress.get() || newVal == null || newVal.isBlank()) {
                                return;
                            }
                            commitCombo.run();
                        });
        combo.getEditor()
                .focusedProperty()
                .addListener(
                        (obs, was, focused) -> {
                            if (!focused) {
                                commitCombo.run();
                            }
                        });
    }

    private static String formatKouteiHint(
            PostProcessingKouteiNaiyoMasterLookup.Snapshot lookup, String rawCode) {
        return formatKouteiHint(
                lookup,
                rawCode,
                PostProcessingKouteiNaiyoMasterLookup.resolveKouteiName(lookup, rawCode));
    }

    private static String formatKouteiHint(
            PostProcessingKouteiNaiyoMasterLookup.Snapshot lookup, String rawCode, String name) {
        String code =
                PostProcessingKouteiNaiyoMasterLookup.normalizeKouteiCode(
                        rawCode != null ? rawCode : "");
        if (code.isEmpty()) {
            return "（工程名）";
        }
        if (name == null || name.isBlank()) {
            return "（マスタに該当なし: " + code + "）";
        }
        return "工程名: " + name;
    }

    private static String formatNaiyoHint(
            PostProcessingKouteiNaiyoMasterLookup.Snapshot lookup, String rawCode) {
        return formatNaiyoHint(
                lookup,
                PostProcessingKouteiNaiyoMasterLookup.resolveNaiyo(lookup, rawCode),
                rawCode);
    }

    private static String formatNaiyoHint(
            PostProcessingKouteiNaiyoMasterLookup.Snapshot lookup,
            PostProcessingKouteiNaiyoMasterLookup.NaiyoEntry entry,
            String rawCode) {
        String code =
                PostProcessingKouteiNaiyoMasterLookup.normalizeNaiyoCode(
                        rawCode != null ? rawCode : "");
        if (code.isEmpty()) {
            return "（加工内容名・工程）";
        }
        if (entry == null) {
            return "（マスタに該当なし: " + code + "）";
        }
        String kouteiPart =
                entry.kouteiName() != null && !entry.kouteiName().isBlank()
                        ? entry.kouteiCode() + " " + entry.kouteiName()
                        : entry.kouteiCode();
        return "内容: " + entry.naiyoName() + "　工程: " + kouteiPart;
    }
}
