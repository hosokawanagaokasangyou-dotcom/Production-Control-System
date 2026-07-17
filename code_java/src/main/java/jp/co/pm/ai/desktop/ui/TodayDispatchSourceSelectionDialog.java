package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.nio.file.Path;
import java.util.function.Function;

import javafx.event.ActionEvent;
import javafx.beans.property.SimpleObjectProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.geometry.Rectangle2D;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Button;
import javafx.scene.control.Alert;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.ToggleGroup;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.scene.layout.HBox;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Screen;
import javafx.stage.Stage;
import javafx.stage.Window;
import javafx.util.StringConverter;

import jp.co.pm.ai.planning.stage2.source.NetworkSourceExtractionCatalog;
import jp.co.pm.ai.planning.stage2.source.NetworkSourceExtractionTimeSupport;
import jp.co.pm.ai.planning.stage2.source.Stage1SourcePairMatcher;

/** 「当日配台する」選択時、段階1直前に加工計画取得時刻と自動ペア日報を選ぶ。 */
public final class TodayDispatchSourceSelectionDialog {

    private static final double COL_SELECT_W = 56;
    private static final double COL_PLAN_TIME_W = 88;
    private static final double COL_DAILY_TIME_W = 110;
    private static final double COL_DELTA_W = 80;
    private static final double COL_MANUAL_BUTTON_W = 128;
    private static final double COL_PAD = 28;
    private static final double DIALOG_CHROME_W = 56;
    private static final double DIALOG_PREF_H = 480;
    private static final double DIALOG_MIN_W = 760;
    private static final double DIALOG_MIN_H = 360;

    private TodayDispatchSourceSelectionDialog() {}

    public static boolean requiresManualDailySelection(Stage1SourcePairMatcher.MatchedPair pair) {
        return pair != null && pair.plan() != null && pair.sameDayDailyCandidates().isEmpty();
    }

    public static boolean canConfirmSelection(Stage1SourcePairMatcher.MatchedPair pair) {
        return pair != null && pair.plan() != null && pair.dailyReport() != null;
    }

    public static Optional<Stage1SourcePairMatcher.MatchedPair> selectManualDailyReport(
            Stage1SourcePairMatcher.MatchedPair base, Path csv) {
        if (base == null || csv == null) {
            return Optional.empty();
        }
        return NetworkSourceExtractionCatalog.resolveDailyReportEntry(csv)
                .map(entry -> Stage1SourcePairMatcher.withDailyOverride(base, entry));
    }

    public static final class Row {
        private final Stage1SourcePairMatcher.MatchedPair initial;
        private NetworkSourceExtractionCatalog.SourceEntry selectedDaily;
        private final SimpleStringProperty planTime = new SimpleStringProperty();
        private final SimpleStringProperty planFile = new SimpleStringProperty();
        private final SimpleStringProperty dailyTime = new SimpleStringProperty();
        private final SimpleStringProperty delta = new SimpleStringProperty();
        private final RadioButton selectRadio = new RadioButton();

        Row(Stage1SourcePairMatcher.MatchedPair initial, ToggleGroup group) {
            this.initial = initial;
            this.selectedDaily = initial != null ? initial.dailyReport() : null;
            selectRadio.setToggleGroup(group);
            refreshTexts();
        }

        void refreshTexts() {
            if (initial == null || initial.plan() == null) {
                planTime.set("—");
                planFile.set("");
                dailyTime.set("—");
                delta.set("");
                return;
            }
            planTime.set(NetworkSourceExtractionTimeSupport.displayTime(initial.plan().extractionTime()));
            planFile.set(initial.plan().fileName());
            if (selectedDaily != null) {
                dailyTime.set(NetworkSourceExtractionTimeSupport.displayTime(selectedDaily.extractionTime()));
                delta.set(
                        NetworkSourceExtractionTimeSupport.deltaMinutes(
                                        initial.plan().extractionTime(), selectedDaily.extractionTime())
                                + "分");
            } else if (initial.sameDayDailyMissing()) {
                dailyTime.set("（同日候補なし）");
                delta.set("—");
            } else {
                dailyTime.set("—");
                delta.set("");
            }
        }

        Stage1SourcePairMatcher.MatchedPair toMatchedPair() {
            if (initial == null) {
                return null;
            }
            if (selectedDaily == null) {
                return initial;
            }
            return Stage1SourcePairMatcher.withDailyOverride(initial, selectedDaily);
        }

        boolean largeDeltaWarning() {
            if (initial == null || selectedDaily == null) {
                return false;
            }
            return NetworkSourceExtractionTimeSupport.isLargePairDelta(
                    NetworkSourceExtractionTimeSupport.deltaMinutes(
                            initial.plan().extractionTime(), selectedDaily.extractionTime()));
        }

        List<NetworkSourceExtractionCatalog.SourceEntry> dailyCandidates() {
            return initial != null ? initial.sameDayDailyCandidates() : List.of();
        }

        RadioButton selectRadio() {
            return selectRadio;
        }

        /** JavaFX {@link PropertyValueFactory} 用（{@code get} 接頭辞必須）。 */
        public String getPlanTime() {
            return planTime.get();
        }

        public String getPlanFile() {
            return planFile.get();
        }

        public String getDailyTime() {
            return dailyTime.get();
        }

        public String getDelta() {
            return delta.get();
        }

        SimpleStringProperty planTimeProperty() {
            return planTime;
        }

        SimpleStringProperty planFileProperty() {
            return planFile;
        }

        SimpleStringProperty dailyTimeProperty() {
            return dailyTime;
        }

        SimpleStringProperty deltaProperty() {
            return delta;
        }
    }

    public static Optional<Stage1SourcePairMatcher.MatchedPair> prompt(
            Window owner, List<Stage1SourcePairMatcher.MatchedPair> pairs) {
        if (pairs == null || pairs.isEmpty()) {
            return Optional.empty();
        }
        ToggleGroup group = new ToggleGroup();
        List<Row> rows = new ArrayList<>();
        for (Stage1SourcePairMatcher.MatchedPair p : pairs) {
            rows.add(new Row(p, group));
        }

        Dialog<Stage1SourcePairMatcher.MatchedPair> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle("当日配台 — ソース選択");
        dialog.setResizable(true);
        dialog.getDialogPane().getButtonTypes().addAll(ButtonType.OK, ButtonType.CANCEL);

        double planFileW =
                Math.max(
                        textWidth("加工計画") + COL_PAD,
                        maxTextWidth(rows, Row::getPlanFile) + COL_PAD);
        double dailyLabelW = Math.max(textWidth("加工日報") + COL_PAD, maxDailyLabelWidth(rows) + COL_PAD);
        boolean anyManual = rows.stream().anyMatch(row -> requiresManualDailySelection(row.initial));
        double dailyPickW = dailyLabelW + (anyManual ? COL_MANUAL_BUTTON_W + 8 : 0);
        double comboPrefW = dailyLabelW;

        TableView<Row> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        table.setPrefHeight(320);
        VBox.setVgrow(table, Priority.ALWAYS);

        TableColumn<Row, Row> selectCol = new TableColumn<>("選択");
        selectCol.setPrefWidth(COL_SELECT_W);
        selectCol.setMinWidth(COL_SELECT_W);
        selectCol.setMaxWidth(COL_SELECT_W);
        selectCol.setResizable(false);
        selectCol.setCellValueFactory(c -> new SimpleObjectProperty<>(c.getValue()));
        selectCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            @Override
                            protected void updateItem(Row item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setGraphic(null);
                                } else {
                                    setGraphic(item.selectRadio());
                                }
                            }
                        });

        TableColumn<Row, String> planTimeCol = new TableColumn<>("計画取得");
        planTimeCol.setPrefWidth(COL_PLAN_TIME_W);
        planTimeCol.setMinWidth(72);
        planTimeCol.setCellValueFactory(c -> c.getValue().planTimeProperty());

        TableColumn<Row, String> planFileCol = new TableColumn<>("加工計画");
        planFileCol.setPrefWidth(planFileW);
        planFileCol.setMinWidth(160);
        planFileCol.setCellValueFactory(c -> c.getValue().planFileProperty());

        TableColumn<Row, String> dailyTimeCol = new TableColumn<>("日報取得");
        dailyTimeCol.setPrefWidth(COL_DAILY_TIME_W);
        dailyTimeCol.setMinWidth(88);
        dailyTimeCol.setCellValueFactory(c -> c.getValue().dailyTimeProperty());

        TableColumn<Row, String> deltaCol = new TableColumn<>("差分");
        deltaCol.setPrefWidth(COL_DELTA_W);
        deltaCol.setMinWidth(64);
        deltaCol.setCellValueFactory(c -> c.getValue().deltaProperty());

        TableColumn<Row, Row> dailyPickCol = new TableColumn<>("加工日報");
        dailyPickCol.setPrefWidth(dailyPickW);
        dailyPickCol.setMinWidth(200);
        dailyPickCol.setCellValueFactory(c -> new SimpleObjectProperty<>(c.getValue()));
        dailyPickCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final ComboBox<NetworkSourceExtractionCatalog.SourceEntry> combo =
                                    new ComboBox<>();
                            private final Button manualButton = new Button("CSVを手動選択…");
                            private final HBox box = new HBox(6, combo, manualButton);

                            {
                                combo.setPrefWidth(comboPrefW);
                                combo.setMaxWidth(Double.MAX_VALUE);
                                HBox.setHgrow(combo, Priority.ALWAYS);
                                combo.setConverter(
                                        new StringConverter<>() {
                                            @Override
                                            public String toString(
                                                    NetworkSourceExtractionCatalog.SourceEntry e) {
                                                return formatDailyLabel(e);
                                            }

                                            @Override
                                            public NetworkSourceExtractionCatalog.SourceEntry fromString(
                                                    String s) {
                                                return null;
                                            }
                                        });
                                combo.valueProperty()
                                        .addListener(
                                                (o, a, b) -> {
                                                    Row row = getTableRow() != null
                                                            ? getTableRow().getItem()
                                                            : null;
                                                    if (row != null) {
                                                        row.selectedDaily = b;
                                                        row.refreshTexts();
                                                    }
                                                });
                                manualButton.setOnAction(
                                        event -> {
                                            Row row = getTableRow() != null ? getTableRow().getItem() : null;
                                            if (row == null) {
                                                return;
                                            }
                                            FileChooser chooser = new FileChooser();
                                            chooser.setTitle("加工日報CSVを選択");
                                            chooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("CSV", "*.csv"));
                                            java.io.File selected = chooser.showOpenDialog(owner);
                                            selectManualDailyReport(row.initial, selected != null ? selected.toPath() : null)
                                                    .ifPresent(pair -> {
                                                        row.selectedDaily = pair.dailyReport();
                                                        row.refreshTexts();
                                                        combo.setValue(pair.dailyReport());
                                                    });
                                        });
                            }

                            @Override
                            protected void updateItem(Row item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setGraphic(null);
                                } else {
                                    combo.setItems(
                                            FXCollections.observableArrayList(
                                                    item.dailyCandidates()));
                                    combo.setValue(item.selectedDaily);
                                    combo.setDisable(item.dailyCandidates().isEmpty());
                                    manualButton.setVisible(requiresManualDailySelection(item.initial));
                                    manualButton.setManaged(manualButton.isVisible());
                                    setGraphic(box);
                                }
                            }
                        });

        table.getColumns().addAll(selectCol, planTimeCol, planFileCol, dailyTimeCol, deltaCol, dailyPickCol);
        if (!rows.isEmpty()) {
            rows.getFirst().selectRadio().setSelected(true);
        }

        double tablePrefW =
                COL_SELECT_W
                        + COL_PLAN_TIME_W
                        + planFileW
                        + COL_DAILY_TIME_W
                        + COL_DELTA_W
                        + dailyPickW
                        + 24;
        table.setPrefWidth(tablePrefW);

        Label hint =
                new Label(
                        "加工計画の取得時刻（行）を1つ選んでください。日報は同日最接近で自動ペアします。"
                                + " 差が大きい場合は警告のみ表示し、実行は可能です。");
        hint.setWrapText(true);
        VBox root = new VBox(8, hint, table);
        root.setPadding(new Insets(12));
        root.setFillWidth(true);
        dialog.getDialogPane().setContent(root);

        Rectangle2D visual = Screen.getPrimary().getVisualBounds();
        double prefW = Math.min(visual.getWidth() * 0.92, tablePrefW + DIALOG_CHROME_W);
        double prefH = Math.min(visual.getHeight() * 0.85, DIALOG_PREF_H);
        dialog.getDialogPane().setPrefSize(prefW, prefH);
        dialog.getDialogPane().setMinWidth(DIALOG_MIN_W);
        dialog.getDialogPane().setMinHeight(DIALOG_MIN_H);
        dialog.setOnShown(
                ev -> {
                    Window w = dialog.getDialogPane().getScene() != null
                            ? dialog.getDialogPane().getScene().getWindow()
                            : null;
                    if (w instanceof Stage stage) {
                        stage.setMinWidth(DIALOG_MIN_W);
                        stage.setMinHeight(DIALOG_MIN_H);
                    }
                });

        dialog.getDialogPane().lookupButton(ButtonType.OK).addEventFilter(
                ActionEvent.ACTION,
                event -> {
                    Stage1SourcePairMatcher.MatchedPair selected = rows.stream()
                            .filter(row -> row.selectRadio().isSelected())
                            .findFirst().map(Row::toMatchedPair).orElse(null);
                    if (!canConfirmSelection(selected)) {
                        event.consume();
                        Alert alert = new Alert(Alert.AlertType.WARNING, "加工日報CSVを選択してください。", ButtonType.OK);
                        if (owner != null) {
                            alert.initOwner(owner);
                        }
                        alert.showAndWait();
                    }
                });

        if (anyManual) {
            Label missingWarning = new Label("同日の加工日報候補がない行があります。CSVを手動選択してください。");
            missingWarning.setStyle("-fx-text-fill: #b45309;");
            root.getChildren().add(1, missingWarning);
        }

        dialog.setResultConverter(
                button -> {
                    if (button != ButtonType.OK) {
                        return null;
                    }
                    for (Row row : rows) {
                        if (row.selectRadio().isSelected()) {
                            Stage1SourcePairMatcher.MatchedPair pair = row.toMatchedPair();
                            if (pair == null || pair.dailyReport() == null) {
                                return null;
                            }
                            return pair;
                        }
                    }
                    return null;
                });

        return dialog.showAndWait();
    }

    static String formatDailyLabel(NetworkSourceExtractionCatalog.SourceEntry e) {
        if (e == null) {
            return "";
        }
        return NetworkSourceExtractionTimeSupport.displayTime(e.extractionTime()) + " " + e.fileName();
    }

    private static double maxTextWidth(List<Row> rows, Function<Row, String> getter) {
        double max = 0;
        for (Row row : rows) {
            max = Math.max(max, textWidth(getter.apply(row)));
        }
        return max;
    }

    private static double maxDailyLabelWidth(List<Row> rows) {
        double max = 0;
        for (Row row : rows) {
            max = Math.max(max, textWidth(formatDailyLabel(row.selectedDaily)));
            for (NetworkSourceExtractionCatalog.SourceEntry e : row.dailyCandidates()) {
                max = Math.max(max, textWidth(formatDailyLabel(e)));
            }
        }
        return max;
    }

    private static double textWidth(String s) {
        Text text = new Text(s == null ? "" : s);
        return Math.ceil(text.getLayoutBounds().getWidth());
    }
}
