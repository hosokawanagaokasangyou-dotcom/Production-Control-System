package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javafx.beans.property.SimpleObjectProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.RadioButton;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.ToggleGroup;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;
import javafx.util.StringConverter;

import jp.co.pm.ai.planning.stage2.source.NetworkSourceExtractionCatalog;
import jp.co.pm.ai.planning.stage2.source.NetworkSourceExtractionTimeSupport;
import jp.co.pm.ai.planning.stage2.source.Stage1SourcePairMatcher;

/** 当日配台 ON 時、段階1直前に加工計画取得時刻と自動ペア日報を選ぶ。 */
public final class TodayDispatchSourceSelectionDialog {

    private TodayDispatchSourceSelectionDialog() {}

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

        String planTime() {
            return planTime.get();
        }

        String planFile() {
            return planFile.get();
        }

        String dailyTime() {
            return dailyTime.get();
        }

        String delta() {
            return delta.get();
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
        dialog.getDialogPane().getButtonTypes().addAll(ButtonType.OK, ButtonType.CANCEL);

        TableView<Row> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        table.setPrefHeight(320);

        TableColumn<Row, Row> selectCol = new TableColumn<>("選択");
        selectCol.setMaxWidth(56);
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
        planTimeCol.setCellValueFactory(new PropertyValueFactory<>("planTime"));

        TableColumn<Row, String> planFileCol = new TableColumn<>("加工計画");
        planFileCol.setCellValueFactory(new PropertyValueFactory<>("planFile"));

        TableColumn<Row, String> dailyTimeCol = new TableColumn<>("日報取得");
        dailyTimeCol.setCellValueFactory(new PropertyValueFactory<>("dailyTime"));

        TableColumn<Row, String> deltaCol = new TableColumn<>("差分");
        deltaCol.setMaxWidth(72);
        deltaCol.setCellValueFactory(new PropertyValueFactory<>("delta"));

        TableColumn<Row, Row> dailyPickCol = new TableColumn<>("加工日報");
        dailyPickCol.setCellValueFactory(c -> new SimpleObjectProperty<>(c.getValue()));
        dailyPickCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final ComboBox<NetworkSourceExtractionCatalog.SourceEntry> combo =
                                    new ComboBox<>();

                            {
                                combo.setConverter(
                                        new StringConverter<>() {
                                            @Override
                                            public String toString(
                                                    NetworkSourceExtractionCatalog.SourceEntry e) {
                                                if (e == null) {
                                                    return "";
                                                }
                                                return NetworkSourceExtractionTimeSupport.displayTime(
                                                                e.extractionTime())
                                                        + " "
                                                        + e.fileName();
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
                                    setGraphic(combo);
                                }
                            }
                        });

        table.getColumns().addAll(selectCol, planTimeCol, planFileCol, dailyTimeCol, deltaCol, dailyPickCol);
        if (!rows.isEmpty()) {
            rows.getFirst().selectRadio().setSelected(true);
        }

        Label hint =
                new Label(
                        "加工計画の取得時刻（行）を1つ選んでください。日報は同日最接近で自動ペアします。"
                                + " 差が大きい場合は警告のみ表示し、実行は可能です。");
        hint.setWrapText(true);
        VBox root = new VBox(8, hint, table);
        root.setPadding(new Insets(12));
        dialog.getDialogPane().setContent(root);

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
}
