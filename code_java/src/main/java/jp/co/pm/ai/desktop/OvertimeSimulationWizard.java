package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;

import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.Tooltip;
import javafx.scene.input.MouseButton;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.dispatch.OvertimeSimulationEditState;
import jp.co.pm.ai.desktop.dispatch.OvertimeSimulationOverridesWriter;

/**
 * 段階3.5 残業シミュレーション: 勤怠表（○/グレー）と残業時間（分）を編集し、段階3実行へ進むウィザード。
 */
public final class OvertimeSimulationWizard {

    private OvertimeSimulationWizard() {}

    public enum RowKind {
        ATTENDANCE("勤怠"),
        OVERTIME("残業時間");

        private final String label;

        RowKind(String label) {
            this.label = label;
        }

        String label() {
            return label;
        }
    }

    public record GridRow(String member, RowKind kind) {}

    /**
     * プレビュー取得済みの状態でウィザードを表示する。
     *
     * @param onRunStage3 確定時に overrides JSON パスを渡す（段階3 起動は呼び出し側）
     */
    public static void show(
            Stage owner,
            MainShellController shell,
            AttendanceOvertimePreview.Preview preview,
            Consumer<Path> onRunStage3) {
        Objects.requireNonNull(preview, "preview");
        Objects.requireNonNull(onRunStage3, "onRunStage3");
        OvertimeSimulationEditState state = new OvertimeSimulationEditState(preview);
        showDialog(owner, shell, state, onRunStage3);
    }

    private static void showDialog(
            Stage owner,
            MainShellController shell,
            OvertimeSimulationEditState state,
            Consumer<Path> onRunStage3) {
        Stage stage = new Stage();
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle("段階3.5 — 残業シミュレーション");

        Label stepLabel = new Label("ステップ 1 / 2");
        stepLabel.setStyle("-fx-font-weight: bold;");

        Label intro =
                new Label(
                        "メンバー勤怠を確認し、残業時間（分）を入力してください。"
                                + " 勤怠行はダブルクリックで ○（出勤）とグレー（休み）を切り替えられます（休日出勤シミュレーション）。"
                                + " 確定後、変更内容を反映して段階3（配台試行）を実行します。");
        intro.setWrapText(true);

        Node gridPanel = buildGridPanel(state);
        VBox.setVgrow(gridPanel, Priority.ALWAYS);

        VBox step1 = new VBox(10, intro, gridPanel);
        step1.setPadding(new Insets(0, 0, 8, 0));

        Label summaryLabel = new Label();
        summaryLabel.setWrapText(true);
        VBox step2 = new VBox(10, summaryLabel);
        step2.setPadding(new Insets(0, 0, 8, 0));

        BorderPane center = new BorderPane();
        center.setPadding(new Insets(12));

        Button backBtn = new Button("戻る");
        backBtn.setDisable(true);
        Button nextBtn = new Button("次へ");
        Button cancelBtn = new Button("キャンセル");

        Runnable showStep1 =
                () -> {
                    stepLabel.setText("ステップ 1 / 2");
                    center.setCenter(step1);
                    backBtn.setDisable(true);
                    nextBtn.setText("次へ");
                };

        Runnable showStep2 =
                () -> {
                    stepLabel.setText("ステップ 2 / 2");
                    summaryLabel.setText(state.buildSummaryText());
                    center.setCenter(step2);
                    backBtn.setDisable(false);
                    nextBtn.setText("段階3を実行");
                };

        showStep1.run();

        nextBtn.setOnAction(
                ev -> {
                    if (center.getCenter() == step1) {
                        showStep2.run();
                    } else {
                        try {
                            Path overrides =
                                    shell.writeOvertimeSimulationOverridesJson(
                                            OvertimeSimulationOverridesWriter.buildFromEditState(
                                                    state));
                            stage.close();
                            onRunStage3.accept(overrides);
                        } catch (Exception ex) {
                            shell.showErrorDialog(
                                    "段階3.5",
                                    "シミュレーション JSON の書き込みに失敗しました。\n"
                                            + (ex.getMessage() != null
                                                    ? ex.getMessage()
                                                    : ex));
                        }
                    }
                });

        backBtn.setOnAction(ev -> showStep1.run());
        cancelBtn.setOnAction(ev -> stage.close());

        HBox buttons = new HBox(8, backBtn, nextBtn, cancelBtn);
        buttons.setAlignment(Pos.CENTER_RIGHT);
        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox topBar = new HBox(12, stepLabel, spacer);

        BorderPane root = new BorderPane();
        root.setTop(new VBox(8, topBar));
        root.setCenter(center);
        root.setBottom(buttons);
        BorderPane.setMargin(buttons, new Insets(8, 12, 12, 12));

        Scene scene = new Scene(root, 960, 560);
        if (shell != null) {
            shell.registerThemeTrackedScene(scene);
            stage.setOnHidden(ev -> shell.unregisterThemeTrackedScene(scene));
        }
        stage.setScene(scene);
        stage.showAndWait();
    }

    private static Node buildGridPanel(OvertimeSimulationEditState state) {
        List<GridRow> rows = new ArrayList<>();
        List<String> members = state.members();
        for (String m : members) {
            rows.add(new GridRow(m, RowKind.ATTENDANCE));
            rows.add(new GridRow(m, RowKind.OVERTIME));
        }

        ObservableList<GridRow> items = FXCollections.observableArrayList(rows);

        TableView<GridRow> fixedTable = new TableView<>(items);
        fixedTable.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        fixedTable.getStyleClass().add("overtime-simulation-grid");
        fixedTable.getStyleClass().add("overtime-simulation-grid-fixed");
        fixedTable.setPrefWidth(220);
        fixedTable.setMinWidth(220);
        fixedTable.setMaxWidth(220);
        fixedTable.getColumns().add(buildNameColumn());
        fixedTable.getColumns().add(buildKindColumn());

        TableView<GridRow> dateTable = new TableView<>(items);
        dateTable.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        dateTable.getStyleClass().add("overtime-simulation-grid");
        dateTable.getStyleClass().add("overtime-simulation-grid-dates");
        for (TableColumn<GridRow, String> dc : buildDateColumns(state, members)) {
            dateTable.getColumns().add(dc);
        }

        double rowHeight = 28.0;
        fixedTable.setFixedCellSize(rowHeight);
        dateTable.setFixedCellSize(rowHeight);
        fixedTable.setPrefHeight(Math.min(420, items.size() * rowHeight + 32));
        dateTable.prefHeightProperty().bind(fixedTable.heightProperty());

        ScrollPane dateScroll = new ScrollPane(dateTable);
        dateScroll.setFitToHeight(true);
        HBox.setHgrow(dateScroll, Priority.ALWAYS);

        HBox panel = new HBox(0, fixedTable, dateScroll);
        panel.getStyleClass().add("overtime-simulation-grid-panel");
        return panel;
    }

    private static TableColumn<GridRow, String> buildNameColumn() {
        TableColumn<GridRow, String> nameCol = new TableColumn<>("氏名");
        nameCol.setCellValueFactory(
                cd -> {
                    GridRow row = cd.getValue();
                    if (row == null) {
                        return new SimpleStringProperty("");
                    }
                    // 2行ブロックの先頭（勤怠）行だけ氏名を表示
                    if (row.kind() == RowKind.ATTENDANCE) {
                        return new SimpleStringProperty(row.member());
                    }
                    return new SimpleStringProperty("");
                });
        nameCol.setPrefWidth(130);
        nameCol.setMinWidth(100);
        nameCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null || item.isBlank()) {
                                    setText(null);
                                } else {
                                    setText(item);
                                }
                            }
                        });
        return nameCol;
    }

    private static TableColumn<GridRow, String> buildKindColumn() {
        TableColumn<GridRow, String> kindCol = new TableColumn<>("区分");
        kindCol.setCellValueFactory(
                cd -> new SimpleStringProperty(cd.getValue().kind().label()));
        kindCol.setPrefWidth(72);
        kindCol.setMinWidth(64);
        return kindCol;
    }

    private static List<TableColumn<GridRow, String>> buildDateColumns(
            OvertimeSimulationEditState state, List<String> members) {
        List<TableColumn<GridRow, String>> cols = new ArrayList<>();
        for (LocalDate d : state.dates()) {
            cols.add(buildDateColumn(state, members, d));
        }
        return cols;
    }

    private static TableColumn<GridRow, String> buildDateColumn(
            OvertimeSimulationEditState state, List<String> members, LocalDate date) {
        TableColumn<GridRow, String> dateCol =
                new TableColumn<>(AttendanceOvertimePreview.formatDateHeader(date));
        dateCol.setPrefWidth(88);
        dateCol.setMinWidth(72);
        dateCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final Label attendanceLabel = new Label();
                            private final TextField overtimeField = new TextField();

                            {
                                attendanceLabel.setMaxWidth(Double.MAX_VALUE);
                                attendanceLabel.setAlignment(Pos.CENTER);
                                attendanceLabel.setOnMouseClicked(
                                        ev -> {
                                            if (ev.getButton() != MouseButton.PRIMARY
                                                    || ev.getClickCount() != 2) {
                                                return;
                                            }
                                            GridRow row = getTableRow().getItem();
                                            if (row == null || row.kind() != RowKind.ATTENDANCE) {
                                                return;
                                            }
                                            String member = memberForRow(getIndex(), members);
                                            if (member == null || member.isBlank()) {
                                                return;
                                            }
                                            state.toggleWorking(date, member);
                                            refresh();
                                            getTableView().refresh();
                                            fixedTableRefreshSibling(getTableView());
                                        });
                                overtimeField.setPrefWidth(64);
                                overtimeField.setMaxWidth(Double.MAX_VALUE);
                                overtimeField
                                        .textProperty()
                                        .addListener(
                                                (obs, o, n) -> {
                                                    GridRow row =
                                                            getTableRow() != null
                                                                    ? getTableRow().getItem()
                                                                    : null;
                                                    if (row == null
                                                            || row.kind() != RowKind.OVERTIME) {
                                                        return;
                                                    }
                                                    String member =
                                                            memberForRow(getIndex(), members);
                                                    if (member == null || member.isBlank()) {
                                                        return;
                                                    }
                                                    String t = n != null ? n.trim() : "";
                                                    if (t.isEmpty()) {
                                                        return;
                                                    }
                                                    try {
                                                        state.setOvertimeMinutes(
                                                                date, member, Integer.parseInt(t));
                                                    } catch (NumberFormatException ignored) {
                                                        // 入力途中
                                                    }
                                                });
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    setGraphic(null);
                                    setText(null);
                                    getStyleClass().remove("overtime-sim-cell-off");
                                    return;
                                }
                                GridRow row = getTableRow().getItem();
                                if (row == null) {
                                    setGraphic(null);
                                    return;
                                }
                                String member = memberForRow(getIndex(), members);
                                if (member == null || member.isBlank()) {
                                    setGraphic(null);
                                    return;
                                }
                                OvertimeSimulationEditState.CellState cs =
                                        state.cell(date, member);
                                if (cs == null) {
                                    setGraphic(null);
                                    return;
                                }
                                if (row.kind() == RowKind.ATTENDANCE) {
                                    setGraphic(attendanceLabel);
                                    attendanceLabel.setText(cs.currentWorking() ? "○" : "");
                                    if (cs.currentWorking()) {
                                        getStyleClass().remove("overtime-sim-cell-off");
                                    } else if (!getStyleClass().contains("overtime-sim-cell-off")) {
                                        getStyleClass().add("overtime-sim-cell-off");
                                    }
                                    setTooltip(
                                            new Tooltip(
                                                    "ダブルクリック: "
                                                            + (cs.currentWorking()
                                                                    ? "○を解除してグレーアウト"
                                                                    : "グレーを解除して○（休日出勤）")));
                                } else {
                                    if (cs.currentWorking()) {
                                        getStyleClass().remove("overtime-sim-cell-off");
                                        overtimeField.setDisable(false);
                                        int show = cs.currentOvertimeMinutes();
                                        String want = show > 0 ? String.valueOf(show) : "";
                                        if (!overtimeField.isFocused()
                                                && !want.equals(overtimeField.getText().trim())) {
                                            overtimeField.setText(want);
                                        }
                                        setGraphic(overtimeField);
                                        setTooltip(new Tooltip("残業時間（分）1〜720"));
                                    } else {
                                        if (!getStyleClass().contains("overtime-sim-cell-off")) {
                                            getStyleClass().add("overtime-sim-cell-off");
                                        }
                                        overtimeField.setDisable(true);
                                        overtimeField.clear();
                                        setGraphic(null);
                                        setText("");
                                    }
                                }
                            }

                            private void refresh() {
                                GridRow row = getTableRow().getItem();
                                if (row == null || row.kind() != RowKind.ATTENDANCE) {
                                    return;
                                }
                                String member = memberForRow(getIndex(), members);
                                if (member == null) {
                                    return;
                                }
                                OvertimeSimulationEditState.CellState cs =
                                        state.cell(date, member);
                                if (cs == null) {
                                    return;
                                }
                                attendanceLabel.setText(cs.currentWorking() ? "○" : "");
                                if (cs.currentWorking()) {
                                    getStyleClass().remove("overtime-sim-cell-off");
                                } else if (!getStyleClass().contains("overtime-sim-cell-off")) {
                                    getStyleClass().add("overtime-sim-cell-off");
                                }
                            }
                        });
        return dateCol;
    }

    /** 日付側テーブル更新時に固定列側も再描画する。 */
    private static void fixedTableRefreshSibling(TableView<GridRow> dateTable) {
        Node p = dateTable.getParent();
        while (p != null) {
            if (p instanceof HBox hb) {
                for (Node c : hb.getChildren()) {
                    if (c instanceof TableView<?> ft && ft != dateTable) {
                        @SuppressWarnings("unchecked")
                        TableView<GridRow> fixed = (TableView<GridRow>) ft;
                        fixed.refresh();
                        return;
                    }
                }
            }
            p = p.getParent();
        }
    }

    /** 2行ブロックのうち行 index からメンバー名を得る。 */
    private static String memberForRow(int rowIndex, List<String> members) {
        if (rowIndex < 0 || members == null || members.isEmpty()) {
            return null;
        }
        int memberIndex = rowIndex / 2;
        if (memberIndex >= members.size()) {
            return null;
        }
        return members.get(memberIndex);
    }
}
