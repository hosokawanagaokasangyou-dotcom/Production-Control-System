package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;

import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TableView;
import javafx.scene.control.ScrollPane;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.dispatch.OvertimeSimulationEditState;
import jp.co.pm.ai.desktop.dispatch.OvertimeSimulationOverridesWriter;

/**
 * 段階2.1 / 段階3.1 残業シミュレーション: 勤怠（チェック）と残業時間（分）を編集し、確定後に各段階を実行するウィザード。
 */
public final class OvertimeSimulationWizard {

    private OvertimeSimulationWizard() {}

    /** ウィザード確定後に起動するパイプライン段階。 */
    public enum Target {
        STAGE21("段階2.1"),
        STAGE31("段階3.1");

        private final String label;

        Target(String label) {
            this.label = label;
        }

        String label() {
            return label;
        }

        String windowTitle() {
            return label + " — 残業/休出シミュ";
        }

        String executeButtonText() {
            return label + "を実行";
        }

        String introConfirmLine() {
            return "確定後、変更内容を反映して" + label + "（残業/休出シミュ）を実行します。";
        }

        String noChangeSummarySuffix() {
            return "\n（変更なし — master 勤怠のまま" + label + "を実行します）";
        }
    }

    public record GridRow(String member) {}

    /**
     * プレビュー取得済みの状態でウィザードを表示する。
     *
     * @param onConfirm 確定時に overrides JSON パスを渡す（段階起動は呼び出し側）
     * @param onDismissWithoutRun キャンセルで閉じたとき（確定前のみ）。省略可。
     */
    public static void show(
            Stage owner,
            MainShellController shell,
            AttendanceOvertimePreview.Preview preview,
            Target target,
            Consumer<Path> onConfirm,
            Runnable onDismissWithoutRun) {
        Objects.requireNonNull(preview, "preview");
        Objects.requireNonNull(target, "target");
        Objects.requireNonNull(onConfirm, "onConfirm");
        AttendanceOvertimePreview.Preview windowed =
                AttendanceOvertimePreview.limitToDefaultOvertimeSimWindow(preview);
        OvertimeSimulationEditState state = new OvertimeSimulationEditState(windowed);
        showDialog(owner, shell, state, target, onConfirm, onDismissWithoutRun);
    }

    public static void show(
            Stage owner,
            MainShellController shell,
            AttendanceOvertimePreview.Preview preview,
            Target target,
            Consumer<Path> onConfirm) {
        show(owner, shell, preview, target, onConfirm, null);
    }

    private static void showDialog(
            Stage owner,
            MainShellController shell,
            OvertimeSimulationEditState state,
            Target target,
            Consumer<Path> onConfirm,
            Runnable onDismissWithoutRun) {
        Stage stage = new Stage();
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        // タイトルバーの×で閉じると onHidden が不整合になり段階ボタンがロックされたままになるため、枠なし＋キャンセルのみで閉じる。
        stage.initStyle(StageStyle.UNDECORATED);
        stage.setTitle(target.windowTitle());

        final boolean[] confirmed = {false};
        final boolean[] closeAllowed = {false};

        Label eyebrow = new Label(target.label());
        eyebrow.getStyleClass().add("overtime-sim-eyebrow");
        Label titleLabel = new Label("残業シミュレーション");
        titleLabel.getStyleClass().add("overtime-sim-title");
        VBox headerTitles = new VBox(2, eyebrow, titleLabel);

        Label stepLabel = new Label("ステップ 1 / 2");
        stepLabel.getStyleClass().add("overtime-sim-step-badge");

        Region headerSpacer = new Region();
        HBox.setHgrow(headerSpacer, Priority.ALWAYS);
        HBox header = new HBox(12, headerTitles, headerSpacer, stepLabel);
        header.setAlignment(Pos.CENTER_LEFT);
        header.getStyleClass().add("overtime-sim-header");

        Label intro =
                new Label(
                        "各行: 出勤にチェックを入れ、残業時間（分・15分刻み）を入力。"
                                + " 日付は本日〜"
                                + AttendanceOvertimePreview.OVERTIME_SIM_DATE_WINDOW_DAYS_AFTER_TODAY
                                + "日後まで。"
                                + " "
                                + target.introConfirmLine());
        intro.setWrapText(true);
        intro.setMaxWidth(Double.MAX_VALUE);
        intro.getStyleClass().add("overtime-sim-intro");

        Label gridHint = new Label("□＝出勤　右＝残業（分・15分刻み、▲▼または直接入力）");
        gridHint.getStyleClass().add("overtime-sim-grid-hint");

        Node gridPanel = buildGridPanel(state);
        VBox.setVgrow(gridPanel, Priority.ALWAYS);

        VBox step1 = new VBox(10, intro, gridHint, gridPanel);

        Label summaryLabel = new Label();
        summaryLabel.setWrapText(true);
        summaryLabel.setMaxWidth(Double.MAX_VALUE);
        summaryLabel.getStyleClass().add("overtime-sim-summary");
        VBox.setVgrow(summaryLabel, Priority.ALWAYS);
        VBox step2 = new VBox(12, summaryLabel);

        BorderPane center = new BorderPane();
        center.setPadding(new Insets(16, 18, 16, 18));

        Button backBtn = new Button("戻る");
        backBtn.getStyleClass().add("overtime-sim-secondary-button");
        backBtn.setDisable(true);
        Button nextBtn = new Button("次へ");
        nextBtn.getStyleClass().add("overtime-sim-primary-button");
        nextBtn.setDefaultButton(true);
        Button cancelBtn = new Button("キャンセル");
        cancelBtn.getStyleClass().add("overtime-sim-secondary-button");

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
                    summaryLabel.setText(state.buildSummaryText(target.noChangeSummarySuffix()));
                    center.setCenter(step2);
                    backBtn.setDisable(false);
                    nextBtn.setText(target.executeButtonText());
                };

        showStep1.run();

        nextBtn.setOnAction(
                ev -> {
                    if (center.getCenter() == step1) {
                        showStep2.run();
                    } else {
                        try {
                            Path overrides =
                                    shell.writeStage21OvertimeSimulationOverridesJson(
                                            OvertimeSimulationOverridesWriter.buildFromEditState(
                                                    state));
                            confirmed[0] = true;
                            closeAllowed[0] = true;
                            stage.close();
                            onConfirm.accept(overrides);
                        } catch (Exception ex) {
                            shell.showErrorDialog(
                                    target.label(),
                                    "シミュレーション JSON の書き込みに失敗しました。\n"
                                            + (ex.getMessage() != null
                                                    ? ex.getMessage()
                                                    : ex));
                        }
                    }
                });

        backBtn.setOnAction(ev -> showStep1.run());
        cancelBtn.setOnAction(
                ev -> {
                    closeAllowed[0] = true;
                    stage.close();
                });

        stage.setOnCloseRequest(
                ev -> {
                    if (!closeAllowed[0]) {
                        ev.consume();
                    }
                });

        Region footerSpacer = new Region();
        HBox.setHgrow(footerSpacer, Priority.ALWAYS);
        HBox footer = new HBox(8, cancelBtn, footerSpacer, backBtn, nextBtn);
        footer.setAlignment(Pos.CENTER_LEFT);
        footer.getStyleClass().add("overtime-sim-footer");

        BorderPane root = new BorderPane();
        root.getStyleClass().add("overtime-sim-dialog");
        root.setTop(header);
        root.setCenter(center);
        root.setBottom(footer);

        Scene scene = new Scene(root, 960, 560);
        if (shell != null) {
            shell.registerOvertimeWizardScene(scene);
        }
        stage.setOnHidden(
                ev -> {
                    if (!confirmed[0] && onDismissWithoutRun != null) {
                        onDismissWithoutRun.run();
                    }
                });
        stage.setScene(scene);
        stage.showAndWait();
    }

    private static Node buildGridPanel(OvertimeSimulationEditState state) {
        List<GridRow> rows = new ArrayList<>();
        for (String m : state.members()) {
            rows.add(new GridRow(m));
        }

        TableView<GridRow> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        table.getStyleClass().addAll("overtime-simulation-grid", "overtime-sim-simple-table");
        table.setPlaceholder(new Label("表示するメンバーがありません"));
        table.getColumns().add(buildNameColumn());
        for (TableColumn<GridRow, String> dc : buildDateColumns(state)) {
            table.getColumns().add(dc);
        }

        double rowHeight = 34.0;
        table.setFixedCellSize(rowHeight);
        double tableBodyHeight = Math.min(400, rows.size() * rowHeight + 36);
        table.setMinHeight(tableBodyHeight);
        table.setPrefHeight(tableBodyHeight);
        double tableMinWidth = 120.0 + state.dates().size() * 100.0;
        table.setMinWidth(tableMinWidth);
        table.setPrefWidth(tableMinWidth);

        ScrollPane scroll = new ScrollPane(table);
        scroll.getStyleClass().add("overtime-sim-table-scroll");
        scroll.setFitToHeight(true);
        scroll.setFitToWidth(false);
        scroll.setHbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);
        scroll.setVbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);
        scroll.setPannable(true);
        VBox.setVgrow(scroll, Priority.ALWAYS);
        return scroll;
    }

    private static TableColumn<GridRow, String> buildNameColumn() {
        TableColumn<GridRow, String> nameCol = new TableColumn<>("氏名");
        nameCol.getStyleClass().add("overtime-sim-name-column");
        nameCol.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().member()));
        nameCol.setMinWidth(96);
        nameCol.setPrefWidth(120);
        nameCol.setMaxWidth(160);
        nameCol.setResizable(true);
        nameCol.setSortable(false);
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
                                setAlignment(Pos.CENTER_LEFT);
                            }
                        });
        return nameCol;
    }

    private static List<TableColumn<GridRow, String>> buildDateColumns(
            OvertimeSimulationEditState state) {
        List<TableColumn<GridRow, String>> cols = new ArrayList<>();
        for (LocalDate d : state.dates()) {
            cols.add(buildDateColumn(state, d));
        }
        return cols;
    }

    private static String formatCompactDateHeader(LocalDate d) {
        String dow =
                switch (d.getDayOfWeek()) {
                    case SATURDAY -> "土";
                    case SUNDAY -> "日";
                    case MONDAY -> "月";
                    case TUESDAY -> "火";
                    case WEDNESDAY -> "水";
                    case THURSDAY -> "木";
                    case FRIDAY -> "金";
                };
        return d.getMonthValue() + "/" + d.getDayOfMonth() + "(" + dow + ")";
    }

    private static TableColumn<GridRow, String> buildDateColumn(
            OvertimeSimulationEditState state, LocalDate date) {
        TableColumn<GridRow, String> dateCol = new TableColumn<>(formatCompactDateHeader(date));
        DayOfWeek dow = date.getDayOfWeek();
        if (dow == DayOfWeek.SATURDAY || dow == DayOfWeek.SUNDAY) {
            dateCol.getStyleClass().add("overtime-sim-weekend-column");
        }
        dateCol.setSortable(false);
        dateCol.setMinWidth(96);
        dateCol.setPrefWidth(100);
        dateCol.setCellFactory(
                col ->
                        new TableCell<>() {
                            private final CheckBox workingCheck = new CheckBox();
                            private final Spinner<Integer> minutesSpinner = new Spinner<>();
                            private final HBox cellBox = new HBox(5, workingCheck, minutesSpinner);
                            private boolean syncing;

                            {
                                cellBox.getStyleClass().add("overtime-sim-day-cell-box");
                                cellBox.setAlignment(Pos.CENTER_LEFT);
                                cellBox.setPadding(new Insets(2, 4, 2, 4));
                                cellBox.setFillHeight(false);
                                minutesSpinner.setEditable(true);
                                minutesSpinner.getStyleClass().add("overtime-sim-minutes-spinner");
                                minutesSpinner.setMinWidth(56);
                                minutesSpinner.setPrefWidth(58);
                                minutesSpinner.setMaxWidth(64);
                                HBox.setHgrow(minutesSpinner, Priority.NEVER);
                                HBox.setHgrow(workingCheck, Priority.NEVER);
                                resetMinutesSpinnerValueFactory(0);
                                workingCheck.getStyleClass().add("overtime-sim-working-check");
                                workingCheck
                                        .selectedProperty()
                                        .addListener(
                                                (obs, was, on) -> {
                                                    if (syncing) {
                                                        return;
                                                    }
                                                    GridRow row = getTableRow().getItem();
                                                    if (row == null) {
                                                        return;
                                                    }
                                                    String member = row.member();
                                                    if (member == null || member.isBlank()) {
                                                        return;
                                                    }
                                                    OvertimeSimulationEditState.CellState cs =
                                                            state.cell(date, member);
                                                    if (cs == null
                                                            || cs.currentWorking() == on) {
                                                        return;
                                                    }
                                                    state.toggleWorking(date, member);
                                                    syncFromState();
                                                });
                                minutesSpinner
                                        .valueProperty()
                                        .addListener(
                                                (obs, oldVal, newVal) -> {
                                                    if (syncing || newVal == null) {
                                                        return;
                                                    }
                                                    commitSpinnerMinutes(newVal);
                                                });
                                minutesSpinner
                                        .focusedProperty()
                                        .addListener(
                                                (obs, wasFocused, focused) -> {
                                                    if (focused || syncing) {
                                                        return;
                                                    }
                                                    Integer v = minutesSpinner.getValue();
                                                    if (v == null) {
                                                        return;
                                                    }
                                                    int snapped =
                                                            OvertimeSimulationEditState
                                                                    .snapOvertimeMinutes(v);
                                                    if (snapped != v) {
                                                        syncing = true;
                                                        try {
                                                            minutesSpinner.getValueFactory()
                                                                    .setValue(snapped);
                                                        } finally {
                                                            syncing = false;
                                                        }
                                                    }
                                                    commitSpinnerMinutes(snapped);
                                                });
                            }

                            private void resetMinutesSpinnerValueFactory(int initial) {
                                minutesSpinner.setValueFactory(
                                        new SpinnerValueFactory.IntegerSpinnerValueFactory(
                                                0,
                                                OvertimeSimulationEditState.OVERTIME_MINUTES_MAX,
                                                OvertimeSimulationEditState.snapOvertimeMinutes(
                                                        initial),
                                                OvertimeSimulationEditState
                                                        .OVERTIME_MINUTES_STEP));
                            }

                            private void commitSpinnerMinutes(int minutes) {
                                if (syncing) {
                                    return;
                                }
                                GridRow row = getTableRow().getItem();
                                if (row == null) {
                                    return;
                                }
                                String member = row.member();
                                if (member == null || member.isBlank()) {
                                    return;
                                }
                                OvertimeSimulationEditState.CellState cs =
                                        state.cell(date, member);
                                if (cs == null || !cs.currentWorking()) {
                                    return;
                                }
                                int snapped =
                                        OvertimeSimulationEditState.snapOvertimeMinutes(minutes);
                                if (snapped == cs.currentOvertimeMinutes()) {
                                    return;
                                }
                                state.setOvertimeMinutes(date, member, snapped);
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty) {
                                    setGraphic(null);
                                    getStyleClass().remove("overtime-sim-cell-off");
                                    return;
                                }
                                GridRow row = getTableRow().getItem();
                                if (row == null) {
                                    setGraphic(null);
                                    return;
                                }
                                String member = row.member();
                                if (member == null || member.isBlank()) {
                                    setGraphic(null);
                                    return;
                                }
                                setGraphic(cellBox);
                                setAlignment(Pos.CENTER);
                                setClip(null);
                                syncFromState();
                            }

                            private void syncFromState() {
                                GridRow row = getTableRow().getItem();
                                if (row == null) {
                                    return;
                                }
                                String member = row.member();
                                OvertimeSimulationEditState.CellState cs =
                                        state.cell(date, member);
                                if (cs == null) {
                                    return;
                                }
                                syncing = true;
                                try {
                                    workingCheck.setSelected(cs.currentWorking());
                                    if (cs.currentWorking()) {
                                        getStyleClass().remove("overtime-sim-cell-off");
                                        minutesSpinner.setDisable(false);
                                        int show =
                                                OvertimeSimulationEditState.snapOvertimeMinutes(
                                                        cs.currentOvertimeMinutes());
                                        Integer current = minutesSpinner.getValue();
                                        if (!minutesSpinner.isFocused()
                                                && (current == null || current != show)) {
                                            resetMinutesSpinnerValueFactory(show);
                                        }
                                    } else {
                                        if (!getStyleClass().contains("overtime-sim-cell-off")) {
                                            getStyleClass().add("overtime-sim-cell-off");
                                        }
                                        minutesSpinner.setDisable(true);
                                        resetMinutesSpinnerValueFactory(0);
                                    }
                                } finally {
                                    syncing = false;
                                }
                            }
                        });
        return dateCol;
    }
}
