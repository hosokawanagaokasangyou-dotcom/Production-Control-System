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
import javafx.stage.StageStyle;

import jp.co.pm.ai.desktop.dispatch.AttendanceOvertimePreview;
import jp.co.pm.ai.desktop.dispatch.OvertimeSimulationEditState;
import jp.co.pm.ai.desktop.dispatch.OvertimeSimulationOverridesWriter;
import jp.co.pm.ai.desktop.ui.SpreadsheetTabularSupport;

/**
 * 段階2.1 / 段階3.1 残業シミュレーション: 勤怠表（○/グレー）と残業時間（分）を編集し、確定後に各段階を実行するウィザード。
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
                        "メンバー勤怠を確認し、残業時間（分）を入力してください。"
                                + " 日付列は本日から "
                                + AttendanceOvertimePreview.OVERTIME_SIM_DATE_WINDOW_DAYS_AFTER_TODAY
                                + " 日後まで（当日を含む）に限定しています。"
                                + " 勤怠行はダブルクリックで ○（出勤）とグレー（休み）を切り替えられます（休日出勤シミュレーション）。"
                                + " "
                                + target.introConfirmLine());
        intro.setWrapText(true);
        intro.setMaxWidth(Double.MAX_VALUE);
        intro.getStyleClass().add("overtime-sim-intro");

        Node gridPanel = buildGridPanel(state);
        VBox.setVgrow(gridPanel, Priority.ALWAYS);

        VBox step1 = new VBox(12, intro, gridPanel);

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

        Scene scene = new Scene(root, 1024, 600);
        if (shell != null) {
            shell.registerThemeTrackedScene(scene);
        }
        stage.setOnHidden(
                ev -> {
                    if (shell != null) {
                        shell.unregisterThemeTrackedScene(scene);
                    }
                    if (!confirmed[0] && onDismissWithoutRun != null) {
                        onDismissWithoutRun.run();
                    }
                });
        stage.setScene(scene);
        stage.showAndWait();
    }

    private static Node buildGridPanel(OvertimeSimulationEditState state) {
        List<String> members = state.members();
        List<GridRow> rows = new ArrayList<>();
        for (String m : members) {
            rows.add(new GridRow(m, RowKind.ATTENDANCE));
            rows.add(new GridRow(m, RowKind.OVERTIME));
        }

        TableView<GridRow> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setColumnResizePolicy(TableView.UNCONSTRAINED_RESIZE_POLICY);
        table.getStyleClass().add("overtime-simulation-grid");
        SpreadsheetTabularSupport.installPmAiReadableTableChrome(table);
        table.getColumns().add(buildNameColumn());
        table.getColumns().add(buildKindColumn());
        for (TableColumn<GridRow, String> dc : buildDateColumns(state)) {
            table.getColumns().add(dc);
        }

        double rowHeight = 28.0;
        table.setFixedCellSize(rowHeight);
        table.setPrefHeight(Math.min(420, rows.size() * rowHeight + 32));
        VBox.setVgrow(table, Priority.ALWAYS);
        return table;
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
            OvertimeSimulationEditState state) {
        List<TableColumn<GridRow, String>> cols = new ArrayList<>();
        for (LocalDate d : state.dates()) {
            cols.add(buildDateColumn(state, d));
        }
        return cols;
    }

    private static TableColumn<GridRow, String> buildDateColumn(
            OvertimeSimulationEditState state, LocalDate date) {
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
                                            String member = row.member();
                                            if (member == null || member.isBlank()) {
                                                return;
                                            }
                                            state.toggleWorking(date, member);
                                            refresh();
                                            getTableView().refresh();
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
                                                    String member = row.member();
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
                                    getStyleClass().remove("pm-ai-readable-date-cell");
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
                                OvertimeSimulationEditState.CellState cs =
                                        state.cell(date, member);
                                if (cs == null) {
                                    setGraphic(null);
                                    return;
                                }
                                if (row.kind() == RowKind.ATTENDANCE) {
                                    setGraphic(attendanceLabel);
                                    attendanceLabel.setText(cs.currentWorking() ? "○" : "");
                                    applyDateCellStyleClasses(cs.currentWorking());
                                    setTooltip(
                                            new Tooltip(
                                                    "ダブルクリック: "
                                                            + (cs.currentWorking()
                                                                    ? "○を解除してグレーアウト"
                                                                    : "グレーを解除して○（休日出勤）")));
                                } else {
                                    if (cs.currentWorking()) {
                                        applyDateCellStyleClasses(true);
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
                                        applyDateCellStyleClasses(false);
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
                                String member = row.member();
                                if (member == null) {
                                    return;
                                }
                                OvertimeSimulationEditState.CellState cs =
                                        state.cell(date, member);
                                if (cs == null) {
                                    return;
                                }
                                attendanceLabel.setText(cs.currentWorking() ? "○" : "");
                                applyDateCellStyleClasses(cs.currentWorking());
                            }

                            private void applyDateCellStyleClasses(boolean working) {
                                // 出勤セルは薄緑にせず白地のまま。休みセルのみグレーで区別する。
                                getStyleClass().remove("pm-ai-readable-date-cell");
                                if (working) {
                                    getStyleClass().remove("overtime-sim-cell-off");
                                } else if (!getStyleClass().contains("overtime-sim-cell-off")) {
                                    getStyleClass().add("overtime-sim-cell-off");
                                }
                            }
                        });
        return dateCol;
    }
}
