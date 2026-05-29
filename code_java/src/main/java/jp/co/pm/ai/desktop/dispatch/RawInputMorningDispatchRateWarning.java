package jp.co.pm.ai.desktop.dispatch;

import java.nio.file.Path;
import java.time.format.DateTimeFormatter;
import java.util.Map;
import java.util.Objects;

import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.BorderPane;
import javafx.stage.Modality;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.MainShellController;
import jp.co.pm.ai.desktop.io.Stage2EquipmentGanttContractPaths;

/**
 * 段階2／段階3完了後、原反投入日制約で午前配台率が 50% 未満の暦日があれば警告ダイアログを出す。
 */
public final class RawInputMorningDispatchRateWarning {

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/M/d");

    private RawInputMorningDispatchRateWarning() {}

    public static void showIfNeeded(
            MainShellController shell,
            Stage owner,
            Path resultDispatchJson,
            Map<String, java.time.LocalDate> rawInputByTaskId) {
        if (shell == null || rawInputByTaskId == null || rawInputByTaskId.isEmpty()) {
            return;
        }
        Path contract = Stage2EquipmentGanttContractPaths.resolveNearResultDispatchJson(resultDispatchJson);
        if (contract == null) {
            return;
        }
        RawInputMorningDispatchRateAnalyzer.AnalysisResult result;
        try {
            result = RawInputMorningDispatchRateAnalyzer.analyze(contract, rawInputByTaskId);
        } catch (Exception ex) {
            return;
        }
        if (!result.hasWarnings()) {
            return;
        }
        showDialog(owner, shell, result);
    }

    private static void showDialog(
            Stage owner,
            MainShellController shell,
            RawInputMorningDispatchRateAnalyzer.AnalysisResult result) {
        TableView<DayRow> tv = new TableView<>();
        tv.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);

        TableColumn<DayRow, String> cDate = new TableColumn<>("暦日");
        cDate.setCellValueFactory(new PropertyValueFactory<>("dateText"));
        cDate.setPrefWidth(100);

        TableColumn<DayRow, String> cRate = new TableColumn<>("午前配台率");
        cRate.setCellValueFactory(new PropertyValueFactory<>("rateText"));
        cRate.setPrefWidth(90);

        TableColumn<DayRow, String> cUsed = new TableColumn<>("午前加工(分)");
        cUsed.setCellValueFactory(new PropertyValueFactory<>("usedText"));
        cUsed.setPrefWidth(100);

        TableColumn<DayRow, String> cCap = new TableColumn<>("午前枠(分)");
        cCap.setCellValueFactory(new PropertyValueFactory<>("capacityText"));
        cCap.setPrefWidth(100);

        TableColumn<DayRow, String> cTasks = new TableColumn<>("原反同日依頼（13:00以降開始）");
        cTasks.setCellValueFactory(new PropertyValueFactory<>("tasksText"));

        tv.getColumns().addAll(cDate, cRate, cUsed, cCap, cTasks);

        for (RawInputMorningDispatchRateAnalyzer.DayLowRate d : result.lowRateDays()) {
            String tasks =
                    String.join(
                            ", ",
                            d.rawInputSameDayTaskIds().stream().limit(8).toList());
            if (d.rawInputSameDayTaskIds().size() > 8) {
                tasks += " …他" + (d.rawInputSameDayTaskIds().size() - 8) + "件";
            }
            tv.getItems()
                    .add(
                            new DayRow(
                                    d.date().format(DATE_FMT),
                                    formatPercent(d.morningRate()),
                                    String.valueOf(d.morningUsedMinutes()),
                                    String.valueOf(d.morningCapacityMinutes()),
                                    tasks));
        }

        Label head =
                new Label(
                        "原反投入日と同日の加工は 13:00 以降にしか開始できないため、"
                                + "午前帯（08:45～13:00）の設備稼働率が "
                                + (int) (RawInputMorningDispatchRateAnalyzer.RATE_THRESHOLD * 100)
                                + "% 未満の暦日があります。"
                                + " 原反投入日の前倒し（タスク入力タブ）を検討してください。");
        head.setWrapText(true);
        head.setStyle("-fx-font-size: 13px;");

        BorderPane root = new BorderPane();
        root.setTop(head);
        BorderPane.setMargin(head, new Insets(10, 14, 8, 14));
        root.setCenter(tv);
        BorderPane.setMargin(tv, new Insets(0, 14, 14, 14));

        Stage st = new Stage();
        if (owner != null) {
            st.initOwner(owner);
        }
        st.initModality(Modality.APPLICATION_MODAL);
        st.setTitle("原反投入日 — 午前配台率警告");
        Scene sc = new Scene(root, 920, 420);
        if (shell != null) {
            shell.registerThemeTrackedScene(sc);
        }
        st.setScene(sc);
        st.setOnHidden(
                ev -> {
                    if (shell != null) {
                        shell.unregisterThemeTrackedScene(sc);
                    }
                });
        st.showAndWait();
    }

    private static String formatPercent(double rate) {
        return String.format("%.1f%%", rate * 100.0);
    }

    public static final class DayRow {
        private final String dateText;
        private final String rateText;
        private final String usedText;
        private final String capacityText;
        private final String tasksText;

        DayRow(String dateText, String rateText, String usedText, String capacityText, String tasksText) {
            this.dateText = dateText;
            this.rateText = rateText;
            this.usedText = usedText;
            this.capacityText = capacityText;
            this.tasksText = tasksText;
        }

        public String getDateText() {
            return dateText;
        }

        public String getRateText() {
            return rateText;
        }

        public String getUsedText() {
            return usedText;
        }

        public String getCapacityText() {
            return capacityText;
        }

        public String getTasksText() {
            return tasksText;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) {
                return true;
            }
            if (!(o instanceof DayRow dayRow)) {
                return false;
            }
            return Objects.equals(dateText, dayRow.dateText)
                    && Objects.equals(rateText, dayRow.rateText)
                    && Objects.equals(usedText, dayRow.usedText)
                    && Objects.equals(capacityText, dayRow.capacityText)
                    && Objects.equals(tasksText, dayRow.tasksText);
        }

        @Override
        public int hashCode() {
            return Objects.hash(dateText, rateText, usedText, capacityText, tasksText);
        }
    }
}
