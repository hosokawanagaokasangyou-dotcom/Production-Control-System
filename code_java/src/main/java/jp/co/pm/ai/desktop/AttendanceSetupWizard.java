package jp.co.pm.ai.desktop;

import java.time.LocalDate;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

import jp.co.pm.ai.desktop.ui.FiscalYearPeriod;

/**
 * 初回／再セットアップ: 祝日取得 → メンバー同期 → 状態表示。
 */
public final class AttendanceSetupWizard {

    private AttendanceSetupWizard() {}

    public static void show(MainShellController shell, Consumer<Boolean> onComplete) {
        if (shell == null) {
            return;
        }
        Stage stage = new Stage(StageStyle.DECORATED);
        stage.initOwner(shell.primaryStageForDialogs());
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle("勤怠カレンダー セットアップ");

        Label status = new Label("祝日取得とメンバー同期を順に実行します。");
        status.setWrapText(true);
        LocalDate today = LocalDate.now();
        int fiscalYear =
                FiscalYearPeriod.fiscalYearLabelFor(
                        today, FiscalYearPeriod.DEFAULT_APRIL_MARCH);
        Spinner<Integer> fiscalYearSpinner =
                new Spinner<>(
                        new SpinnerValueFactory.IntegerSpinnerValueFactory(
                                2020, 2040, fiscalYear));
        Spinner<Integer> monthSpinner =
                new Spinner<>(
                        new SpinnerValueFactory.IntegerSpinnerValueFactory(
                                1, 12, today.getMonthValue()));

        final boolean[] fetchOk = {false};
        final boolean[] syncOk = {false};

        Button fetchBtn = new Button("1. 祝日を取得（年度・週末含む）");
        Button syncBtn = new Button("2. メンバー勤怠を同期");
        Button closeBtn = new Button("閉じる");
        closeBtn.setDefaultButton(true);

        fetchBtn.setOnAction(
                e -> {
                    status.setText("祝日取得中…");
                    fetchBtn.setDisable(true);
                    shell.runAttendanceDataIoAsync(
                            shell.buildAttendanceDataIoRequest(
                                    "fetch_holidays_fiscal",
                                    Integer.toString(fiscalYearSpinner.getValue()),
                                    "4",
                                    "1",
                                    "--weekends"),
                            node -> {
                                fetchOk[0] = true;
                                status.setText(
                                        "祝日: 適用 "
                                                + node.path("applied").asInt(0)
                                                + " 日。次にメンバー同期を実行してください。");
                                fetchBtn.setDisable(false);
                                shell.refreshAttendanceReadiness();
                            },
                            err -> {
                                status.setText("祝日取得失敗: " + err);
                                fetchBtn.setDisable(false);
                            });
                });

        syncBtn.setOnAction(
                e -> {
                    status.setText("メンバー同期中…");
                    syncBtn.setDisable(true);
                    int year = fiscalYearSpinner.getValue();
                    int month = monthSpinner.getValue();
                    shell.runAttendanceDataIoAsync(
                            shell.buildAttendanceDataIoRequest(
                                    "sync_members",
                                    Integer.toString(year),
                                    Integer.toString(month)),
                            node -> {
                                syncOk[0] = true;
                                status.setText(
                                        "同期完了: 適用 "
                                                + node.path("applied").asInt(0)
                                                + " / スキップ "
                                                + node.path("skipped").asInt(0));
                                syncBtn.setDisable(false);
                                shell.refreshAttendanceReadiness();
                            },
                            err -> {
                                status.setText("同期失敗: " + err);
                                syncBtn.setDisable(false);
                            });
                });

        closeBtn.setOnAction(
                e -> {
                    stage.close();
                    if (onComplete != null) {
                        onComplete.accept(fetchOk[0] || syncOk[0]);
                    }
                });

        HBox yearRow =
                new HBox(
                        8,
                        new Label("年度"),
                        fiscalYearSpinner,
                        new Label("同期月"),
                        monthSpinner);
        yearRow.setAlignment(Pos.CENTER_LEFT);
        Region spacer = new Region();
        VBox.setVgrow(spacer, Priority.ALWAYS);
        VBox root =
                new VBox(
                        12,
                        new Label("会社カレンダー（祝日・公休）とメンバー勤怠の初期生成"),
                        yearRow,
                        fetchBtn,
                        syncBtn,
                        status,
                        spacer,
                        closeBtn);
        root.setPadding(new Insets(16));
        root.setPrefWidth(420);
        stage.setScene(new javafx.scene.Scene(root));
        stage.showAndWait();
    }
}
