package jp.co.pm.ai.desktop;

import java.time.LocalDate;
import java.util.function.Consumer;

import com.fasterxml.jackson.databind.JsonNode;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Separator;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

import jp.co.pm.ai.desktop.ui.AttendanceSyncStatusPane;
import jp.co.pm.ai.desktop.ui.FiscalYearPeriod;

/**
 * 初回／再セットアップ: 祝日取得 → メンバー同期 → 状態表示。
 *
 * <p>祝日・週末公休の取得およびメンバー勤怠の会社カレンダー同期は本ウィザードに一本化。
 */
public final class AttendanceSetupWizard {

    private static final int WIZARD_MIN_WIDTH = 720;
    private static final int WIZARD_PREF_WIDTH = 780;

    private AttendanceSetupWizard() {}

    public static void show(MainShellController shell, Consumer<Boolean> onComplete) {
        LocalDate today = LocalDate.now();
        int fiscalYear =
                FiscalYearPeriod.fiscalYearLabelFor(
                        today, FiscalYearPeriod.DEFAULT_APRIL_MARCH);
        show(
                shell,
                fiscalYear,
                FiscalYearPeriod.DEFAULT_APRIL_MARCH,
                today.getYear(),
                today.getMonthValue(),
                onComplete);
    }

    public static void show(
            MainShellController shell,
            int fiscalYear,
            FiscalYearPeriod fiscalPeriod,
            Consumer<Boolean> onComplete) {
        LocalDate today = LocalDate.now();
        show(
                shell,
                fiscalYear,
                fiscalPeriod,
                today.getYear(),
                today.getMonthValue(),
                onComplete);
    }

    public static void show(
            MainShellController shell,
            int fiscalYear,
            FiscalYearPeriod fiscalPeriod,
            int syncYear,
            int syncMonth,
            Consumer<Boolean> onComplete) {
        if (shell == null) {
            return;
        }
        FiscalYearPeriod period =
                fiscalPeriod != null ? fiscalPeriod : FiscalYearPeriod.DEFAULT_APRIL_MARCH;

        Stage stage = new Stage(StageStyle.DECORATED);
        stage.initOwner(shell.primaryStageForDialogs());
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle("勤怠カレンダー セットアップ");
        stage.setMinWidth(WIZARD_MIN_WIDTH);

        Label status = new Label("祝日・週末公休の取得とメンバー勤怠の会社カレンダー同期を順に実行します。");
        status.setWrapText(true);
        status.setMaxWidth(Double.MAX_VALUE);

        Spinner<Integer> fiscalYearSpinner = yearSpinner(fiscalYear);
        Spinner<Integer> fiscalStartMonthSpinner = monthSpinner(period.startMonth());
        Spinner<Integer> fiscalStartDaySpinner = daySpinner(period.startDay());
        Spinner<Integer> syncYearSpinner = yearSpinner(syncYear);
        Spinner<Integer> monthSpinner = monthSpinner(syncMonth);

        final boolean[] fetchOk = {false};
        final boolean[] syncOk = {false};
        final boolean[] running = {false};

        Button fetchBtn = new Button("1. 祝日・週末公休を取得");
        Button syncBtn = new Button("2. メンバー勤怠を会社カレンダーに合わせる（1ヶ月）");
        Button syncFiscalBtn = new Button("2b. メンバー勤怠を年度一括で同期");
        Button runAllBtn = new Button("すべて実行（1 → 2：1ヶ月）");
        Button runAllFiscalBtn = new Button("すべて実行（1 → 2：年度一括）");
        runAllFiscalBtn.setDefaultButton(true);
        Button closeBtn = new Button("閉じる");

        AttendanceSyncStatusPane syncStatusPane = new AttendanceSyncStatusPane();

        final boolean[] onCompleteInvoked = {false};
        Runnable invokeOnCompleteOnce =
                () -> {
                    if (onCompleteInvoked[0]) {
                        return;
                    }
                    onCompleteInvoked[0] = true;
                    if (onComplete != null) {
                        onComplete.accept(fetchOk[0] || syncOk[0]);
                    }
                };

        Runnable refreshReadiness =
                () ->
                        shell.runAttendanceDataIoAsync(
                                shell.buildAttendanceDataIoRequest(
                                        "readiness",
                                        Integer.toString(syncYearSpinner.getValue()),
                                        Integer.toString(monthSpinner.getValue())),
                                syncStatusPane::updateFromReadiness,
                                err -> {});

        Consumer<Boolean> setRunning =
                busy -> {
                    running[0] = busy;
                    fetchBtn.setDisable(busy);
                    syncBtn.setDisable(busy);
                    syncFiscalBtn.setDisable(busy);
                    runAllBtn.setDisable(busy);
                    runAllFiscalBtn.setDisable(busy);
                    closeBtn.setDisable(busy);
                    fiscalYearSpinner.setDisable(busy);
                    fiscalStartMonthSpinner.setDisable(busy);
                    fiscalStartDaySpinner.setDisable(busy);
                    syncYearSpinner.setDisable(busy);
                    monthSpinner.setDisable(busy);
                };

        Runnable closeAfterSuccess =
                () -> {
                    setRunning.accept(false);
                    stage.close();
                };

        Runnable runFetch =
                () -> {
                    setRunning.accept(true);
                    status.setText("1/2 祝日・週末公休を取得中…");
                    shell.runAttendanceDataIoAsync(
                            shell.buildAttendanceDataIoRequest(
                                    "fetch_holidays_fiscal",
                                    Integer.toString(fiscalYearSpinner.getValue()),
                                    Integer.toString(fiscalStartMonthSpinner.getValue()),
                                    Integer.toString(fiscalStartDaySpinner.getValue()),
                                    "--weekends"),
                            node -> {
                                fetchOk[0] = true;
                                status.setText(
                                        "1/2 完了: 適用 "
                                                + node.path("applied").asInt(0)
                                                + " 日 / スキップ "
                                                + node.path("skipped").asInt(0));
                                shell.refreshAttendanceReadiness();
                                refreshReadiness.run();
                                setRunning.accept(false);
                            },
                            err -> {
                                status.setText("祝日取得失敗: " + err);
                                setRunning.accept(false);
                            });
                };

        Runnable runSyncMonth =
                () -> {
                    setRunning.accept(true);
                    status.setText("2/2 メンバー勤怠を同期中（1ヶ月）…");
                    int year = syncYearSpinner.getValue();
                    int month = monthSpinner.getValue();
                    shell.runAttendanceDataIoAsync(
                            shell.buildAttendanceDataIoRequest(
                                    "sync_members",
                                    Integer.toString(year), Integer.toString(month)),
                            node -> {
                                syncOk[0] = true;
                                status.setText(
                                        "2/2 完了（1ヶ月）: 適用 "
                                                + node.path("applied").asInt(0)
                                                + " / スキップ "
                                                + node.path("skipped").asInt(0));
                                shell.refreshAttendanceReadiness();
                                refreshReadiness.run();
                                closeAfterSuccess.run();
                            },
                            err -> {
                                status.setText("同期失敗: " + err);
                                setRunning.accept(false);
                            });
                };

        Runnable runSyncFiscal =
                () -> {
                    setRunning.accept(true);
                    int fy = fiscalYearSpinner.getValue();
                    int sm = fiscalStartMonthSpinner.getValue();
                    int sd = fiscalStartDaySpinner.getValue();
                    status.setText("2/2 メンバー勤怠を同期中（会計年度 " + fy + "）…");
                    shell.runAttendanceDataIoAsync(
                            shell.buildAttendanceDataIoRequest(
                                    "sync_members_fiscal",
                                    Integer.toString(fy),
                                    Integer.toString(sm),
                                    Integer.toString(sd)),
                            node -> {
                                syncOk[0] = true;
                                status.setText(
                                        "2/2 完了（年度一括）: 適用 "
                                                + node.path("applied").asInt(0)
                                                + " / スキップ "
                                                + node.path("skipped").asInt(0)
                                                + " （"
                                                + node.path("fiscal_start").asText("")
                                                + " ～ "
                                                + node.path("fiscal_end").asText("")
                                                + "）");
                                shell.refreshAttendanceReadiness();
                                refreshReadiness.run();
                                closeAfterSuccess.run();
                            },
                            err -> {
                                status.setText("年度一括同期失敗: " + err);
                                setRunning.accept(false);
                            });
                };

        fetchBtn.setOnAction(e -> runFetch.run());
        syncBtn.setOnAction(e -> runSyncMonth.run());
        syncFiscalBtn.setOnAction(e -> runSyncFiscal.run());

        runAllBtn.setOnAction(
                e -> {
                    setRunning.accept(true);
                    status.setText("1/2 祝日・週末公休を取得中…");
                    shell.runAttendanceDataIoAsync(
                            shell.buildAttendanceDataIoRequest(
                                    "fetch_holidays_fiscal",
                                    Integer.toString(fiscalYearSpinner.getValue()),
                                    Integer.toString(fiscalStartMonthSpinner.getValue()),
                                    Integer.toString(fiscalStartDaySpinner.getValue()),
                                    "--weekends"),
                            node -> {
                                fetchOk[0] = true;
                                status.setText("1/2 完了。2/2 メンバー勤怠を同期中（1ヶ月）…");
                                shell.refreshAttendanceReadiness();
                                int year = syncYearSpinner.getValue();
                                int month = monthSpinner.getValue();
                                shell.runAttendanceDataIoAsync(
                                        shell.buildAttendanceDataIoRequest(
                                                "sync_members",
                                                Integer.toString(year),
                                                Integer.toString(month)),
                                        syncNode -> {
                                            syncOk[0] = true;
                                            status.setText(
                                                    "すべて完了（1ヶ月）: 祝日 "
                                                            + node.path("applied").asInt(0)
                                                            + " 日 / メンバー適用 "
                                                            + syncNode.path("applied").asInt(0));
                                            shell.refreshAttendanceReadiness();
                                            refreshReadiness.run();
                                            closeAfterSuccess.run();
                                        },
                                        err -> {
                                            status.setText("同期失敗: " + err);
                                            setRunning.accept(false);
                                        });
                            },
                            err -> {
                                status.setText("祝日取得失敗: " + err);
                                setRunning.accept(false);
                            });
                });

        runAllFiscalBtn.setOnAction(
                e -> {
                    setRunning.accept(true);
                    status.setText("1/2 祝日・週末公休を取得中…");
                    shell.runAttendanceDataIoAsync(
                            shell.buildAttendanceDataIoRequest(
                                    "fetch_holidays_fiscal",
                                    Integer.toString(fiscalYearSpinner.getValue()),
                                    Integer.toString(fiscalStartMonthSpinner.getValue()),
                                    Integer.toString(fiscalStartDaySpinner.getValue()),
                                    "--weekends"),
                            node -> {
                                fetchOk[0] = true;
                                status.setText("1/2 完了。2/2 メンバー勤怠を同期中（年度一括）…");
                                shell.refreshAttendanceReadiness();
                                int fy = fiscalYearSpinner.getValue();
                                int sm = fiscalStartMonthSpinner.getValue();
                                int sd = fiscalStartDaySpinner.getValue();
                                shell.runAttendanceDataIoAsync(
                                        shell.buildAttendanceDataIoRequest(
                                                "sync_members_fiscal",
                                                Integer.toString(fy),
                                                Integer.toString(sm),
                                                Integer.toString(sd)),
                                        syncNode -> {
                                            syncOk[0] = true;
                                            status.setText(
                                                    "すべて完了（年度一括）: 祝日 "
                                                            + node.path("applied").asInt(0)
                                                            + " 日 / メンバー適用 "
                                                            + syncNode.path("applied").asInt(0));
                                            shell.refreshAttendanceReadiness();
                                            refreshReadiness.run();
                                            closeAfterSuccess.run();
                                        },
                                        err -> {
                                            status.setText("年度一括同期失敗: " + err);
                                            setRunning.accept(false);
                                        });
                            },
                            err -> {
                                status.setText("祝日取得失敗: " + err);
                                setRunning.accept(false);
                            });
                });

        closeBtn.setOnAction(
                e -> {
                    if (running[0]) {
                        return;
                    }
                    stage.close();
                });

        stage.setOnCloseRequest(
                e -> {
                    if (running[0]) {
                        e.consume();
                        return;
                    }
                    invokeOnCompleteOnce.run();
                });

        Label titleLabel = new Label("会社カレンダー（祝日・公休）とメンバー勤怠の初期設定");
        titleLabel.getStyleClass().add("pm-attendance-setup-wizard-title");
        titleLabel.setWrapText(true);
        titleLabel.setMaxWidth(Double.MAX_VALUE);

        Label holidaySection = sectionLabel("1. 祝日・週末公休の取得");
        HBox holidayRow =
                labeledRow(
                        "会計年度",
                        fiscalYearSpinner,
                        "期間開始",
                        fiscalStartMonthSpinner,
                        "日",
                        fiscalStartDaySpinner);

        Label monthSyncSection = sectionLabel("2. メンバー同期（1ヶ月）");
        Label monthHelp =
                helpLabel(
                        "対象年月の1ヶ月分のみ更新します。未編集セルに会社カレンダーの休日設定を反映します。");
        HBox monthRow =
                labeledRow("対象年", syncYearSpinner, "月", monthSpinner);

        Label fiscalSyncSection = sectionLabel("2b. メンバー同期（会計年度一括）");
        Label fiscalHelp =
                helpLabel(
                        "「1. 祝日取得」と同じ会計年度・期間開始の全日を一括更新します。"
                                + "手動編集済みセルはスキップします。");

        HBox runAllRow = new HBox(12, runAllBtn, runAllFiscalBtn);
        runAllRow.setAlignment(Pos.CENTER_LEFT);

        Region spacer = new Region();
        VBox.setVgrow(spacer, Priority.ALWAYS);

        VBox root =
                new VBox(
                        10,
                        titleLabel,
                        new Separator(),
                        holidaySection,
                        holidayRow,
                        fetchBtn,
                        new Separator(),
                        monthSyncSection,
                        monthHelp,
                        monthRow,
                        syncBtn,
                        new Separator(),
                        fiscalSyncSection,
                        fiscalHelp,
                        syncFiscalBtn,
                        new Separator(),
                        runAllRow,
                        status,
                        syncStatusPane,
                        spacer,
                        closeBtn);
        root.getStyleClass().add("pm-attendance-setup-wizard");
        root.setPadding(new Insets(16));
        root.setPrefWidth(WIZARD_PREF_WIDTH);
        root.setFillWidth(true);

        javafx.scene.Scene scene = new javafx.scene.Scene(root);
        shell.applyStylesheetsToScene(scene);
        stage.setScene(scene);
        refreshReadiness.run();
        stage.showAndWait();
    }

    private static Label sectionLabel(String text) {
        Label label = new Label(text);
        label.getStyleClass().add("pm-attendance-setup-wizard-section");
        label.setWrapText(true);
        label.setMaxWidth(Double.MAX_VALUE);
        return label;
    }

    private static Label helpLabel(String text) {
        Label label = new Label(text);
        label.setWrapText(true);
        label.setMaxWidth(Double.MAX_VALUE);
        label.getStyleClass().add("pm-attendance-setup-wizard-help");
        return label;
    }

    private static Spinner<Integer> yearSpinner(int value) {
        Spinner<Integer> spinner =
                new Spinner<>(new SpinnerValueFactory.IntegerSpinnerValueFactory(2020, 2040, value));
        spinner.setEditable(true);
        spinner.setPrefWidth(92);
        spinner.setMinWidth(92);
        return spinner;
    }

    private static Spinner<Integer> monthSpinner(int value) {
        Spinner<Integer> spinner =
                new Spinner<>(new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 12, value));
        spinner.setEditable(true);
        spinner.setPrefWidth(64);
        spinner.setMinWidth(64);
        return spinner;
    }

    private static Spinner<Integer> daySpinner(int value) {
        Spinner<Integer> spinner =
                new Spinner<>(new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 31, value));
        spinner.setEditable(true);
        spinner.setPrefWidth(64);
        spinner.setMinWidth(64);
        return spinner;
    }

  /** ラベルとスピナーのペアを横並びにする（可変長）。 */
    private static HBox labeledRow(Object... labelSpinnerPairs) {
        HBox row = new HBox(12);
        row.setAlignment(Pos.CENTER_LEFT);
        for (int i = 0; i < labelSpinnerPairs.length; i += 2) {
            if (i + 1 >= labelSpinnerPairs.length) {
                break;
            }
            String labelText = (String) labelSpinnerPairs[i];
            Spinner<Integer> spinner = (Spinner<Integer>) labelSpinnerPairs[i + 1];
            row.getChildren().add(fieldPair(labelText, spinner));
        }
        return row;
    }

    private static HBox fieldPair(String labelText, Spinner<Integer> spinner) {
        Label label = new Label(labelText);
        label.setMinWidth(Region.USE_PREF_SIZE);
        label.setWrapText(false);
        HBox pair = new HBox(6, label, spinner);
        pair.setAlignment(Pos.CENTER_LEFT);
        return pair;
    }
}
