package jp.co.pm.ai.desktop;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.RadioButton;
import javafx.scene.control.ToggleGroup;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.application.Platform;
import javafx.stage.Modality;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.dispatch.DispatchTrialShortages;
import jp.co.pm.ai.desktop.dispatch.ResultTaskUnassignedLoader;

/**
 * 配台試行後に「配台不可」タスクがあるとき、次に取るアクションをウィザード形式で選ばせる。
 */
public final class DispatchTrialUnassignedWizard {

    private DispatchTrialUnassignedWizard() {}

    private enum NextStepAction {
        CLOSE,
        PLAN_VIEWER,
        EQUIPMENT_GANTT,
        INTERACTIVE_EDIT,
        OPEN_OUTPUT_FOLDER
    }

    /**
     * {@code dispatch_trial_shortages.json} と結果_タスク一覧を参照し、配台不可があればモーダルウィザードを表示する。
     */
    public static void showIfNeeded(Stage owner, MainShellController shell, Path shortageJsonPath) {
        Objects.requireNonNull(shell, "shell");
        if (shortageJsonPath == null) {
            return;
        }
        Map<String, String> uiSnap = shell.snapshotUiEnv();
        Thread worker =
                new Thread(
                        () -> {
                            try {
                                DispatchTrialShortages.Paths paths =
                                        DispatchTrialShortages.read(shortageJsonPath);
                                if (paths.productionPlan() == null || paths.productionPlan().isBlank()) {
                                    return;
                                }
                                List<ResultTaskUnassignedLoader.UnassignedRow> rows =
                                        ResultTaskUnassignedLoader.loadUnassigned(
                                                uiSnap, paths.productionPlan());
                                if (rows.isEmpty()) {
                                    return;
                                }
                                Platform.runLater(() -> showDialog(owner, shell, paths, rows));
                            } catch (Exception e) {
                                Platform.runLater(
                                        () ->
                                                shell.appendLog(
                                                        "[dispatch-wizard] 配台不可ウィザードを開けませんでした: "
                                                                + (e.getMessage() != null
                                                                        ? e.getMessage()
                                                                        : e)));
                            }
                        },
                        "dispatch-unassigned-wizard");
        worker.setDaemon(true);
        worker.start();
    }

    private static void showDialog(
            Stage owner,
            MainShellController shell,
            DispatchTrialShortages.Paths paths,
            List<ResultTaskUnassignedLoader.UnassignedRow> rows) {
        Stage stage = new Stage();
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle("配台不可タスク — 次の操作");

        Label stepLabel = new Label("ステップ 1 / 2");
        stepLabel.setStyle("-fx-font-weight: bold;");

        Label intro =
                new Label(
                        "配台試行の結果、計画上「配台不可」のタスクが "
                                + rows.size()
                                + " 件あります。"
                                + " 一覧を確認し、次のステップで対応を選んでください。");
        intro.setWrapText(true);

        ListView<String> list = new ListView<>();
        for (ResultTaskUnassignedLoader.UnassignedRow r : rows) {
            list.getItems()
                    .add(
                            r.taskId()
                                    + "  |  "
                                    + r.processName()
                                    + "  |  "
                                    + r.machineName());
        }
        list.setPrefHeight(Math.min(280, 36 + rows.size() * 26));
        VBox.setVgrow(list, Priority.ALWAYS);

        VBox step1 = new VBox(10, intro, list);
        step1.setPadding(new Insets(0, 0, 8, 0));

        ToggleGroup group = new ToggleGroup();
        RadioButton rClose =
                new RadioButton("このダイアログだけ閉じる（結果の確認は別途）");
        RadioButton rViewer =
                new RadioButton("計画結果ビューアで production_plan / member_schedule の JSON を開く");
        RadioButton rGantt = new RadioButton("設備ガントで計画を確認する");
        RadioButton rEdit =
                new RadioButton("配台計画手動修正タブへ戻り、表を編集して保存後に再度配台試行する");
        RadioButton rFolder = new RadioButton("出力フォルダをファイルマネージャで開く");
        rClose.setToggleGroup(group);
        rViewer.setToggleGroup(group);
        rGantt.setToggleGroup(group);
        rEdit.setToggleGroup(group);
        rFolder.setToggleGroup(group);
        rViewer.setSelected(true);

        Label step2Title =
                new Label(
                        "ステップ 2: どれを実行しますか？（タブ切替やフォルダを開く処理をその場で行います）");
        step2Title.setWrapText(true);

        VBox step2 = new VBox(8, step2Title, rClose, rViewer, rGantt, rEdit, rFolder);
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
                    center.setCenter(step2);
                    backBtn.setDisable(false);
                    nextBtn.setText("実行");
                };

        showStep1.run();

        nextBtn.setOnAction(
                ev -> {
                    if (center.getCenter() == step1) {
                        showStep2.run();
                    } else {
                        NextStepAction choice;
                        if (rClose.isSelected()) {
                            choice = NextStepAction.CLOSE;
                        } else if (rViewer.isSelected()) {
                            choice = NextStepAction.PLAN_VIEWER;
                        } else if (rGantt.isSelected()) {
                            choice = NextStepAction.EQUIPMENT_GANTT;
                        } else if (rEdit.isSelected()) {
                            choice = NextStepAction.INTERACTIVE_EDIT;
                        } else if (rFolder.isSelected()) {
                            choice = NextStepAction.OPEN_OUTPUT_FOLDER;
                        } else {
                            choice = NextStepAction.PLAN_VIEWER;
                        }
                        applyChoice(shell, paths, choice);
                        stage.close();
                    }
                });

        backBtn.setOnAction(ev -> showStep1.run());

        cancelBtn.setOnAction(ev -> stage.close());

        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox bottom = new HBox(8, cancelBtn, spacer, backBtn, nextBtn);
        bottom.setAlignment(Pos.CENTER_RIGHT);
        bottom.setPadding(new Insets(8, 12, 12, 12));

        BorderPane root = new BorderPane();
        root.setTop(stepLabel);
        BorderPane.setMargin(stepLabel, new Insets(12, 12, 0, 12));
        root.setCenter(center);
        root.setBottom(bottom);

        Scene scene = new Scene(root, 640, 520);
        shell.registerThemeTrackedScene(scene);
        stage.setScene(scene);
        stage.setOnHidden(ev -> shell.unregisterThemeTrackedScene(scene));
        stage.show();
    }

    private static void applyChoice(MainShellController shell, DispatchTrialShortages.Paths paths, NextStepAction c) {
        String plan = paths.productionPlan() != null ? paths.productionPlan() : "";
        String mem = paths.memberSchedule() != null ? paths.memberSchedule() : "";
        switch (c) {
            case CLOSE -> shell.appendLog("[dispatch-wizard] ユーザー: ダイアログのみ閉じる");
            case PLAN_VIEWER -> {
                shell.appendLog("[dispatch-wizard] ユーザー: 計画結果ビューアへ");
                shell.navigatePlanResultViewerWithArtifacts(plan, mem);
            }
            case EQUIPMENT_GANTT -> {
                shell.appendLog("[dispatch-wizard] ユーザー: 設備ガントへ");
                shell.navigateEquipmentGanttWithArtifacts(plan, mem);
            }
            case INTERACTIVE_EDIT -> {
                shell.appendLog("[dispatch-wizard] ユーザー: 配台計画手動修正タブへ");
                shell.navigateDispatchInteractiveTab();
            }
            case OPEN_OUTPUT_FOLDER -> {
                shell.appendLog("[dispatch-wizard] ユーザー: 出力フォルダを開く");
                shell.openDefaultPlanningOutputFolderInOs();
            }
        }
    }
}
