package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.MasterDispatchSetupPrompt;
import jp.co.pm.ai.desktop.ui.MasterDispatchSetupCompleteness.EquipmentRef;
import jp.co.pm.ai.desktop.ui.MasterDispatchSetupCompleteness.EquipmentStatus;
import jp.co.pm.ai.desktop.ui.MasterDispatchSetupCompleteness.Evaluation;
import jp.co.pm.ai.desktop.ui.MasterDispatchSetupCompleteness.Step;

/**
 * 配台マスタ設定ウィザード（skills → need → 組み合わせ表 → 加工速度）。
 */
public final class MasterDispatchSetupWizardDialog {

    public enum Mode {
        /** 段階1後: 後でスキップ可 */
        AFTER_STAGE1,
        /** 段階2前: 未完了なら実行不可 */
        BEFORE_STAGE2
    }

    public enum Outcome {
        CANCELLED,
        SKIPPED,
        COMPLETED
    }

    public record Session(
            List<List<String>> skills,
            List<List<String>> need,
            List<List<String>> combinations,
            List<List<String>> speed) {}

    public record Result(Outcome outcome, Session session) {}

    private MasterDispatchSetupWizardDialog() {}

    /**
     * @param initial 編集開始時の4シート（null なら空）
     * @param equipment 計画上の工程+機械
     */
    public static Optional<Result> run(
            Window owner, Mode mode, List<EquipmentRef> equipment, Session initial) {
        Session session =
                initial != null
                        ? initial
                        : new Session(List.of(), List.of(), List.of(), List.of());
        Evaluation eval =
                MasterDispatchSetupCompleteness.evaluate(
                        equipment,
                        session.skills(),
                        session.need(),
                        session.combinations(),
                        session.speed());
        if (eval.allComplete()) {
            return Optional.of(new Result(Outcome.COMPLETED, session));
        }

        if (!confirmStart(owner, mode, eval)) {
            return Optional.of(
                    new Result(
                            mode == Mode.AFTER_STAGE1 ? Outcome.SKIPPED : Outcome.CANCELLED,
                            session));
        }

        List<EquipmentStatus> todo = new ArrayList<>(eval.incomplete());
        for (EquipmentStatus status : todo) {
            EquipmentRef eq = status.equipment();
            for (Step step : Step.values()) {
                if (!status.incompleteSteps().contains(step)) {
                    continue;
                }
                Optional<Session> next = runStep(owner, eq, step, session);
                if (next.isEmpty()) {
                    return Optional.of(new Result(Outcome.CANCELLED, session));
                }
                session = next.get();
            }
        }

        Evaluation after =
                MasterDispatchSetupCompleteness.evaluate(
                        equipment,
                        session.skills(),
                        session.need(),
                        session.combinations(),
                        session.speed());
        if (!after.allComplete()) {
            Alert alert = new Alert(Alert.AlertType.WARNING);
            if (owner != null) {
                alert.initOwner(owner);
            }
            alert.setTitle("配台マスタ設定");
            alert.setHeaderText("まだ未完了の項目があります。");
            alert.setContentText(after.summaryJa(12));
            alert.showAndWait();
            return Optional.of(new Result(Outcome.CANCELLED, session));
        }
        return Optional.of(new Result(Outcome.COMPLETED, session));
    }

    private static boolean confirmStart(Window owner, Mode mode, Evaluation eval) {
        Dialog<ButtonType> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle(
                mode == Mode.BEFORE_STAGE2 ? "段階2 — 配台マスタ設定が必要" : "配台マスタ設定ウィザード");
        dialog.setHeaderText(
                mode == Mode.BEFORE_STAGE2
                        ? "配台マスタの設定が完了していないため、段階2は実行できません。"
                        : "計画タスクに、配台マスタ未設定の工程+機械があります。");
        Label body =
                new Label(
                        "設定順: ①資格 → ②必要人数 → ③組み合わせ表 → ④加工速度\n\n"
                                + eval.summaryJa(16)
                                + (mode == Mode.BEFORE_STAGE2
                                        ? "\n\n「設定を開始」でウィザードを開きます。キャンセルすると段階2は実行しません。"
                                        : "\n\n「設定を開始」でウィザードを開きます。「後で」でスキップできます。"));
        body.setWrapText(true);
        dialog.getDialogPane().setContent(body);
        dialog.getDialogPane().setPrefWidth(640);
        ButtonType start = new ButtonType("設定を開始", ButtonBar.ButtonData.OK_DONE);
        if (mode == Mode.AFTER_STAGE1) {
            ButtonType later = new ButtonType("後で", ButtonBar.ButtonData.CANCEL_CLOSE);
            dialog.getDialogPane().getButtonTypes().setAll(start, later);
        } else {
            dialog.getDialogPane().getButtonTypes().setAll(start, ButtonType.CANCEL);
        }
        Optional<ButtonType> choice = dialog.showAndWait();
        return choice.isPresent() && choice.get() == start;
    }

    private static Optional<Session> runStep(
            Window owner, EquipmentRef eq, Step step, Session session) {
        return switch (step) {
            case SKILLS -> runSkillsStep(owner, eq, session);
            case NEED -> runNeedStep(owner, eq, session);
            case COMBINATIONS -> runCombinationsStep(owner, eq, session);
            case SPEED -> runSpeedStep(owner, eq, session);
        };
    }

    public static final class SkillRow {
        private final String memberName;
        private String role;

        public SkillRow(String memberName, String role) {
            this.memberName = memberName;
            this.role = role != null ? role : "";
        }

        public String getMemberName() {
            return memberName;
        }

        public String getRole() {
            return role;
        }

        public void setRole(String role) {
            this.role = role != null ? role : "";
        }
    }

    private static Optional<Session> runSkillsStep(Window owner, EquipmentRef eq, Session session) {
        List<List<String>> skills =
                MasterDispatchSheetEditRules.addEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS,
                        session.skills(),
                        eq.process(),
                        eq.machine());
        List<String> names = MasterDispatchSheetEditRules.skillsMemberNames(skills);
        List<String> skilled =
                MasterDispatchSetupCompleteness.skilledOpAsMembers(
                        skills, eq.process(), eq.machine());
        List<SkillRow> rows = new ArrayList<>();
        for (String name : names) {
            String role = "";
            for (String s : skilled) {
                String n = MasterDispatchSheetEditRules.combinationMemberName(s);
                if (name.equals(n)) {
                    if (s.toUpperCase().startsWith("OP")) {
                        role = "OP";
                    } else if (s.toUpperCase().startsWith("AS")) {
                        role = "AS";
                    }
                    break;
                }
            }
            rows.add(new SkillRow(name, role));
        }
        if (rows.isEmpty()) {
            rows.add(new SkillRow("（メンバーを追加）", ""));
        }

        Dialog<ButtonType> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle("① 資格（skills）");
        dialog.setHeaderText(eq.display() + " — OP/AS を1人以上設定してください。");

        TableView<SkillRow> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setEditable(true);
        table.setPrefHeight(Math.min(400, 80 + rows.size() * 28.0));
        TableColumn<SkillRow, String> cName = new TableColumn<>("メンバー");
        cName.setCellValueFactory(new PropertyValueFactory<>("memberName"));
        cName.setEditable(false);
        TableColumn<SkillRow, String> cRole = new TableColumn<>("資格");
        cRole.setCellValueFactory(new PropertyValueFactory<>("role"));
        cRole.setCellFactory(
                col ->
                        new javafx.scene.control.TableCell<>() {
                            private final ComboBox<String> combo =
                                    new ComboBox<>(
                                            FXCollections.observableArrayList("", "OP", "AS"));

                            {
                                combo.setMaxWidth(Double.MAX_VALUE);
                                combo.setOnAction(
                                        e -> {
                                            SkillRow item = getTableView().getItems().get(getIndex());
                                            if (item != null) {
                                                item.setRole(combo.getValue());
                                            }
                                        });
                            }

                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || getIndex() < 0 || getIndex() >= getTableView().getItems().size()) {
                                    setGraphic(null);
                                    return;
                                }
                                SkillRow row = getTableView().getItems().get(getIndex());
                                combo.setValue(row.getRole() != null ? row.getRole() : "");
                                setGraphic(combo);
                            }
                        });
        table.getColumns().setAll(cName, cRole);

        Label hint = new Label("次へ の前に、少なくとも1人に OP または AS を選んでください。");
        hint.setWrapText(true);
        VBox root = new VBox(8, hint, table);
        VBox.setVgrow(table, Priority.ALWAYS);
        root.setPadding(new Insets(4, 0, 0, 0));
        dialog.getDialogPane().setContent(root);
        dialog.getDialogPane().setPrefWidth(520);
        ButtonType next = new ButtonType("次へ", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().setAll(next, ButtonType.CANCEL);

        while (true) {
            Optional<ButtonType> choice = dialog.showAndWait();
            if (choice.isEmpty() || choice.get() != next) {
                return Optional.empty();
            }
            List<List<String>> updated = skills;
            int assigned = 0;
            for (SkillRow row : table.getItems()) {
                String name = row.getMemberName();
                if (name == null || name.isBlank() || name.startsWith("（")) {
                    continue;
                }
                String role = row.getRole() != null ? row.getRole().strip() : "";
                updated =
                        MasterDispatchSheetEditRules.setSkillRoleForMember(
                                updated, eq.process(), eq.machine(), name, role);
                if (!role.isEmpty()) {
                    assigned++;
                }
            }
            if (assigned < 1) {
                Alert warn = new Alert(Alert.AlertType.WARNING);
                if (owner != null) {
                    warn.initOwner(owner);
                }
                warn.setTitle("① 資格");
                warn.setHeaderText("OP/AS が未設定です。");
                warn.setContentText("少なくとも1人に OP または AS を設定してください。");
                warn.showAndWait();
                continue;
            }
            return Optional.of(
                    new Session(
                            updated, session.need(), session.combinations(), session.speed()));
        }
    }

    private static Optional<Session> runNeedStep(Window owner, EquipmentRef eq, Session session) {
        int current =
                MasterDispatchSetupCompleteness.readBaseRequiredHeadcount(
                                session.need(), eq.process(), eq.machine())
                        .orElse(1);
        Dialog<ButtonType> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle("② 必要人数（need）");
        dialog.setHeaderText(eq.display() + " — 基本必要人数を入力してください。");
        Spinner<Integer> spinner = new Spinner<>();
        spinner.setValueFactory(new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 99, current));
        spinner.setEditable(true);
        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.setPadding(new Insets(8));
        grid.add(new Label("基本必要人数"), 0, 0);
        grid.add(spinner, 1, 0);
        dialog.getDialogPane().setContent(grid);
        ButtonType next = new ButtonType("次へ", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().setAll(next, ButtonType.CANCEL);
        Optional<ButtonType> choice = dialog.showAndWait();
        if (choice.isEmpty() || choice.get() != next) {
            return Optional.empty();
        }
        int k = spinner.getValue() != null ? spinner.getValue() : 1;
        List<List<String>> need =
                MasterDispatchSheetEditRules.setBaseRequiredHeadcount(
                        session.need(), eq.process(), eq.machine(), k);
        return Optional.of(
                new Session(session.skills(), need, session.combinations(), session.speed()));
    }

    private static Optional<Session> runCombinationsStep(
            Window owner, EquipmentRef eq, Session session) {
        int k =
                MasterDispatchSheetEditRules.baseRequiredHeadcount(
                        session.need(), eq.process(), eq.machine());
        List<String> members =
                MasterDispatchSetupCompleteness.skilledOpAsMembers(
                        session.skills(), eq.process(), eq.machine());
        Dialog<ButtonType> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle("③ 組み合わせ表");
        dialog.setHeaderText(eq.display() + " — スキル保持者の組合せ行を追加します。");
        Label body =
                new Label(
                        "必要人数 "
                                + k
                                + " 人、スキル保持者 "
                                + members.size()
                                + " 人から組合せを生成します。\n"
                                + (members.size() >= k
                                        ? "「次へ」で組み合わせ表へ行を追加します。"
                                        : "スキル人数が不足しているため、メンバー空の行を1行追加します。"
                                                + " ①に戻って資格を増やしてください。"));
        body.setWrapText(true);
        dialog.getDialogPane().setContent(body);
        ButtonType next = new ButtonType("次へ", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().setAll(next, ButtonType.CANCEL);
        Optional<ButtonType> choice = dialog.showAndWait();
        if (choice.isEmpty() || choice.get() != next) {
            return Optional.empty();
        }
        List<List<String>> combo =
                MasterDispatchSheetEditRules.ensureSkillCombinations(
                        session.combinations(),
                        eq.process(),
                        eq.machine(),
                        session.skills(),
                        session.need());
        return Optional.of(new Session(session.skills(), session.need(), combo, session.speed()));
    }

    private static Optional<Session> runSpeedStep(Window owner, EquipmentRef eq, Session session) {
        Dialog<ButtonType> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle("④ 加工速度（speed）");
        dialog.setHeaderText(eq.display() + " — 基本速度を入力してください（0.0〜99.0）。");
        TextField field = new TextField("20.0");
        field.setPrefColumnCount(8);
        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.setPadding(new Insets(8));
        grid.add(new Label("基本速度"), 0, 0);
        grid.add(field, 1, 0);
        dialog.getDialogPane().setContent(grid);
        ButtonType next = new ButtonType("完了（この設備）", ButtonBar.ButtonData.OK_DONE);
        dialog.getDialogPane().getButtonTypes().setAll(next, ButtonType.CANCEL);

        while (true) {
            Optional<ButtonType> choice = dialog.showAndWait();
            if (choice.isEmpty() || choice.get() != next) {
                return Optional.empty();
            }
            String v = field.getText() != null ? field.getText().strip() : "";
            if (!MasterDispatchSheetEditRules.isSpeedBaseDecimalValid(v)) {
                Alert warn = new Alert(Alert.AlertType.WARNING);
                if (owner != null) {
                    warn.initOwner(owner);
                }
                warn.setTitle("④ 加工速度");
                warn.setHeaderText("基本速度が不正です。");
                warn.setContentText("0.0〜99.0（小数第一位まで）で入力してください。");
                warn.showAndWait();
                continue;
            }
            List<List<String>> speed =
                    MasterDispatchSheetEditRules.setBaseSpeed(
                            session.speed(), eq.process(), eq.machine(), v);
            return Optional.of(
                    new Session(
                            session.skills(), session.need(), session.combinations(), speed));
        }
    }

    public static MasterDispatchSetupPrompt.SheetBundle toSheetBundle(Session session) {
        return new MasterDispatchSetupPrompt.SheetBundle(
                session.skills(), session.need(), session.combinations(), session.speed());
    }
}
