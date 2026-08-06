package jp.co.pm.ai.desktop.ui;

import java.util.HashMap;
import java.util.Map;
import java.util.function.Consumer;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.StageStyle;

/** 1 日の時間別勤務モード編集（6:00〜22:00）。 */
public final class MemberHourlyAttendanceDialog {

    public static final String MODE_AVAILABLE = "available";
    public static final String MODE_BREAK = "break";
    public static final String MODE_LEAVE = "leave";
    public static final String MODE_AWAY = "away";
    public static final String MODE_OFF_SHIFT = "off_shift";

    private MemberHourlyAttendanceDialog() {}

    public static void show(
            Stage owner,
            String member,
            String dateKey,
            Map<String, String> initialHourly,
            Consumer<Map<String, String>> onSave) {
        Stage stage = new Stage(StageStyle.DECORATED);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle(member + " — " + dateKey + " 時間別");

        Map<String, ComboBox<String>> pickers = new HashMap<>();
        GridPane grid = new GridPane();
        grid.setHgap(6);
        grid.setVgap(4);
        grid.setPadding(new Insets(8));

        String[] labels = {"稼働", "休憩", "休暇", "離席", "非シフト"};
        String[] modes =
                new String[] {
                    MODE_AVAILABLE, MODE_BREAK, MODE_LEAVE, MODE_AWAY, MODE_OFF_SHIFT
                };

        int row = 0;
        for (int hour = 6; hour <= 22; hour++) {
            String key = String.format("%02d:00", hour);
            Label hl = new Label(key);
            ComboBox<String> box = new ComboBox<>();
            for (int i = 0; i < labels.length; i++) {
                box.getItems().add(labels[i]);
            }
            String init = initialHourly != null ? initialHourly.get(key) : null;
            if (init == null || init.isBlank()) {
                box.getSelectionModel().select(0);
            } else {
                boolean matched = false;
                for (int i = 0; i < modes.length; i++) {
                    if (modes[i].equals(init)) {
                        box.getSelectionModel().select(i);
                        matched = true;
                        break;
                    }
                }
                if (!matched) {
                    box.getSelectionModel().select(0);
                }
            }
            pickers.put(key, box);
            grid.add(hl, 0, row);
            grid.add(box, 1, row);
            row++;
        }

        Button ok = new Button("適用");
        Button cancel = new Button("キャンセル");
        ok.setOnAction(
                e -> {
                    Map<String, String> out = new HashMap<>();
                    for (var ent : pickers.entrySet()) {
                        int idx = ent.getValue().getSelectionModel().getSelectedIndex();
                        if (idx >= 0 && idx < modes.length) {
                            String mode = modes[idx];
                            if (!MODE_AVAILABLE.equals(mode)) {
                                out.put(ent.getKey(), mode);
                            }
                        }
                    }
                    if (onSave != null) {
                        onSave.accept(out);
                    }
                    stage.close();
                });
        cancel.setOnAction(e -> stage.close());

        HBox buttons = new HBox(8, ok, cancel);
        buttons.setAlignment(Pos.CENTER_RIGHT);
        VBox root = new VBox(8, new Label("時間帯ごとのモード（空欄相当は稼働）"), grid, buttons);
        root.setPadding(new Insets(12));
        stage.setScene(new javafx.scene.Scene(root, 320, 520));
        stage.showAndWait();
    }
}
