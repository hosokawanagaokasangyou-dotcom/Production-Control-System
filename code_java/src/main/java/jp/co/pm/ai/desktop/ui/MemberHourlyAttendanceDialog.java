package jp.co.pm.ai.desktop.ui;

import java.util.HashMap;
import java.util.Map;
import java.util.function.Consumer;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.layout.BorderPane;
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

    private static final int DIALOG_WIDTH = 340;
    private static final int DIALOG_HEIGHT = 480;
    private static final int SCROLL_VIEWPORT_HEIGHT = 360;

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
        grid.setPadding(new Insets(4, 8, 4, 8));

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
        Button clear = new Button("時間別をクリア");
        Button cancel = new Button("キャンセル");
        clear.setOnAction(
                e -> {
                    for (ComboBox<String> box : pickers.values()) {
                        box.getSelectionModel().select(0);
                    }
                });
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

        HBox buttons = new HBox(8, clear, ok, cancel);
        buttons.setAlignment(Pos.CENTER_RIGHT);
        buttons.setMaxWidth(Double.MAX_VALUE);

        Label header = new Label("時間帯ごとのモード（空欄相当は稼働）");
        VBox top = new VBox(header);
        top.setPadding(new Insets(12, 12, 4, 12));

        ScrollPane scroll = new ScrollPane(grid);
        scroll.setFitToWidth(true);
        scroll.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        scroll.setVbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);
        scroll.setPrefViewportHeight(SCROLL_VIEWPORT_HEIGHT);
        scroll.setMinViewportHeight(160);
        scroll.setMaxHeight(SCROLL_VIEWPORT_HEIGHT);

        VBox bottom = new VBox(8, buttons);
        bottom.setPadding(new Insets(8, 12, 12, 12));

        BorderPane root = new BorderPane();
        root.setTop(top);
        root.setCenter(scroll);
        root.setBottom(bottom);

        Scene scene = new Scene(root, DIALOG_WIDTH, DIALOG_HEIGHT);
        stage.setMinWidth(300);
        stage.setMinHeight(320);
        stage.setScene(scene);
        stage.showAndWait();
    }
}
