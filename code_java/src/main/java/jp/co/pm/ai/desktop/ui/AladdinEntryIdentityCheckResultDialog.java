package jp.co.pm.ai.desktop.ui;

import java.util.List;

import javafx.beans.property.ReadOnlyStringWrapper;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.io.AladdinEntryDispatchPlanIdentityCheck;
import jp.co.pm.ai.desktop.io.AladdinEntryIdentityCheckDiffTable;

/**
 * 同一化チェックの差異を表で表示し、カンマ区切り／リッチテキストでコピーする。
 */
public final class AladdinEntryIdentityCheckResultDialog {

    private AladdinEntryIdentityCheckResultDialog() {}

    public static void show(Window owner, AladdinEntryDispatchPlanIdentityCheck.Result result) {
        if (result == null) {
            return;
        }
        List<AladdinEntryDispatchPlanIdentityCheck.Diff> diffs =
                result.diffs() != null ? result.diffs() : List.of();

        Stage stage = new Stage();
        stage.initModality(Modality.WINDOW_MODAL);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.setTitle("同一化チェック");

        Label header =
                new Label(
                        result.badgeText() != null && !result.badgeText().isBlank()
                                ? result.badgeText()
                                : (result.message() != null ? result.message() : "差異"));
        header.setWrapText(true);
        header.setStyle("-fx-font-size: 16px; -fx-font-weight: bold;");

        TableView<AladdinEntryDispatchPlanIdentityCheck.Diff> table = new TableView<>();
        table.setItems(FXCollections.observableArrayList(diffs));
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_FLEX_LAST_COLUMN);
        table.setPrefHeight(Math.min(420, 80 + diffs.size() * 28));
        addColumn(table, "機械", 140, d -> nz(d.machineName()));
        addColumn(table, "依頼NO", 110, d -> nz(d.taskId()));
        addColumn(table, "工程", 90, d -> nz(d.processName()));
        addColumn(table, "日付", 110, d -> d.date() != null ? d.date().toString() : "");
        addColumn(table, "シス計", 80, d -> AladdinEntryIdentityCheckDiffTable.formatQty(d.systemQty()));
        addColumn(table, "加工計画", 90, d -> AladdinEntryIdentityCheckDiffTable.formatQty(d.planQty()));

        Button copyCsv = new Button("カンマ・改行でコピー");
        copyCsv.setOnAction(e -> copyCsv(diffs));
        Button copyRich = new Button("リッチテキストでコピー");
        copyRich.setOnAction(e -> copyRich(diffs));
        Button ok = new Button("OK");
        ok.setDefaultButton(true);
        ok.setOnAction(e -> stage.close());

        HBox buttons = new HBox(10, copyCsv, copyRich, ok);
        buttons.setAlignment(Pos.CENTER_RIGHT);

        VBox root = new VBox(12, header, table, buttons);
        root.setPadding(new Insets(16));
        VBox.setVgrow(table, Priority.ALWAYS);
        Scene scene = new Scene(root, 780, 520);
        if (owner != null && owner.getScene() != null) {
            scene.getStylesheets().setAll(owner.getScene().getStylesheets());
        }
        stage.setScene(scene);
        stage.showAndWait();
    }

    private static void addColumn(
            TableView<AladdinEntryDispatchPlanIdentityCheck.Diff> table,
            String title,
            double minWidth,
            java.util.function.Function<AladdinEntryDispatchPlanIdentityCheck.Diff, String> value) {
        TableColumn<AladdinEntryDispatchPlanIdentityCheck.Diff, String> col = new TableColumn<>(title);
        col.setMinWidth(minWidth);
        col.setCellValueFactory(
                c -> new ReadOnlyStringWrapper(value.apply(c.getValue())));
        table.getColumns().add(col);
    }

    static void copyCsv(List<AladdinEntryDispatchPlanIdentityCheck.Diff> diffs) {
        String csv = AladdinEntryIdentityCheckDiffTable.toCsv(diffs);
        if (csv.isBlank()) {
            return;
        }
        ClipboardContent content = new ClipboardContent();
        content.putString(csv);
        Clipboard.getSystemClipboard().setContent(content);
    }

    static void copyRich(List<AladdinEntryDispatchPlanIdentityCheck.Diff> diffs) {
        String tsv = AladdinEntryIdentityCheckDiffTable.toTsv(diffs);
        String html = AladdinEntryIdentityCheckDiffTable.toHtmlTable(diffs);
        if (tsv.isBlank() || html.isBlank()) {
            return;
        }
        ClipboardTableSupport.copyTabularForRichTextPaste(tsv, html);
    }

    private static String nz(String s) {
        return s != null ? s : "";
    }
}
