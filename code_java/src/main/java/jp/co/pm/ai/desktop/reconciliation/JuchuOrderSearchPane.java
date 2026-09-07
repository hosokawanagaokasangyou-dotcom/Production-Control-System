package jp.co.pm.ai.desktop.reconciliation;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Supplier;

import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Orientation;
import javafx.scene.Parent;
import javafx.scene.control.Button;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.SplitPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;

/**
 * 受注検索（閲覧専用）。左条件・右結果のレイアウト。
 * 編集フォームへのジャンプはしない。
 */
public final class JuchuOrderSearchPane {

    private JuchuOrderSearchPane() {}

    public static Parent build(Supplier<List<OrderRecord>> recordsSupplier) {
        Objects.requireNonNull(recordsSupplier, "recordsSupplier");

        DatePicker from = new DatePicker();
        DatePicker to = new DatePicker();
        TextField product = new TextField();
        product.setPromptText("製品名（部分一致）");
        TextField raw = new TextField();
        raw.setPromptText("投入原反（部分一致）");
        Button search = new Button("検索");
        Label statusMessage = new Label("");
        statusMessage.setWrapText(true);

        VBox conditions = new VBox(8);
        conditions.setPadding(new Insets(12));
        conditions.getChildren()
                .addAll(
                        labeled("納期 From", from),
                        labeled("納期 To", to),
                        labeled("製品", product),
                        labeled("投入原反", raw),
                        search,
                        statusMessage);

        ScrollPane leftScroll = new ScrollPane(conditions);
        leftScroll.setFitToWidth(true);
        leftScroll.getStyleClass().add("form-scroll-pane");

        Label countLabel = new Label("0 件");
        ObservableList<OrderRecord> items = FXCollections.observableArrayList();
        TableView<OrderRecord> table = new TableView<>(items);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        table.setPlaceholder(new Label("条件を指定して検索してください"));
        table.getColumns()
                .addAll(
                        col("依頼No", r -> nullToEmpty(r.getReqNo())),
                        col("希望納期", r -> dbValue(r, "希望納期")),
                        col("調整納期", r -> dbValue(r, "調整納期")),
                        col("製品", r -> dbValue(r, "製品")),
                        col("原反", r -> JuchuOrderSearch.displayRawMaterial(r.getDbValues())),
                        col("ユーザー", r -> nullToEmpty(r.getUser())),
                        col("入力日", r -> dbValue(r, "入力日")));

        VBox right = new VBox(8, countLabel, table);
        right.setPadding(new Insets(12));
        VBox.setVgrow(table, Priority.ALWAYS);

        search.setOnAction(
                e -> {
                    var c =
                            new JuchuOrderSearchCriteria(
                                    from.getValue(),
                                    to.getValue(),
                                    product.getText(),
                                    raw.getText());
                    Optional<String> err = c.validationError();
                    if (err.isPresent()) {
                        statusMessage.setText(err.get());
                        items.clear();
                        countLabel.setText("0 件");
                        return;
                    }
                    List<OrderRecord> hits = JuchuOrderSearch.filter(recordsSupplier.get(), c);
                    items.setAll(hits);
                    String countText = hits.size() + " 件";
                    statusMessage.setText(countText);
                    countLabel.setText(countText);
                });

        SplitPane split = new SplitPane(leftScroll, right);
        split.setOrientation(Orientation.HORIZONTAL);
        split.setDividerPositions(0.32);
        SplitPane.setResizableWithParent(leftScroll, Boolean.TRUE);
        SplitPane.setResizableWithParent(right, Boolean.TRUE);

        VBox root = new VBox(split);
        VBox.setVgrow(split, Priority.ALWAYS);
        root.setMaxWidth(Double.MAX_VALUE);
        root.setMaxHeight(Double.MAX_VALUE);
        return root;
    }

    private static VBox labeled(String caption, javafx.scene.Node field) {
        Label label = new Label(caption);
        VBox box = new VBox(4, label, field);
        HBox.setHgrow(field, Priority.ALWAYS);
        if (field instanceof TextField tf) {
            tf.setMaxWidth(Double.MAX_VALUE);
        } else if (field instanceof DatePicker dp) {
            dp.setMaxWidth(Double.MAX_VALUE);
        }
        return box;
    }

    private static TableColumn<OrderRecord, String> col(
            String title, java.util.function.Function<OrderRecord, String> value) {
        TableColumn<OrderRecord, String> column = new TableColumn<>(title);
        column.setCellValueFactory(
                cd -> {
                    OrderRecord row = cd.getValue();
                    return new SimpleStringProperty(row != null ? value.apply(row) : "");
                });
        return column;
    }

    private static String dbValue(OrderRecord record, String key) {
        Map<String, String> db = record.getDbValues();
        if (db == null) {
            return "";
        }
        return nullToEmpty(db.get(key));
    }

    private static String nullToEmpty(String value) {
        return value != null ? value : "";
    }
}
