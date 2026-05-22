package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Optional;

import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Separator;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.CheckBoxTableCell;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.CodeDispatchLookupTablesBlankPrompt.ProductInput;
import jp.co.pm.ai.desktop.CodeDispatchLookupTablesBlankPrompt.ProductPromptRow;
import jp.co.pm.ai.desktop.CodeDispatchLookupTablesBlankPrompt.PromptBundle;
import jp.co.pm.ai.desktop.CodeDispatchLookupTablesBlankPrompt.UsedRawInput;
import jp.co.pm.ai.desktop.CodeDispatchLookupTablesBlankPrompt.UsedRawPromptRow;

/**
 * 段階1完了後: 新規の製品名・使用原反について材料テーブル値を入力し、行ごとに確認チェックを付けて OK する。
 */
public final class Stage1NewMaterialLookupDialog {

    private Stage1NewMaterialLookupDialog() {}

    public record Result(List<ProductInput> products, List<UsedRawInput> usedRaws) {}

    public static final class ProductRow {
        private final String productName;
        private final boolean needRollLength;
        private final boolean needWidth;
        private final boolean needThickness;
        private final boolean needLength;
        private final StringProperty rollLength = new SimpleStringProperty();
        private final StringProperty productWidth = new SimpleStringProperty();
        private final StringProperty thickness = new SimpleStringProperty();
        private final StringProperty productLength = new SimpleStringProperty();
        private final BooleanProperty confirmed = new SimpleBooleanProperty(false);

        ProductRow(ProductPromptRow src) {
            this.productName = src.productName();
            this.needRollLength = src.needRollLength();
            this.needWidth = src.needWidth();
            this.needThickness = src.needThickness();
            this.needLength = src.needLength();
            rollLength.set(src.suggestedRollLength());
            productWidth.set(src.suggestedWidth());
            thickness.set(src.suggestedThickness());
            productLength.set(src.suggestedLength());
        }

        public String productName() {
            return productName;
        }

        boolean needRollLength() {
            return needRollLength;
        }

        boolean needWidth() {
            return needWidth;
        }

        boolean needThickness() {
            return needThickness;
        }

        boolean needLength() {
            return needLength;
        }

        public StringProperty rollLengthProperty() {
            return rollLength;
        }

        public StringProperty productWidthProperty() {
            return productWidth;
        }

        public StringProperty thicknessProperty() {
            return thickness;
        }

        public StringProperty productLengthProperty() {
            return productLength;
        }

        public BooleanProperty confirmedProperty() {
            return confirmed;
        }

        ProductInput toInput() {
            return new ProductInput(
                    productName,
                    needRollLength ? rollLength.get() : "",
                    needWidth ? productWidth.get() : "",
                    needThickness ? thickness.get() : "",
                    needLength ? productLength.get() : "");
        }
    }

    public static final class UsedRawRow {
        private final String usedRaw;
        private final boolean needRollLength;
        private final boolean needRawWidth;
        private final StringProperty rollLength = new SimpleStringProperty();
        private final StringProperty rawWidth = new SimpleStringProperty();
        private final BooleanProperty confirmed = new SimpleBooleanProperty(false);

        UsedRawRow(UsedRawPromptRow src) {
            this.usedRaw = src.usedRaw();
            this.needRollLength = src.needRollLength();
            this.needRawWidth = src.needRawWidth();
            rollLength.set(src.suggestedRollLength());
            rawWidth.set(src.suggestedRawWidth());
        }

        public String usedRaw() {
            return usedRaw;
        }

        boolean needRollLength() {
            return needRollLength;
        }

        boolean needRawWidth() {
            return needRawWidth;
        }

        public StringProperty rollLengthProperty() {
            return rollLength;
        }

        public StringProperty rawWidthProperty() {
            return rawWidth;
        }

        public BooleanProperty confirmedProperty() {
            return confirmed;
        }

        UsedRawInput toInput() {
            return new UsedRawInput(
                    usedRaw,
                    needRollLength ? rollLength.get() : "",
                    needRawWidth ? rawWidth.get() : "");
        }
    }

    public static Optional<Result> prompt(Window owner, PromptBundle bundle) {
        if (bundle == null || bundle.empty()) {
            return Optional.empty();
        }

        List<ProductRow> productRows = new ArrayList<>();
        if (bundle.products() != null) {
            for (ProductPromptRow p : bundle.products()) {
                productRows.add(new ProductRow(p));
            }
        }
        List<UsedRawRow> usedRawRows = new ArrayList<>();
        if (bundle.usedRaws() != null) {
            for (UsedRawPromptRow u : bundle.usedRaws()) {
                usedRawRows.add(new UsedRawRow(u));
            }
        }

        Dialog<ButtonType> dialog = new Dialog<>();
        dialog.initOwner(owner);
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.setTitle("段階1 — 新規材料・製品種類の登録");
        dialog.setHeaderText(
                "加工計画DATA に未登録の製品名・使用原反が見つかりました。"
                        + " 各項目を入力し、行末の「確認」にチェックを付けて OK してください。");

        Label hint =
                new Label(
                        "空欄のまま段階2・段階3は実行できません。"
                                + " 推定値が入っている列は必要に応じて修正してください。");
        hint.setWrapText(true);
        hint.setStyle("-fx-font-size: 11px; -fx-text-fill: derive(-fx-text-inner-color, 22%);");

        VBox content = new VBox(10);
        content.setPadding(new Insets(4, 0, 0, 0));

        if (!productRows.isEmpty()) {
            content.getChildren().add(new Label("新規 製品名（" + productRows.size() + " 件）"));
            content.getChildren().add(buildProductTable(productRows));
        }
        if (!productRows.isEmpty() && !usedRawRows.isEmpty()) {
            content.getChildren().add(new Separator());
        }
        if (!usedRawRows.isEmpty()) {
            content.getChildren().add(new Label("新規 使用原反（" + usedRawRows.size() + " 件）"));
            content.getChildren().add(buildUsedRawTable(usedRawRows));
        }

        ScrollPane scroll = new ScrollPane(content);
        scroll.setFitToWidth(true);
        scroll.setPrefViewportHeight(Math.min(520, 120 + (productRows.size() + usedRawRows.size()) * 30.0));
        VBox root = new VBox(10, hint, scroll);
        VBox.setVgrow(scroll, Priority.ALWAYS);
        dialog.getDialogPane().setContent(root);
        dialog.getDialogPane().setPrefWidth(920);
        dialog.getDialogPane().getButtonTypes().setAll(ButtonType.OK, ButtonType.CANCEL);
        dialog.getDialogPane()
                .lookupButton(ButtonType.OK)
                .addEventFilter(
                        javafx.event.ActionEvent.ACTION,
                        ev -> {
                            commitPendingEdits(content);
                            Optional<String> err = validate(productRows, usedRawRows);
                            if (err.isPresent()) {
                                ev.consume();
                                showValidationError(dialog, err.get());
                            }
                        });

        Optional<ButtonType> choice = dialog.showAndWait();
        if (choice.isEmpty() || choice.get() != ButtonType.OK) {
            return Optional.empty();
        }
        List<ProductInput> products = new ArrayList<>(productRows.size());
        for (ProductRow r : productRows) {
            products.add(r.toInput());
        }
        List<UsedRawInput> usedRaws = new ArrayList<>(usedRawRows.size());
        for (UsedRawRow r : usedRawRows) {
            usedRaws.add(r.toInput());
        }
        return Optional.of(new Result(products, usedRaws));
    }

    private static TableView<ProductRow> buildProductTable(List<ProductRow> rows) {
        TableView<ProductRow> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        table.setEditable(true);
        table.setPrefHeight(Math.min(280, 56 + rows.size() * 28.0));

        TableColumn<ProductRow, String> cName = new TableColumn<>("製品名");
        cName.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().productName()));
        cName.setEditable(false);
        cName.setPrefWidth(220);

        TableColumn<ProductRow, String> cRoll = productNumColumn("ロール長(m)", ProductRow::rollLengthProperty);
        cRoll.setCellFactory(col -> conditionalProductTextCell(ProductRow::needRollLength));

        TableColumn<ProductRow, String> cWidth = productNumColumn("製品幅(mm)", ProductRow::productWidthProperty);
        cWidth.setCellFactory(col -> conditionalProductTextCell(ProductRow::needWidth));

        TableColumn<ProductRow, String> cThick = productNumColumn("厚み(mm)", ProductRow::thicknessProperty);
        cThick.setCellFactory(col -> conditionalProductTextCell(ProductRow::needThickness));

        TableColumn<ProductRow, String> cLen = productNumColumn("製品長(mm)", ProductRow::productLengthProperty);
        cLen.setCellFactory(col -> conditionalProductTextCell(ProductRow::needLength));

        TableColumn<ProductRow, Boolean> cConfirm = new TableColumn<>("確認");
        cConfirm.setCellValueFactory(cd -> cd.getValue().confirmedProperty());
        cConfirm.setCellFactory(CheckBoxTableCell.forTableColumn(cConfirm));
        cConfirm.setEditable(true);
        cConfirm.setPrefWidth(56);

        table.getColumns().setAll(cName, cRoll, cWidth, cThick, cLen, cConfirm);
        return table;
    }

    private static TableView<UsedRawRow> buildUsedRawTable(List<UsedRawRow> rows) {
        TableView<UsedRawRow> table = new TableView<>(FXCollections.observableArrayList(rows));
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        table.setEditable(true);
        table.setPrefHeight(Math.min(220, 56 + rows.size() * 28.0));

        TableColumn<UsedRawRow, String> cName = new TableColumn<>("使用原反");
        cName.setCellValueFactory(cd -> new SimpleStringProperty(cd.getValue().usedRaw()));
        cName.setEditable(false);
        cName.setPrefWidth(260);

        TableColumn<UsedRawRow, String> cRoll = usedRawNumColumn("ロール長(m)", UsedRawRow::rollLengthProperty);
        cRoll.setCellFactory(col -> conditionalUsedRawTextCell(UsedRawRow::needRollLength));

        TableColumn<UsedRawRow, String> cWidth = usedRawNumColumn("原反幅(mm)", UsedRawRow::rawWidthProperty);
        cWidth.setCellFactory(col -> conditionalUsedRawTextCell(UsedRawRow::needRawWidth));

        TableColumn<UsedRawRow, Boolean> cConfirm = new TableColumn<>("確認");
        cConfirm.setCellValueFactory(cd -> cd.getValue().confirmedProperty());
        cConfirm.setCellFactory(CheckBoxTableCell.forTableColumn(cConfirm));
        cConfirm.setEditable(true);
        cConfirm.setPrefWidth(56);

        table.getColumns().setAll(cName, cRoll, cWidth, cConfirm);
        return table;
    }

    private interface ProductField {
        StringProperty get(ProductRow row);
    }

    private interface UsedRawField {
        StringProperty get(UsedRawRow row);
    }

    private interface ProductNeed {
        boolean test(ProductRow row);
    }

    private interface UsedRawNeed {
        boolean test(UsedRawRow row);
    }

    private static TableColumn<ProductRow, String> productNumColumn(String title, ProductField field) {
        TableColumn<ProductRow, String> col = new TableColumn<>(title);
        col.setCellValueFactory(cd -> field.get(cd.getValue()));
        col.setCellFactory(TextFieldTableCell.forTableColumn());
        col.setOnEditCommit(
                ev -> {
                    if (ev.getNewValue() != null) {
                        field.get(ev.getRowValue()).set(ev.getNewValue());
                    }
                });
        col.setEditable(true);
        return col;
    }

    private static TableColumn<UsedRawRow, String> usedRawNumColumn(String title, UsedRawField field) {
        TableColumn<UsedRawRow, String> col = new TableColumn<>(title);
        col.setCellValueFactory(cd -> field.get(cd.getValue()));
        col.setCellFactory(TextFieldTableCell.forTableColumn());
        col.setOnEditCommit(
                ev -> {
                    if (ev.getNewValue() != null) {
                        field.get(ev.getRowValue()).set(ev.getNewValue());
                    }
                });
        col.setEditable(true);
        return col;
    }

    private static TextFieldTableCell<ProductRow, String> conditionalProductTextCell(ProductNeed needed) {
        return new TextFieldTableCell<>() {
            @Override
            public void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || getTableRow() == null || getTableRow().getItem() == null) {
                    setDisable(true);
                    setStyle("");
                    return;
                }
                ProductRow row = getTableRow().getItem();
                boolean active = needed.test(row);
                setDisable(!active);
                setStyle(active ? "" : "-fx-opacity: 0.45;");
            }
        };
    }

    private static TextFieldTableCell<UsedRawRow, String> conditionalUsedRawTextCell(UsedRawNeed needed) {
        return new TextFieldTableCell<>() {
            @Override
            public void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || getTableRow() == null || getTableRow().getItem() == null) {
                    setDisable(true);
                    setStyle("");
                    return;
                }
                UsedRawRow row = getTableRow().getItem();
                boolean active = needed.test(row);
                setDisable(!active);
                setStyle(active ? "" : "-fx-opacity: 0.45;");
            }
        };
    }

    private static void commitPendingEdits(VBox content) {
        for (var node : content.lookupAll(".table-view")) {
            if (node instanceof TableView<?> tv) {
                tv.edit(-1, null);
            }
        }
    }

    private static Optional<String> validate(List<ProductRow> products, List<UsedRawRow> usedRaws) {
        for (ProductRow r : products) {
            if (r.needRollLength && !isPositiveNumber(r.rollLengthProperty().get())) {
                return Optional.of("製品名「" + r.productName() + "」のロール長 (m) に正の数値を入力してください。");
            }
            if (r.needWidth && !isPositiveNumber(r.productWidthProperty().get())) {
                return Optional.of("製品名「" + r.productName() + "」の製品幅 (mm) に正の数値を入力してください。");
            }
            if (r.needThickness && !isPositiveNumber(r.thicknessProperty().get())) {
                return Optional.of("製品名「" + r.productName() + "」の厚み (mm) に正の数値を入力してください。");
            }
            if (r.needLength && !isPositiveNumber(r.productLengthProperty().get())) {
                return Optional.of("製品名「" + r.productName() + "」の製品長 (mm) に正の数値を入力してください。");
            }
            if (!r.confirmedProperty().get()) {
                return Optional.of("製品名「" + r.productName() + "」の行で「確認」にチェックを付けてください。");
            }
        }
        for (UsedRawRow r : usedRaws) {
            if (r.needRollLength && !isPositiveNumber(r.rollLengthProperty().get())) {
                return Optional.of("使用原反「" + r.usedRaw() + "」のロール長 (m) に正の数値を入力してください。");
            }
            if (r.needRawWidth && !isPositiveNumber(r.rawWidthProperty().get())) {
                return Optional.of("使用原反「" + r.usedRaw() + "」の原反幅 (mm) に正の数値を入力してください。");
            }
            if (!r.confirmedProperty().get()) {
                return Optional.of("使用原反「" + r.usedRaw() + "」の行で「確認」にチェックを付けてください。");
            }
        }
        return Optional.empty();
    }

    private static boolean isPositiveNumber(String raw) {
        if (raw == null || raw.isBlank()) {
            return false;
        }
        try {
            return Double.parseDouble(raw.strip().replace(",", "")) > 0;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private static void showValidationError(Dialog<?> parent, String message) {
        Dialog<Void> err = new Dialog<>();
        err.initOwner(parent.getDialogPane().getScene().getWindow());
        err.initModality(Modality.WINDOW_MODAL);
        err.setTitle("入力エラー");
        err.setHeaderText("材料・製品種類の入力を確認してください");
        err.setContentText(message);
        err.getDialogPane().getButtonTypes().setAll(ButtonType.OK);
        err.showAndWait();
    }
}
