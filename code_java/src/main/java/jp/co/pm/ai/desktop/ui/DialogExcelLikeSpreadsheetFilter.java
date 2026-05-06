package jp.co.pm.ai.desktop.ui;

import java.util.BitSet;
import java.util.Comparator;
import java.util.HashSet;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import java.util.TreeSet;
import java.util.stream.Collectors;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.control.MenuButton;
import javafx.scene.control.MenuItem;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.Window;
import javafx.util.Callback;

import org.controlsfx.control.spreadsheet.Filter;
import org.controlsfx.control.spreadsheet.SpreadsheetCell;
import org.controlsfx.control.spreadsheet.SpreadsheetCellType;
import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * Column filter opening a modal stage with 適用 / OK / キャンセル. Row-hide semantics match
 * {@link ExcelLikeSpreadsheetFilter}; キャンセル restores hidden rows from when the dialog opened.
 */
public final class DialogExcelLikeSpreadsheetFilter implements Filter {

    private static final String SORT_ASC = "昇順で並べ替え";
    private static final String SORT_DESC = "降順で並べ替え";
    private static final String SORT_CLEAR = "並べ替えを解除";
    private static final String SEARCH_PROMPT = "値を検索…";
    private static final String SELECT_ALL = "すべて選択";
    private static final String CLEAR_ALL = "すべて解除";

    private final SpreadsheetView spv;
    private final int column;

    private MenuButton menuButton;

    private final Set<String> stringSet = new HashSet<>();
    private final Set<String> copySet = new HashSet<>();

    private MenuItem sortItem;
    private final Comparator<ObservableList<SpreadsheetCell>> ascendingComp;
    private final Comparator<ObservableList<SpreadsheetCell>> descendingComp;

    public DialogExcelLikeSpreadsheetFilter(SpreadsheetView spv, int column) {
        this.spv = Objects.requireNonNull(spv);
        this.column = column;
        this.ascendingComp =
                (o1, o2) -> compareRowsForSort(spv.getFilteredRow(), o1, o2, column, false);
        this.descendingComp =
                (o1, o2) -> compareRowsForSort(spv.getFilteredRow(), o1, o2, column, true);
    }

    /** After global clear, sync header filter menu sort labels with {@link SpreadsheetView#getComparator()}. */
    public static void resetAllColumnSortMenus(SpreadsheetView spv) {
        if (spv == null) {
            return;
        }
        for (var col : spv.getColumns()) {
            Filter f = col.getFilter();
            if (f instanceof DialogExcelLikeSpreadsheetFilter d) {
                d.resetSortMenuState();
            }
        }
    }

    private void resetSortMenuState() {
        if (sortItem != null) {
            sortItem.setText(SORT_ASC);
        }
    }

    private static int compareRowsForSort(
            int filteredRow,
            ObservableList<SpreadsheetCell> o1,
            ObservableList<SpreadsheetCell> o2,
            int columnIndex,
            boolean descending) {
        SpreadsheetCell cell1 = o1.get(columnIndex);
        SpreadsheetCell cell2 = o2.get(columnIndex);
        if (cell1.getRow() <= filteredRow) {
            return Integer.compare(cell1.getRow(), cell2.getRow());
        }
        if (cell2.getRow() <= filteredRow) {
            return Integer.compare(cell1.getRow(), cell2.getRow());
        }
        int cmp = compareCellValuesNumericAware(cell1, cell2);
        return descending ? -cmp : cmp;
    }

    private static int compareCellValuesNumericAware(SpreadsheetCell cell1, SpreadsheetCell cell2) {
        if (cell1.getCellType() == SpreadsheetCellType.INTEGER
                && cell2.getCellType() == SpreadsheetCellType.INTEGER) {
            return Integer.compare((Integer) cell1.getItem(), (Integer) cell2.getItem());
        }
        if (cell1.getCellType() == SpreadsheetCellType.DOUBLE
                && cell2.getCellType() == SpreadsheetCellType.DOUBLE) {
            return Double.compare((Double) cell1.getItem(), (Double) cell2.getItem());
        }
        Double n1 = numericSortKey(cell1);
        Double n2 = numericSortKey(cell2);
        if (n1 != null && n2 != null) {
            return Double.compare(n1, n2);
        }
        if (n1 != null) {
            return -1;
        }
        if (n2 != null) {
            return 1;
        }
        String t1 = cell1.getText() != null ? cell1.getText() : "";
        String t2 = cell2.getText() != null ? cell2.getText() : "";
        return t1.compareToIgnoreCase(t2);
    }

    private static Double numericSortKey(SpreadsheetCell c) {
        if (c.getCellType() == SpreadsheetCellType.INTEGER && c.getItem() instanceof Integer) {
            return ((Integer) c.getItem()).doubleValue();
        }
        if (c.getCellType() == SpreadsheetCellType.DOUBLE && c.getItem() instanceof Double) {
            return (Double) c.getItem();
        }
        return tryParseDouble(c.getText());
    }

    private static Double tryParseDouble(String raw) {
        if (raw == null) {
            return null;
        }
        String s = raw.trim();
        if (s.isEmpty()) {
            return null;
        }
        try {
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    @Override
    public MenuButton getMenuButton() {
        if (menuButton == null) {
            menuButton = new MenuButton();
            menuButton.getStyleClass().add("filter-menu-button");
            menuButton
                    .showingProperty()
                    .addListener(
                            (obs, oldVal, newVal) -> {
                                if (Boolean.TRUE.equals(newVal)) {
                                    menuButton.hide();
                                    javafx.application.Platform.runLater(this::openFilterDialog);
                                }
                            });
        }
        return menuButton;
    }

    private void openFilterDialog() {
        Window owner = spv.getScene() != null ? spv.getScene().getWindow() : null;

        rebuildUniqueValues();
        refreshCopySetFromVisibleRows();

        ObservableList<String> displayedItems = FXCollections.observableArrayList();

        TextField searchField = new TextField();
        searchField.setPromptText(SEARCH_PROMPT);
        HBox.setHgrow(searchField, Priority.ALWAYS);

        sortItem = new MenuItem(SORT_ASC);
        MenuButton sortMb = new MenuButton("並べ替え");
        sortMb.getItems().add(sortItem);
        sortItem.setOnAction(this::onSortAction);

        ListView<String> listView = new ListView<>(displayedItems);
        listView.setPrefHeight(260);
        listView.setCellFactory(
                new Callback<>() {
                    @Override
                    public ListCell<String> call(ListView<String> param) {
                        return new ListCell<>() {
                            @Override
                            protected void updateItem(String item, boolean empty) {
                                super.updateItem(item, empty);
                                if (empty || item == null) {
                                    setText(null);
                                    setGraphic(null);
                                    return;
                                }
                                setText(item);
                                CheckBox checkBox = new CheckBox();
                                checkBox.setSelected(copySet.contains(item));
                                checkBox
                                        .selectedProperty()
                                        .addListener(
                                                (obs, oldValue, newValue) -> {
                                                    if (Boolean.TRUE.equals(newValue)) {
                                                        copySet.add(item);
                                                    } else {
                                                        copySet.remove(item);
                                                    }
                                                });
                                setGraphic(checkBox);
                            }
                        };
                    }
                });

        Runnable refillDisplayed =
                () -> {
                    applySearchFilter(displayedItems, searchField.getText());
                    listView.refresh();
                };

        searchField
                .textProperty()
                .addListener((obs, o, n) -> refillDisplayed.run());

        Button selectAllBtn = new Button(SELECT_ALL);
        selectAllBtn.setMaxWidth(Double.MAX_VALUE);
        selectAllBtn.setOnAction(
                e -> {
                    copySet.addAll(displayedItems);
                    listView.refresh();
                });
        Button clearAllBtn = new Button(CLEAR_ALL);
        clearAllBtn.setMaxWidth(Double.MAX_VALUE);
        clearAllBtn.setOnAction(
                e -> {
                    copySet.clear();
                    listView.refresh();
                });
        HBox bulkRow = new HBox(8, selectAllBtn, clearAllBtn);

        refillDisplayed.run();

        BitSet snapshotHidden = new BitSet();
        snapshotHidden.or(spv.getHiddenRows());

        Runnable commit =
                () ->
                        SpreadsheetMultiColumnFilterCoordinator.commitColumnSelection(
                                spv, column, new HashSet<>(copySet));

        Runnable restore =
                () -> {
                    BitSet r = new BitSet();
                    r.or(snapshotHidden);
                    spv.setHiddenRows(r);
                };

        Button applyBtn = new Button("適用");
        applyBtn.setOnAction(
                e -> {
                    commit.run();
                    e.consume();
                });

        Button okBtn = new Button("OK");
        okBtn.setDefaultButton(true);
        okBtn.setOnAction(
                e -> {
                    commit.run();
                    ((Stage) okBtn.getScene().getWindow()).close();
                    e.consume();
                });

        Button cancelBtn = new Button("キャンセル");
        cancelBtn.setCancelButton(true);
        cancelBtn.setOnAction(
                e -> {
                    restore.run();
                    ((Stage) cancelBtn.getScene().getWindow()).close();
                    e.consume();
                });

        HBox buttonBar = new HBox(10, applyBtn, okBtn, cancelBtn);
        buttonBar.setAlignment(Pos.CENTER_RIGHT);
        buttonBar.setPadding(new Insets(8, 0, 0, 0));

        VBox root =
                new VBox(
                        8,
                        new Label("列フィルタ"),
                        sortMb,
                        searchField,
                        bulkRow,
                        listView,
                        buttonBar);
        root.setPadding(new Insets(12));
        VBox.setVgrow(listView, Priority.ALWAYS);

        Stage stage = new Stage();
        stage.initOwner(owner);
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle("列フィルタ");
        stage.setScene(new Scene(root, 420, 480));
        stage.showAndWait();
    }

    private void rebuildUniqueValues() {
        stringSet.clear();
        stringSet.addAll(
                SpreadsheetMultiColumnFilterCoordinator.distinctValuesForColumnRespectingOtherFilters(
                        spv, column));
    }

    private void refreshCopySetFromVisibleRows() {
        copySet.clear();
        int n = spv.getGrid().getRowCount();
        for (int i = spv.getFilteredRow() + 1; i < n; ++i) {
            if (!spv.getHiddenRows().get(i)) {
                copySet.add(spv.getGrid().getRows().get(i).get(column).getText());
            }
        }
    }

    private void applySearchFilter(ObservableList<String> displayedItems, String query) {
        String q = query != null ? query.trim().toLowerCase(Locale.ROOT) : "";
        TreeSet<String> sorted =
                new TreeSet<>(Comparator.nullsFirst(DialogExcelLikeSpreadsheetFilter::compareFilterChoiceStrings));
        sorted.addAll(stringSet);
        if (q.isEmpty()) {
            displayedItems.setAll(sorted);
        } else {
            displayedItems.setAll(
                    sorted.stream()
                            .filter(v -> v != null && v.toLowerCase(Locale.ROOT).contains(q))
                            .collect(Collectors.toList()));
        }
    }

    private static int compareFilterChoiceStrings(String a, String b) {
        if (a == null && b == null) {
            return 0;
        }
        if (a == null) {
            return -1;
        }
        if (b == null) {
            return 1;
        }
        Double da = tryParseDouble(a);
        Double db = tryParseDouble(b);
        if (da != null && db != null) {
            return Double.compare(da, db);
        }
        if (da != null) {
            return -1;
        }
        if (db != null) {
            return 1;
        }
        return a.compareToIgnoreCase(b);
    }

    private void onSortAction(ActionEvent event) {
        if (spv.getComparator() == ascendingComp) {
            spv.setComparator(descendingComp);
            sortItem.setText(SORT_CLEAR);
        } else if (spv.getComparator() == descendingComp) {
            spv.setComparator(null);
            sortItem.setText(SORT_ASC);
        } else {
            spv.setComparator(ascendingComp);
            sortItem.setText(SORT_DESC);
        }
    }
}
