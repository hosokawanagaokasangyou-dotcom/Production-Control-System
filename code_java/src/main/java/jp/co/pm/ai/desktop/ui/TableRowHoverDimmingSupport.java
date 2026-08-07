package jp.co.pm.ai.desktop.ui;

import java.lang.ref.WeakReference;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import javafx.application.Platform;
import javafx.scene.Node;
import javafx.scene.Parent;
import javafx.scene.control.TableRow;
import javafx.scene.control.TableView;
import javafx.util.Callback;

import org.controlsfx.control.spreadsheet.SpreadsheetView;

/**
 * {@link TableView} および ControlsFX {@link SpreadsheetView} 内表の行ホバー暗転。
 *
 * <p>SpreadsheetView は {@code setGrid} 後に内側 {@link TableView} の {@code rowFactory} を差し替えるため、
 * {@link #installOnSpreadsheet} と {@link #rehookSpreadsheet} で再ラップする。
 */
public final class TableRowHoverDimmingSupport {

  public static final String STYLE_ROW_FOCUSED = "pm-grid-row-hover-row-focused";
  public static final String STYLE_ROW_DIMMED = "pm-grid-row-hover-row-dimmed";

  private static final String HOOK_KEY = "jp.co.pm.ai.desktop.tableRowHoverDimmingHook";
  private static final String HOOK_EXIT_KEY = "jp.co.pm.ai.desktop.tableRowHoverDimmingExitHook";
  private static final String HOVERED_ROW_KEY = "jp.co.pm.ai.desktop.tableRowHoveredIndex";

  private static final List<WeakReference<TableView<?>>> REGISTERED_TABLES = new ArrayList<>();

  private TableRowHoverDimmingSupport() {}

  public static void install(TableView<?> table) {
    Objects.requireNonNull(table, "table");
    registerTable(table);
    wrapRowFactory(table);
    installTableExitHandler(table);
  }

  public static void installOnSpreadsheet(SpreadsheetView view) {
    if (view == null) {
      return;
    }
    if (view.getProperties().putIfAbsent(HOOK_KEY + ".spreadsheet", Boolean.TRUE) == null) {
      view.sceneProperty()
          .addListener(
              (obs, oldScene, newScene) -> {
                if (newScene != null) {
                  scheduleSpreadsheetRehook(view);
                }
              });
      view.skinProperty()
          .addListener((obs, oldSkin, newSkin) -> scheduleSpreadsheetRehook(view));
      view.gridProperty()
          .addListener((obs, oldGrid, newGrid) -> scheduleSpreadsheetRehook(view));
    }
    scheduleSpreadsheetRehook(view);
  }

  /** {@link SpreadsheetView#setGrid} やスキン再構築後に内側 {@link TableView} へ再適用する。 */
  public static void rehookSpreadsheet(SpreadsheetView view) {
    if (view == null) {
      return;
    }
    List<TableView<?>> tables = new ArrayList<>();
    for (Node ch : view.getChildrenUnmodifiable()) {
      if (ch instanceof TableView<?> tv) {
        tables.add(tv);
      }
    }
    collectEmbeddedTableViews(view, 0, tables);
    for (TableView<?> tv : tables) {
      install(tv);
    }
  }

  public static void scheduleSpreadsheetRehook(SpreadsheetView view) {
    if (view == null) {
      return;
    }
    Platform.runLater(
        () -> {
          rehookSpreadsheet(view);
          Platform.runLater(() -> rehookSpreadsheet(view));
        });
  }

  public static void refreshAllRegistered() {
    REGISTERED_TABLES.removeIf(ref -> ref.get() == null);
    for (WeakReference<TableView<?>> ref : REGISTERED_TABLES) {
      TableView<?> table = ref.get();
      if (table == null) {
        continue;
      }
      if (!UiRowHoverDimmingSettings.enabled()) {
        table.getProperties().remove(HOVERED_ROW_KEY);
      }
      refreshRows(table);
    }
  }

  private static void registerTable(TableView<?> table) {
    for (WeakReference<TableView<?>> ref : REGISTERED_TABLES) {
      if (ref.get() == table) {
        return;
      }
    }
    REGISTERED_TABLES.add(new WeakReference<>(table));
  }

  @SuppressWarnings({"rawtypes", "unchecked"})
  private static void wrapRowFactory(TableView<?> table) {
    Callback current = table.getRowFactory();
    if (current instanceof DimmingRowFactoryWrapper) {
      return;
    }
    table.setRowFactory(new DimmingRowFactoryWrapper(current));
  }

  private static void installTableExitHandler(TableView<?> table) {
    if (table.getProperties().putIfAbsent(HOOK_EXIT_KEY, Boolean.TRUE) != null) {
      return;
    }
    table.setOnMouseExited(
        e -> {
          if (!table.isHover()) {
            table.getProperties().remove(HOVERED_ROW_KEY);
            refreshRows(table);
          }
        });
  }

  private static void attachRowHoverHandlers(TableView<?> table, TableRow<?> row) {
    row.setOnMouseEntered(
        e -> {
          if (!UiRowHoverDimmingSettings.enabled()) {
            return;
          }
          int idx = row.getIndex();
          if (idx >= 0) {
            table.getProperties().put(HOVERED_ROW_KEY, idx);
            refreshRows(table);
          }
        });
  }

  private static void collectEmbeddedTableViews(Node n, int depth, List<TableView<?>> out) {
    if (n == null || depth > 32) {
      return;
    }
    if (n instanceof TableView<?> tv) {
      if (!out.contains(tv)) {
        out.add(tv);
      }
    }
    if (n instanceof Parent p) {
      for (Node c : p.getChildrenUnmodifiable()) {
        collectEmbeddedTableViews(c, depth + 1, out);
      }
    }
  }

  private static void refreshRows(TableView<?> table) {
    Object hoveredObj = table.getProperties().get(HOVERED_ROW_KEY);
    int hovered =
        hoveredObj instanceof Integer i && UiRowHoverDimmingSettings.enabled() ? i : -1;
    for (Node n : table.lookupAll("TableRow")) {
      if (n instanceof TableRow<?> row) {
        int idx = row.getIndex();
        if (idx < 0) {
          continue;
        }
        boolean focused = hovered >= 0 && idx == hovered;
        boolean dim = hovered >= 0 && !focused;
        toggleStyleClass(row, STYLE_ROW_FOCUSED, focused);
        toggleStyleClass(row, STYLE_ROW_DIMMED, dim);
      }
    }
  }

  private static void toggleStyleClass(Node node, String styleClass, boolean add) {
    if (add) {
      if (!node.getStyleClass().contains(styleClass)) {
        node.getStyleClass().add(styleClass);
      }
    } else {
      node.getStyleClass().remove(styleClass);
    }
  }

  @SuppressWarnings("rawtypes")
  private static final class DimmingRowFactoryWrapper implements Callback {

    private final Callback delegate;

    private DimmingRowFactoryWrapper(Callback delegate) {
      this.delegate = delegate;
    }

    @Override
    public Object call(Object param) {
      TableView table = (TableView) param;
      TableRow row = delegate != null ? (TableRow) delegate.call(param) : new TableRow();
      attachRowHoverHandlers(table, row);
      return row;
    }
  }
}
