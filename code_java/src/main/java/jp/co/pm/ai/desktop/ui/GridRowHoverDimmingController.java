package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;

import javafx.scene.Node;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.layout.Region;

/**
 * GridPane 上の行帯＋セルに、メンバー勤怠と同様の行ホバー暗転を適用する。
 */
public final class GridRowHoverDimmingController {

  public static final String STYLE_BAND = "pm-grid-row-hover-band";
  public static final String STYLE_BAND_FOCUSED = "pm-grid-row-hover-band-focused";
  public static final String STYLE_BAND_DIMMED = "pm-grid-row-hover-band-dimmed";
  public static final String STYLE_NODE_DIMMED = "pm-grid-row-hover-node-dimmed";

  private record Row(Region band, Label nameLabel, List<Node> nodes) {}

  private final List<Row> rows = new ArrayList<>();
  private int hoveredRow = -1;

  public void clear() {
    rows.clear();
    hoveredRow = -1;
  }

  public void addRow(Region band, Label nameLabel, List<Node> cellNodes) {
    rows.add(new Row(band, nameLabel, cellNodes != null ? cellNodes : List.of()));
  }

  public void installHover(Node node, int rowIndex) {
    if (node == null) {
      return;
    }
    node.setOnMouseEntered(e -> setHoveredRow(rowIndex));
  }

  public void installScrollClearOnExit(ScrollPane scroll) {
    if (scroll == null) {
      return;
    }
    scroll.setOnMouseExited(
        e -> {
          if (!scroll.isHover()) {
            setHoveredRow(-1);
          }
        });
  }

  public void setHoveredRow(int rowIndex) {
    if (!UiRowHoverDimmingSettings.enabled()) {
      if (hoveredRow == -1) {
        return;
      }
      hoveredRow = -1;
      applyStyles();
      return;
    }
    if (hoveredRow == rowIndex) {
      return;
    }
    hoveredRow = rowIndex;
    applyStyles();
  }

  public void refresh() {
    applyStyles();
  }

  private void applyStyles() {
    boolean en = UiRowHoverDimmingSettings.enabled();
    for (int i = 0; i < rows.size(); i++) {
      Row row = rows.get(i);
      boolean focused = en && i == hoveredRow;
      boolean dim = en && hoveredRow >= 0 && !focused;
      if (row.band() != null) {
        toggleStyleClass(row.band(), STYLE_BAND_FOCUSED, focused);
        toggleStyleClass(row.band(), STYLE_BAND_DIMMED, dim);
      }
      if (row.nameLabel() != null) {
        toggleStyleClass(row.nameLabel(), STYLE_NODE_DIMMED, dim);
      }
      for (Node n : row.nodes()) {
        toggleStyleClass(n, STYLE_NODE_DIMMED, dim);
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
}
