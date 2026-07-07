package jp.co.pm.ai.desktop.ui;

import java.util.ArrayList;
import java.util.List;

import javafx.application.Platform;
import javafx.scene.Node;
import javafx.scene.Parent;

import org.controlsfx.control.spreadsheet.SpreadsheetView;

import jp.co.pm.ai.desktop.dispatch.ResultDispatchRequestFormOriginalColumns;

/**
 * ControlsFX {@link SpreadsheetView} の横列見出し（{@code HorizontalHeaderColumn}）に、依頼書原本由来列の紫色を付ける。
 */
public final class SpreadsheetRequestFormOriginalHeaderStyle {

    private SpreadsheetRequestFormOriginalHeaderStyle() {}

    public static void applyWhenReady(SpreadsheetView view, List<String> headersInVisualOrder) {
        if (view == null || headersInVisualOrder == null || headersInVisualOrder.isEmpty()) {
            return;
        }
        List<String> headers = new ArrayList<>(headersInVisualOrder);
        Platform.runLater(() -> applyNow(view, headers));
    }

    private static void applyNow(SpreadsheetView view, List<String> headers) {
        if (view.getColumns().isEmpty()) {
            return;
        }
        List<Node> horizontal = new ArrayList<>();
        collectHorizontalColumnHeaders(view, horizontal);
        if (horizontal.isEmpty()) {
            return;
        }
        int n = Math.min(horizontal.size(), headers.size());
        String styleClass = ResultDispatchRequestFormOriginalColumns.HEADER_STYLE_CLASS;
        for (int i = 0; i < n; i++) {
            Node header = horizontal.get(i);
            boolean original =
                    ResultDispatchRequestFormOriginalColumns.isDerivedFromRequestFormOriginal(
                            headers.get(i));
            if (original) {
                if (!header.getStyleClass().contains(styleClass)) {
                    header.getStyleClass().add(styleClass);
                }
            } else {
                header.getStyleClass().remove(styleClass);
            }
        }
    }

    private static void collectHorizontalColumnHeaders(Node n, List<Node> out) {
        collectHorizontalColumnHeaders(n, out, 0);
    }

    private static void collectHorizontalColumnHeaders(Node n, List<Node> out, int depth) {
        if (n == null || depth > 32) {
            return;
        }
        if (isTableColumnHeaderNode(n) && isUnderHorizontalHeaderColumn(n)) {
            out.add(n);
            return;
        }
        if (n instanceof Parent p) {
            for (Node c : p.getChildrenUnmodifiable()) {
                collectHorizontalColumnHeaders(c, out, depth + 1);
            }
        }
    }

    private static boolean isTableColumnHeaderNode(Node n) {
        return n.getClass().getName().endsWith("TableColumnHeader");
    }

    private static boolean isUnderHorizontalHeaderColumn(Node n) {
        for (Node p = n; p != null; p = p.getParent()) {
            String cn = p.getClass().getName();
            if (cn.contains("HorizontalHeaderColumn")) {
                return true;
            }
            if (cn.contains("VerticalHeader")) {
                return false;
            }
        }
        return false;
    }
}
