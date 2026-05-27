package jp.co.pm.ai.desktop.dispatch.rules.ui.editor;

import java.util.List;
import java.util.function.Consumer;

import javafx.scene.canvas.Canvas;
import javafx.scene.canvas.GraphicsContext;
import javafx.scene.input.MouseButton;
import javafx.scene.layout.Pane;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.TextAlignment;

import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleEdge;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleGraph;
import jp.co.pm.ai.desktop.dispatch.rules.model.DispatchRuleNode;

/** Canvas graph editor with category colors. */
public final class DispatchRuleGraphEditorPane extends Pane {

    private static final double NODE_W = 140;
    private static final double NODE_H = 56;

    private final Canvas canvas = new Canvas(800, 480);
    private DispatchRuleGraph graph = new DispatchRuleGraph();
    private String highlightedNodeId = "";
    private double wipOverlayValue = -1;
    private double wipOverlayThreshold = 20;
    private Consumer<String> nodeSelectionHandler = id -> {};

    public DispatchRuleGraphEditorPane() {
        getChildren().add(canvas);
        canvas.widthProperty().bind(widthProperty());
        canvas.heightProperty().bind(heightProperty());
        canvas.setOnMouseClicked(
                e -> {
                    if (e.getButton() != MouseButton.PRIMARY) {
                        return;
                    }
                    for (DispatchRuleNode node : graph.nodes) {
                        if (hit(node, e.getX(), e.getY())) {
                            nodeSelectionHandler.accept(node.id);
                            redraw();
                            return;
                        }
                    }
                });
        widthProperty().addListener((o, a, b) -> redraw());
        heightProperty().addListener((o, a, b) -> redraw());
    }

    public void setGraph(DispatchRuleGraph graph) {
        this.graph = graph != null ? graph : new DispatchRuleGraph();
        redraw();
    }

    public void setHighlightedNodeId(String nodeId) {
        this.highlightedNodeId = nodeId != null ? nodeId : "";
        redraw();
    }

    public void setWipOverlay(double wip, double threshold) {
        this.wipOverlayValue = wip;
        if (threshold > 0) {
            this.wipOverlayThreshold = threshold;
        }
        redraw();
    }

    public void clearWipOverlay() {
        this.wipOverlayValue = -1;
        redraw();
    }

    public void setOnNodeSelected(Consumer<String> handler) {
        this.nodeSelectionHandler = handler != null ? handler : id -> {};
    }

    private boolean hit(DispatchRuleNode node, double x, double y) {
        return x >= node.x && x <= node.x + NODE_W && y >= node.y && y <= node.y + NODE_H;
    }

    public void redraw() {
        double w = canvas.getWidth();
        double h = canvas.getHeight();
        if (w <= 0 || h <= 0) {
            return;
        }
        GraphicsContext g = canvas.getGraphicsContext2D();
        g.setFill(Color.web("#f5f5f5"));
        g.fillRect(0, 0, w, h);
        drawGrid(g, w, h);
        for (DispatchRuleEdge edge : graph.edges) {
            DispatchRuleNode from = findNode(edge.from);
            DispatchRuleNode to = findNode(edge.to);
            if (from == null || to == null) {
                continue;
            }
            g.setStroke(Color.web("#666"));
            g.setLineWidth(2);
            double x1 = from.x + NODE_W;
            double y1 = from.y + NODE_H / 2;
            double x2 = to.x;
            double y2 = to.y + NODE_H / 2;
            g.strokeLine(x1, y1, x2, y2);
        }
        for (DispatchRuleNode node : graph.nodes) {
            drawNode(g, node);
        }
    }

    private void drawGrid(GraphicsContext g, double w, double h) {
        g.setStroke(Color.web("#ddd"));
        g.setLineWidth(0.5);
        for (double x = 0; x < w; x += 20) {
            g.strokeLine(x, 0, x, h);
        }
        for (double y = 0; y < h; y += 20) {
            g.strokeLine(0, y, w, y);
        }
    }

    private void drawNode(GraphicsContext g, DispatchRuleNode node) {
        Color band = categoryColor(node.type);
        g.setFill(Color.WHITE);
        g.fillRoundRect(node.x, node.y, NODE_W, NODE_H, 8, 8);
        g.setFill(band);
        g.fillRoundRect(node.x, node.y, 8, NODE_H, 8, 8);
        g.fillRect(node.x + 4, node.y, 4, NODE_H);
        if (node.id.equals(highlightedNodeId)) {
            g.setStroke(Color.web("#E74C3C"));
            g.setLineWidth(3);
        } else {
            g.setStroke(Color.web("#333"));
            g.setLineWidth(1);
        }
        g.strokeRoundRect(node.x, node.y, NODE_W, NODE_H, 8, 8);
        g.setFill(Color.web("#222"));
        g.setFont(Font.font(11));
        g.setTextAlign(TextAlignment.CENTER);
        String text = node.label != null && !node.label.isBlank() ? node.label : node.type;
        g.fillText(truncate(text, 16), node.x + NODE_W / 2, node.y + NODE_H / 2 + 4);
        if (wipOverlayValue >= 0
                && node.type != null
                && (node.type.startsWith("metric.wip") || node.type.startsWith("compare.threshold"))) {
            g.setFill(
                    wipOverlayValue >= wipOverlayThreshold
                            ? Color.web("#C0392B")
                            : Color.web("#2980B9"));
            g.setFont(Font.font(12));
            g.fillText(
                    String.format("WIP %.0f", wipOverlayValue),
                    node.x + NODE_W / 2,
                    node.y + NODE_H - 8);
        }
    }

    private static String truncate(String s, int max) {
        if (s.length() <= max) {
            return s;
        }
        return s.substring(0, max - 1) + "…";
    }

    private DispatchRuleNode findNode(String id) {
        for (DispatchRuleNode n : graph.nodes) {
            if (id.equals(n.id)) {
                return n;
            }
        }
        return null;
    }

    public static Color categoryColor(String type) {
        if (type == null) {
            return Color.web("#95A5A6");
        }
        if (type.startsWith("scope.")) {
            return Color.web("#4A90D9");
        }
        if (type.startsWith("filter.")) {
            return Color.web("#9B59B6");
        }
        if (type.startsWith("metric.")) {
            return Color.web("#E67E22");
        }
        if (type.startsWith("compare.")) {
            return Color.web("#F1C40F");
        }
        if (type.startsWith("action.")) {
            return Color.web("#E74C3C");
        }
        return Color.web("#95A5A6");
    }

    public List<DispatchRuleNode> nodes() {
        return graph.nodes;
    }
}
