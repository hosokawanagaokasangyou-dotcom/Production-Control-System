package jp.co.pm.ai.desktop.reconciliation;

import java.util.List;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.control.Label;
import javafx.scene.control.OverrunStyle;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.shape.Line;
import javafx.scene.text.Font;
import javafx.scene.text.FontPosture;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import javafx.scene.text.TextAlignment;
import javafx.scene.text.TextFlow;

final class RequestFormPreviewNodeFactory {

    private RequestFormPreviewNodeFactory() {}

    static Node buildCellNode(
            String text,
            RequestFormPreviewCellStyle style,
            List<RequestFormPreviewTextRun> richRuns,
            HorizontalAlignment hAlign,
            double width,
            double height) {
        Node content;
        if (richRuns != null && richRuns.size() > 1) {
            TextFlow flow = buildTextFlow(richRuns, hAlign, width, style.wrapText());
            content = wrapDecorations(flow, richRuns.get(0).style(), width, height);
        } else {
            Label label = new Label(text != null ? text : "");
            label.setWrapText(style != null && style.wrapText());
            label.setTextOverrun(OverrunStyle.CLIP);
            label.setMaxWidth(width);
            label.setMinWidth(width);
            label.setPrefWidth(width);
            if (style != null && style.wrapText()) {
                label.setMaxHeight(height);
            }
            applyLabelStyle(label, style);
            content = wrapDecorations(label, style, width, height);
        }
        return wrapCellContainer(content, style, hAlign, width, height);
    }

    static Node buildTextRuns(List<RequestFormPreviewTextRun> runs, double width, double height) {
        TextFlow flow = buildTextFlow(runs, HorizontalAlignment.CENTER, width, true);
        flow.setMaxWidth(width);
        flow.setPrefWidth(width);
        RequestFormPreviewCellStyle style =
                runs.isEmpty() ? RequestFormPreviewCellStyle.defaults() : runs.get(0).style();
        return wrapDecorations(flow, style, width, height);
    }

    private static TextFlow buildTextFlow(
            List<RequestFormPreviewTextRun> runs,
            HorizontalAlignment hAlign,
            double width,
            boolean wrapText) {
        TextFlow flow = new TextFlow();
        flow.setTextAlignment(toTextAlignment(hAlign));
        flow.setMaxWidth(width);
        flow.setPrefWidth(width);
        if (wrapText) {
            flow.setLineSpacing(0);
        }
        for (RequestFormPreviewTextRun run : runs) {
            Text text = new Text(run.text() != null ? run.text() : "");
            applyTextStyle(text, run.style());
            flow.getChildren().add(text);
        }
        return flow;
    }

    private static Node wrapCellContainer(
            Node content,
            RequestFormPreviewCellStyle style,
            HorizontalAlignment hAlign,
            double width,
            double height) {
        StackPane stack = new StackPane(content);
        stack.setMinSize(width, height);
        stack.setPrefSize(width, height);
        stack.setMaxSize(width, height);
        stack.setAlignment(toPos(hAlign, style != null ? style.verticalAlignment() : null));
        stack.setPadding(cellPadding(style));
        stack.setStyle(containerCss(style));
        return stack;
    }

    private static Insets cellPadding(RequestFormPreviewCellStyle style) {
        double top = 1.0;
        if (style != null && style.fontSizePx() > 0) {
            top = Math.max(1.0, Math.min(4.0, style.fontSizePx() * 0.08));
        }
        return new Insets(top, 3, 1, 3);
    }

    private static Node wrapDecorations(
            Node content, RequestFormPreviewCellStyle style, double width, double height) {
        if (style == null || (!style.doubleStrike() && !style.doubleUnderline())) {
            return content;
        }
        StackPane stack = new StackPane(content);
        stack.setMinSize(width, height);
        stack.setPrefSize(width, height);
        if (style.doubleStrike()) {
            for (Line line : doubleStrikeLines(width, height, style.foreground())) {
                stack.getChildren().add(line);
            }
        }
        if (style.doubleUnderline()) {
            for (Line line : doubleUnderlineLines(width, height, style.foreground())) {
                stack.getChildren().add(line);
            }
        }
        return stack;
    }

    private static Line[] doubleStrikeLines(double width, double height, String color) {
        Color stroke = color(color, Color.BLACK);
        Line upper = new Line(2, height * 0.38, width - 2, height * 0.38);
        Line lower = new Line(2, height * 0.52, width - 2, height * 0.52);
        upper.setStroke(stroke);
        lower.setStroke(stroke);
        upper.setStrokeWidth(1.0);
        lower.setStrokeWidth(1.0);
        return new Line[] {upper, lower};
    }

    private static Line[] doubleUnderlineLines(double width, double height, String color) {
        Color stroke = color(color, Color.BLACK);
        Line upper = new Line(2, height - 4, width - 2, height - 4);
        Line lower = new Line(2, height - 1, width - 2, height - 1);
        upper.setStroke(stroke);
        lower.setStroke(stroke);
        upper.setStrokeWidth(1.0);
        lower.setStrokeWidth(1.0);
        return new Line[] {upper, lower};
    }

    static void applyLabelStyle(Label label, RequestFormPreviewCellStyle style) {
        if (style == null) {
            return;
        }
        label.setStyle(textCss(style));
    }

    private static void applyTextStyle(Text text, RequestFormPreviewCellStyle style) {
        if (style == null) {
            return;
        }
        text.setFill(color(style.foreground(), Color.BLACK));
        text.setFont(toFont(style));
        text.setUnderline(style.underline() && !style.doubleUnderline());
        text.setStrikethrough(style.strike() && !style.doubleStrike());
    }

    static String containerCss(RequestFormPreviewCellStyle style) {
        if (style == null) {
            return "-fx-background-color: white;";
        }
        StringBuilder css = new StringBuilder();
        css.append("-fx-background-color: ")
                .append(style.background() != null ? style.background() : "#FFFFFF")
                .append(';');
        if (style.borderCss() != null && !style.borderCss().isBlank()) {
            css.append(style.borderCss());
        }
        return css.toString();
    }

    static String textCss(RequestFormPreviewCellStyle style) {
        return buildCss(style, false);
    }

    static String buildCss(RequestFormPreviewCellStyle style, boolean includeBackground) {
        StringBuilder css = new StringBuilder();
        if (includeBackground) {
            css.append("-fx-background-color: ")
                    .append(style.background() != null ? style.background() : "#FFFFFF")
                    .append(';');
            if (style.borderCss() != null && !style.borderCss().isBlank()) {
                css.append(style.borderCss());
            }
        }
        css.append("-fx-text-fill: ")
                .append(style.foreground() != null ? style.foreground() : "#000000")
                .append(';');
        css.append("-fx-font-size: ").append(Math.max(6.0, style.fontSizePx())).append("px;");
        if (style.bold()) {
            css.append("-fx-font-weight: bold;");
        }
        if (style.italic()) {
            css.append("-fx-font-style: italic;");
        }
        if (style.underline() && !style.doubleUnderline()) {
            css.append("-fx-underline: true;");
        }
        if (style.strike() && !style.doubleStrike()) {
            css.append("-fx-strikethrough: true;");
        }
        if (style.fontFamily() != null && !style.fontFamily().isBlank()) {
            css.append("-fx-font-family: '").append(style.fontFamily()).append("';");
        }
        return css.toString();
    }

    static Color color(String hex, Color fallback) {
        if (hex == null || hex.isBlank()) {
            return fallback;
        }
        try {
            return Color.web(hex);
        } catch (IllegalArgumentException ex) {
            return fallback;
        }
    }

    private static Font toFont(RequestFormPreviewCellStyle style) {
        FontWeight weight = style.bold() ? FontWeight.BOLD : FontWeight.NORMAL;
        FontPosture posture = style.italic() ? FontPosture.ITALIC : FontPosture.REGULAR;
        double size = Math.max(6.0, style.fontSizePx());
        if (style.fontFamily() != null && !style.fontFamily().isBlank()) {
            return Font.font(style.fontFamily(), weight, posture, size);
        }
        return Font.font(size);
    }

    private static Pos toPos(HorizontalAlignment hAlign, VerticalAlignment vAlign) {
        boolean top = vAlign == VerticalAlignment.TOP || vAlign == VerticalAlignment.JUSTIFY;
        boolean bottom = vAlign == VerticalAlignment.BOTTOM || vAlign == VerticalAlignment.DISTRIBUTED;
        boolean centerV = !top && !bottom;

        if (hAlign == HorizontalAlignment.RIGHT) {
            if (top) {
                return Pos.TOP_RIGHT;
            }
            if (bottom) {
                return Pos.BOTTOM_RIGHT;
            }
            return Pos.CENTER_RIGHT;
        }
        if (hAlign == HorizontalAlignment.CENTER
                || hAlign == HorizontalAlignment.CENTER_SELECTION
                || hAlign == HorizontalAlignment.FILL
                || hAlign == HorizontalAlignment.JUSTIFY
                || hAlign == HorizontalAlignment.DISTRIBUTED) {
            if (top) {
                return Pos.TOP_CENTER;
            }
            if (bottom) {
                return Pos.BOTTOM_CENTER;
            }
            return Pos.CENTER;
        }
        if (top) {
            return Pos.TOP_LEFT;
        }
        if (bottom) {
            return Pos.BOTTOM_LEFT;
        }
        return centerV ? Pos.CENTER_LEFT : Pos.CENTER_LEFT;
    }

    private static TextAlignment toTextAlignment(HorizontalAlignment alignment) {
        if (alignment == null) {
            return TextAlignment.LEFT;
        }
        return switch (alignment) {
            case CENTER, CENTER_SELECTION, FILL, JUSTIFY, DISTRIBUTED -> TextAlignment.CENTER;
            case RIGHT -> TextAlignment.RIGHT;
            default -> TextAlignment.LEFT;
        };
    }
}
