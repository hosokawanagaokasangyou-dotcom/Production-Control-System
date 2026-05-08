package jp.co.pm.ai.desktop;

import javax.swing.SwingUtilities;

import javafx.embed.swing.SwingNode;
import javafx.fxml.FXML;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

/** Embeds sample JFreeChart in {@link SwingNode}; layout is {@code ChartTab.fxml}. */
public final class ChartTabController {

    @FXML
    private SwingNode chartSwingNode;

    @FXML
    private void initialize() {
        embedSampleChart();
    }

    private void embedSampleChart() {
        DefaultCategoryDataset ds = new DefaultCategoryDataset();
        ds.addValue(12, "actual", "M-A");
        ds.addValue(8, "actual", "M-B");
        ds.addValue(15, "plan", "M-A");
        SwingUtilities.invokeLater(
                () -> {
                    JFreeChart chart =
                            ChartFactory.createBarChart(
                                    "sample by equipment (JFreeChart)",
                                    "equipment",
                                    "qty",
                                    ds,
                                    PlotOrientation.VERTICAL,
                                    true,
                                    true,
                                    false);
                    ChartPanel panel = new ChartPanel(chart);
                    panel.setFillZoomRectangle(true);
                    javafx.application.Platform.runLater(() -> chartSwingNode.setContent(panel));
                });
    }
}
