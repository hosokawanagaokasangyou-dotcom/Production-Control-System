package jp.co.pm.ai.desktop.io.actuals;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * 加工トレンド Excel 用の簡易折れ線チャート生成（POI XDDF）。
 */
final class ProcessingTrendWorkbookCharts {

    private ProcessingTrendWorkbookCharts() {}

    /**
     * カテゴリ列と複数の数値系列から折れ線チャートを追加する。
     *
     * @param categoryCol カテゴリ（日付・年月）列 index（0-based）
     * @param seriesCols 数値系列の列 index
     * @param seriesTitles 系列名（seriesCols と同長）
     * @param firstDataRow データ開始行（ヘッダの次）
     * @param lastDataRow データ終了行（合計行の直前）
     */
    static void addLineChart(
            XSSFSheet sheet,
            String title,
            int anchorCol1,
            int anchorRow1,
            int anchorCol2,
            int anchorRow2,
            int categoryCol,
            int[] seriesCols,
            String[] seriesTitles,
            int firstDataRow,
            int lastDataRow) {
        if (sheet == null || seriesCols == null || seriesTitles == null) {
            return;
        }
        if (seriesCols.length == 0 || seriesCols.length != seriesTitles.length) {
            return;
        }
        if (lastDataRow < firstDataRow) {
            return;
        }

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor =
                drawing.createAnchor(
                        0,
                        0,
                        0,
                        0,
                        anchorCol1,
                        anchorRow1,
                        anchorCol2,
                        anchorRow2);
        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(title);
        chart.setTitleOverlay(false);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);

        XDDFCategoryAxis bottom = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottom.setTitle("");
        XDDFValueAxis left = chart.createValueAxis(AxisPosition.LEFT);
        left.setTitle("加工長 (m)");
        left.setCrosses(AxisCrosses.AUTO_ZERO);

        XDDFCategoryDataSource cats =
                XDDFDataSourcesFactory.fromStringCellRange(
                        sheet, new CellRangeAddress(firstDataRow, lastDataRow, categoryCol, categoryCol));

        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottom, left);
        for (int i = 0; i < seriesCols.length; i++) {
            int col = seriesCols[i];
            XDDFNumericalDataSource<Double> vals =
                    XDDFDataSourcesFactory.fromNumericCellRange(
                            sheet, new CellRangeAddress(firstDataRow, lastDataRow, col, col));
            XDDFChartData.Series series = data.addSeries(cats, vals);
            series.setTitle(seriesTitles[i], null);
            if (series instanceof XDDFLineChartData.Series lineSeries) {
                lineSeries.setSmooth(false);
                lineSeries.setMarkerStyle(MarkerStyle.CIRCLE);
            }
        }
        chart.plot(data);
    }
}
