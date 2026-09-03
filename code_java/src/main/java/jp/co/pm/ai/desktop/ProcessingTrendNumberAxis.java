package jp.co.pm.ai.desktop;

import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import javafx.beans.property.DoubleProperty;
import javafx.beans.property.SimpleDoubleProperty;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.ValueAxis;
import javafx.util.StringConverter;

/**
 * 加工トレンド用の数値軸（{@link NumberAxis} は final のため {@link ValueAxis} から最小実装）。
 * <p>
 * JavaFX の {@code Axis} は目盛の表示位置が {@code 0 <= pos <= ceil(length)} の範囲外だとラベルを非表示にする。
 * 上限値の位置は {@code length + (upper - lower) * (-(length / (upper - lower)))} で求まり、軸の高さによっては
 * 浮動小数誤差で {@code -1e-13} 程度の負値になって最上段の目盛ラベルだけが消える（高さが変わると出たり消えたりする）。
 * {@link #getDisplayPosition(Number)} で 0 付近の誤差を 0 に丸めて安定させる。
 * <p>
 * 目盛は {@code lowerBound} から {@code tickUnit} 刻み。{@code upperBound} が刻みに乗らない場合のみ上限を追加する。
 * 自動レンジは {@link ProcessingTrendChartSupport#niceRange(double)} に委譲する（本タブは通常 {@code autoRanging=false}）。
 */
public class ProcessingTrendNumberAxis extends ValueAxis<Number> {

    private static final double ZERO_EPSILON = 1e-6;
    private static final int MAX_MAJOR_TICKS = 2000;

    private final DecimalFormat defaultFormat = new DecimalFormat("#,##0.###");

    private final DoubleProperty tickUnit = new SimpleDoubleProperty(this, "tickUnit", 5.0) {
        @Override
        protected void invalidated() {
            if (!isAutoRanging()) {
                invalidateRange();
                requestAxisLayout();
            }
        }
    };

    public ProcessingTrendNumberAxis() {
        super();
    }

    public final double getTickUnit() {
        return tickUnit.get();
    }

    public final void setTickUnit(double value) {
        tickUnit.set(value);
    }

    public final DoubleProperty tickUnitProperty() {
        return tickUnit;
    }

    @Override
    public double getDisplayPosition(Number value) {
        double pos = super.getDisplayPosition(value);
        return Math.abs(pos) < ZERO_EPSILON ? 0.0 : pos;
    }

    @Override
    protected String getTickMarkLabel(Number value) {
        StringConverter<Number> formatter = getTickLabelFormatter();
        if (formatter != null) {
            return formatter.toString(value);
        }
        return defaultFormat.format(value.doubleValue());
    }

    @Override
    protected Object getRange() {
        return new double[] {getLowerBound(), getUpperBound(), getTickUnit(), getScale()};
    }

    @Override
    protected void setRange(Object range, boolean animate) {
        double[] r = (double[]) range;
        setLowerBound(r[0]);
        setUpperBound(r[1]);
        setTickUnit(r[2]);
        currentLowerBound.set(r[0]);
        setScale(r[3]);
    }

    @Override
    protected Object autoRange(double minValue, double maxValue, double length, double labelSize) {
        double lower = Math.min(0.0, minValue);
        ProcessingTrendChartSupport.NiceRange nice = ProcessingTrendChartSupport.niceRange(Math.max(0.0, maxValue));
        double upper = Math.max(nice.upperBound(), lower + nice.tickUnit());
        return new double[] {lower, upper, nice.tickUnit(), calculateNewScale(length, lower, upper)};
    }

    @Override
    protected List<Number> calculateTickValues(double length, Object range) {
        double[] r = (double[]) range;
        double lower = r[0];
        double upper = r[1];
        double unit = r[2];
        List<Number> ticks = new ArrayList<>();
        if (!(upper > lower) || !(unit > 0)) {
            ticks.add(lower);
            if (upper != lower) {
                ticks.add(upper);
            }
            return ticks;
        }
        int count = (int) Math.floor((upper - lower) / unit + 1e-9);
        if (count > MAX_MAJOR_TICKS) {
            ticks.add(lower);
            ticks.add(upper);
            return ticks;
        }
        for (int i = 0; i <= count; i++) {
            ticks.add(lower + i * unit);
        }
        double last = lower + count * unit;
        if (upper - last > unit * 1e-6) {
            ticks.add(upper);
        }
        return ticks;
    }

    @Override
    protected List<Number> calculateMinorTickMarks() {
        return List.of();
    }
}
