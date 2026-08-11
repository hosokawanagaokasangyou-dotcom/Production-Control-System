package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.List;

import javafx.geometry.Rectangle2D;

import org.junit.jupiter.api.Test;

class MainStageScreenGeometryTest {

    @Test
    void resolveScreenBounds_matchesSavedMonitorWhenWindowCoordsAreOffScreen() {
        Rectangle2D primary = new Rectangle2D(0, 0, 1920, 1080);
        Rectangle2D secondary = new Rectangle2D(1920, 0, 1920, 1080);
        MainStageScreenGeometry.Snapshot geometry =
                new MainStageScreenGeometry.Snapshot(800, 600, 9999, 0, 1920, 0);

        Rectangle2D resolved =
                MainStageScreenGeometry.resolveScreenBoundsForTest(
                        geometry, List.of(primary, secondary));

        assertEquals(1920, resolved.getMinX());
        assertEquals(0, resolved.getMinY());
    }

    @Test
    void resolveScreenBounds_usesWindowIntersectionWhenSavedMonitorUnknown() {
        Rectangle2D primary = new Rectangle2D(0, 0, 1920, 1080);
        Rectangle2D secondary = new Rectangle2D(1920, 0, 1920, 1080);
        MainStageScreenGeometry.Snapshot geometry =
                new MainStageScreenGeometry.Snapshot(800, 600, 2000, 100, Double.NaN, Double.NaN);

        Rectangle2D resolved =
                MainStageScreenGeometry.resolveScreenBoundsForTest(
                        geometry, List.of(primary, secondary));

        assertEquals(1920, resolved.getMinX());
        assertEquals(0, resolved.getMinY());
    }
}
