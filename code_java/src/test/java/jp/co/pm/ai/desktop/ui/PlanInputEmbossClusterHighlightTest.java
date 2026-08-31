package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class PlanInputEmbossClusterHighlightTest {

    @Test
    void loadWithEmboss_lightsClusterNotSave() {
        PlanInputEmbossClusterHighlight h = new PlanInputEmbossClusterHighlight();
        h.resetForLoadedTable();
        assertTrue(h.clusterHot(true));
        assertFalse(h.saveHot());
    }

    @Test
    void afterCluster_stopsClusterLight_lightsSave() {
        PlanInputEmbossClusterHighlight h = new PlanInputEmbossClusterHighlight();
        h.resetForLoadedTable();
        h.markClusterApplied();
        assertFalse(h.clusterHot(true));
        assertTrue(h.saveHot());
    }

    @Test
    void afterSave_stopsSaveLight_keepsClusterOffWhileEmbossRemains() {
        PlanInputEmbossClusterHighlight h = new PlanInputEmbossClusterHighlight();
        h.resetForLoadedTable();
        h.markClusterApplied();
        h.markSaved();
        assertFalse(h.saveHot());
        assertFalse(h.clusterHot(true));
    }

    @Test
    void reloadWithEmboss_lightsClusterAgain() {
        PlanInputEmbossClusterHighlight h = new PlanInputEmbossClusterHighlight();
        h.resetForLoadedTable();
        h.markClusterApplied();
        h.markSaved();
        h.resetForLoadedTable();
        assertTrue(h.clusterHot(true));
        assertFalse(h.saveHot());
    }
}
