package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage3PlanningMetaStoreTest {

    @TempDir Path dir;

    @Test
    void deleteSidecar_removesStaleVariantSoPlanningStageIsStage2() throws Exception {
        Path dispatchJson = dir.resolve("結果_配台表.json");
        Files.writeString(dispatchJson, "{\"columns\":[],\"rows\":[]}", StandardCharsets.UTF_8);
        Stage3PlanningMetaStore.writeVariant(dispatchJson, Stage3PlanningMetaStore.Variant.STAGE3_1);
        assertTrue(Stage3PlanningMetaStore.hasPipelinePlanningVariant(dispatchJson));
        assertEquals(
                ResultDispatchStage3Support.PlanningStage.STAGE3,
                ResultDispatchStage3Support.detectPlanningStage(dispatchJson));

        Stage3PlanningMetaStore.deleteSidecar(dispatchJson);

        assertFalse(Stage3PlanningMetaStore.hasPipelinePlanningVariant(dispatchJson));
        assertFalse(Files.isRegularFile(Stage3PlanningMetaStore.sidecarPath(dispatchJson)));
        assertEquals(
                ResultDispatchStage3Support.PlanningStage.STAGE2,
                ResultDispatchStage3Support.detectPlanningStage(dispatchJson));
        assertEquals(
                ResultDispatchStage3Support.Stage3PlanningVariant.NONE,
                Stage3PlanningMetaStore.readPlanningVariant(dispatchJson));
    }
}
