package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PlanningCoreMaterialTableAppendProbeTest {

    @Test
    void detect_repoFacadePlanningCore_isCurrent() {
        Path repoRoot = AppPaths.resolveRepoRoot(java.util.Map.of());
        Path codePython = repoRoot.resolve("code").resolve("python");
        if (!Files.isDirectory(codePython)) {
            return;
        }
        PlanningCoreMaterialTableAppendProbe.Result result =
                PlanningCoreMaterialTableAppendProbe.detect(codePython);
        assertEquals(PlanningCoreMaterialTableAppendProbe.Spec.CURRENT, result.spec());
        assertTrue(result.buildId().isPresent());
        assertTrue(result.buildId().get().contains("write-canonical"));
    }

    @Test
    void detect_legacyMonolithicCore_isLegacy(@TempDir Path tmp) throws Exception {
        Path planningCore = tmp.resolve("planning_core");
        Files.createDirectories(planningCore);
        Files.writeString(
                planningCore.resolve("_core.py"),
                "# old monolithic\nprint('hello')\n",
                StandardCharsets.UTF_8);
        PlanningCoreMaterialTableAppendProbe.Result result =
                PlanningCoreMaterialTableAppendProbe.detect(tmp);
        assertEquals(PlanningCoreMaterialTableAppendProbe.Spec.LEGACY, result.spec());
        assertFalse(result.buildId().isPresent());
    }

    @Test
    void detect_facadeWithColumnsMarker_isCurrent(@TempDir Path tmp) throws Exception {
        Path planningCore = tmp.resolve("planning_core");
        Path coreDir = planningCore.resolve("core");
        Files.createDirectories(coreDir);
        Files.writeString(
                planningCore.resolve("_core.py"),
                """
                _MODULE_ORDER = ["columns"]
                def _exec_into_ns(name): pass
                """,
                StandardCharsets.UTF_8);
        Files.writeString(
                coreDir.resolve("columns.py"),
                "_STAGE1_MATERIAL_TABLE_APPEND_BUILD = \"test-build-id\"\n",
                StandardCharsets.UTF_8);
        PlanningCoreMaterialTableAppendProbe.Result result =
                PlanningCoreMaterialTableAppendProbe.detect(tmp);
        assertEquals(PlanningCoreMaterialTableAppendProbe.Spec.CURRENT, result.spec());
        assertEquals("test-build-id", result.buildId().orElseThrow());
    }
}
