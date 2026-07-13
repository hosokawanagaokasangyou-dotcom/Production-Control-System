package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PostProcessingPlanMachineLookupTest {

    @TempDir Path tempDir;

    @Test
    void snapshotFromFile_buildsCodeToNameFromPlanColumns() throws Exception {
        Path csv = tempDir.resolve("plan.csv");
        Files.writeString(
                csv,
                "機械,機械名,依頼NO\n"
                        + "M01,スライス機1　湖南,R1\n"
                        + "M02,検査機,R2\n");

        PostProcessingPlanMachineLookup.invalidate();
        PostProcessingPlanMachineLookup.Snapshot snap =
                PostProcessingPlanMachineLookup.snapshotFromFile(csv);

        assertTrue(snap.loaded());
        assertTrue(snap.hasCodeColumn());
        assertTrue(snap.hasNameColumn());
        assertEquals("スライス機1　湖南", snap.machineCodeToName().get("M01"));
        assertEquals("M01", PostProcessingPlanMachineLookup.resolveMachineCodeFromName(snap, "スライス機1　湖南"));
        assertEquals("M02", PostProcessingPlanMachineLookup.resolveCodeFromComboInput(snap, "M02 検査機"));
        assertTrue(PostProcessingPlanMachineLookup.isMachineCodeColumn("機械コード3"));
    }

    @Test
    void legacyMachineCodeColumnHeader_stillWorks() throws Exception {
        Path csv = tempDir.resolve("plan-legacy.csv");
        Files.writeString(
                csv,
                "機械コード,機械名,依頼NO\n" + "L01,旧見出し,R1\n");

        PostProcessingPlanMachineLookup.invalidate();
        PostProcessingPlanMachineLookup.Snapshot snap =
                PostProcessingPlanMachineLookup.snapshotFromFile(csv);

        assertTrue(snap.loaded());
        assertTrue(snap.hasCodeColumn());
        assertEquals("旧見出し", snap.machineCodeToName().get("L01"));
    }

    @Test
    void nameOnlyColumn_usesNameAsKey() throws Exception {
        Path csv = tempDir.resolve("plan-name-only.csv");
        Files.writeString(csv, "機械名,依頼NO\nスライス機1,R1\n");

        PostProcessingPlanMachineLookup.invalidate();
        PostProcessingPlanMachineLookup.Snapshot snap =
                PostProcessingPlanMachineLookup.snapshotFromFile(csv);

        assertTrue(snap.loaded());
        assertEquals("スライス機1", snap.machineCodeToName().get("スライス機1"));
    }
}
