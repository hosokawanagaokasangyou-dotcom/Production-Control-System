package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class DispatchResultPathsTest {

    @TempDir Path temp;

    @Test
    void snapshotDisplayDoesNotSubstituteLocalFiles() throws Exception {
        Path share = temp.resolve("share");
        Path local = temp.resolve("local");
        Files.createDirectories(local);
        Files.writeString(local.resolve("結果_配台表.json"), "{}", StandardCharsets.UTF_8);
        Files.writeString(local.resolve("計画2609230900000001.json"), "{}", StandardCharsets.UTF_8);
        String gen = "20260923-090000-001_pc_stage2";
        Path genDir = share.resolve("森岡").resolve(gen);
        Files.createDirectories(genDir);
        Files.writeString(genDir.resolve("計画2609230900000002.json"), "{}", StandardCharsets.UTF_8);

        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_DISPATCH_SNAPSHOT_DIR,
                        share.toString(),
                        AppPaths.KEY_PM_AI_OUTPUT_DIR,
                        local.toString());
        DispatchResultSelection selection = new DispatchResultSelection();
        assertTrue(selection.selectSnapshot("森岡", gen, "他者の結果: 森岡", "detail", false));

        Path plan = DispatchResultPaths.planJson(ui, selection);
        Path dispatch = DispatchResultPaths.dispatchJson(ui, selection);
        Path member = DispatchResultPaths.memberJson(ui, selection);
        Path shaped = DispatchResultPaths.shapedAladdin(ui, selection);
        Path localDispatch = AppPaths.resolveResultDispatchTableJsonPath(ui);

        assertNotNull(plan);
        assertTrue(plan.startsWith(genDir));
        assertNotNull(dispatch);
        assertTrue(dispatch.startsWith(genDir));
        assertFalse(dispatch.equals(localDispatch));
        assertNull(member);
        assertNotNull(shaped);
        assertTrue(shaped.startsWith(genDir));
    }

    @Test
    void localLatestStillUsesOutputDir() throws Exception {
        Path local = temp.resolve("local");
        Files.createDirectories(local);
        Path dispatchFile = local.resolve("結果_配台表.json");
        Files.writeString(dispatchFile, "{}", StandardCharsets.UTF_8);
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_OUTPUT_DIR, local.toString());
        DispatchResultSelection selection = new DispatchResultSelection();
        assertTrue(selection.selectLocal());

        Path dispatch = DispatchResultPaths.dispatchJson(ui, selection);
        assertEquals(dispatchFile.toAbsolutePath().normalize(), dispatch);
    }

    @Test
    void localMemberFollowsPlanStampNotNewestMember() throws Exception {
        Path local = temp.resolve("local");
        Files.createDirectories(local);
        Path plan = local.resolve("計画2609230900000002.json");
        Path memberSame = local.resolve("人員2609230900000002.json");
        Path memberOther = local.resolve("人員2609230900000001.json");
        Files.writeString(plan, "{}", StandardCharsets.UTF_8);
        Files.writeString(memberSame, "{}", StandardCharsets.UTF_8);
        Files.writeString(memberOther, "{}", StandardCharsets.UTF_8);
        Files.setLastModifiedTime(memberOther, java.nio.file.attribute.FileTime.fromMillis(9_000));
        Files.setLastModifiedTime(plan, java.nio.file.attribute.FileTime.fromMillis(5_000));
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_OUTPUT_DIR, local.toString());
        DispatchResultSelection selection = new DispatchResultSelection();

        assertEquals(memberSame.toAbsolutePath().normalize(), DispatchResultPaths.memberJson(ui, selection));
    }
}
