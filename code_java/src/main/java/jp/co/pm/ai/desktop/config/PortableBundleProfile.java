package jp.co.pm.ai.desktop.config;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/** ポータブル配布の種別（配台 PMD / RDP ランチャー）。 */
public enum PortableBundleProfile {

    PMD(
            "PMD_version_upgrade.zip",
            "PMD.exe",
            "launch-pm-ai-desktop.bat",
            "pmd-apply-portable-update.ps1",
            ".pm-ai-desktop",
            "pending-portable-update.json",
            "pending-portable-update-staging",
            "pending-portable-update-apply.log",
            "pending-portable-upgrade-followup.json",
            true,
            List.of(),
            List.of()),

    RDP_LAUNCHER(
            AppPaths.RDP_DESKTOP_LAUNCHER_VERSION_UPGRADE_ZIP,
            AppPaths.RDP_DESKTOP_LAUNCHER_EXE_BASENAME,
            "launch-pm-ai-rpa-luncher.bat",
            "rdp-apply-portable-update.ps1",
            ".pm-ai-desktop-rdp",
            "pending-rdp-portable-update.json",
            "pending-rdp-portable-update-staging",
            "pending-rdp-portable-update-apply.log",
            null,
            false,
            List.of("README_PORTABLE_RDP.txt"),
            List.of("launcher-deploy-seed"));

    private final String upgradeZipName;
    private final String desktopExeName;
    private final String launcherBatName;
    private final String applyScriptName;
    private final String userStateDirName;
    private final String pendingManifestFileName;
    private final String stagingDirName;
    private final String applyLogFileName;
    private final String upgradeFollowUpFileName;
    private final boolean syncsPmAiData;
    private final List<String> extraDesktopLeafFiles;
    private final List<String> extraDesktopDirectories;

    PortableBundleProfile(
            String upgradeZipName,
            String desktopExeName,
            String launcherBatName,
            String applyScriptName,
            String userStateDirName,
            String pendingManifestFileName,
            String stagingDirName,
            String applyLogFileName,
            String upgradeFollowUpFileName,
            boolean syncsPmAiData,
            List<String> extraDesktopLeafFiles,
            List<String> extraDesktopDirectories) {
        this.upgradeZipName = upgradeZipName;
        this.desktopExeName = desktopExeName;
        this.launcherBatName = launcherBatName;
        this.applyScriptName = applyScriptName;
        this.userStateDirName = userStateDirName;
        this.pendingManifestFileName = pendingManifestFileName;
        this.stagingDirName = stagingDirName;
        this.applyLogFileName = applyLogFileName;
        this.upgradeFollowUpFileName = upgradeFollowUpFileName;
        this.syncsPmAiData = syncsPmAiData;
        this.extraDesktopLeafFiles = extraDesktopLeafFiles;
        this.extraDesktopDirectories = extraDesktopDirectories;
    }

    public String upgradeZipName() {
        return upgradeZipName;
    }

    public String desktopExeName() {
        return desktopExeName;
    }

    public String launcherBatName() {
        return launcherBatName;
    }

    public String applyScriptName() {
        return applyScriptName;
    }

    public String userStateDirName() {
        return userStateDirName;
    }

    public String pendingManifestFileName() {
        return pendingManifestFileName;
    }

    public String stagingDirName() {
        return stagingDirName;
    }

    public String applyLogFileName() {
        return applyLogFileName;
    }

    public String upgradeFollowUpFileName() {
        return upgradeFollowUpFileName;
    }

    public boolean syncsPmAiData() {
        return syncsPmAiData;
    }

    public List<String> extraDesktopLeafFiles() {
        return extraDesktopLeafFiles;
    }

    public List<String> extraDesktopDirectories() {
        return extraDesktopDirectories;
    }

    /** {@code cwd} が当該ポータブル配布レイアウトか。 */
    public boolean isPortableBundleLayout(Path cwd) {
        if (cwd == null) {
            return false;
        }
        Path abs = cwd.toAbsolutePath().normalize();
        if (syncsPmAiData) {
            Path marker =
                    abs.resolve("pm-ai-data")
                            .resolve("code")
                            .resolve("python")
                            .resolve("task_extract_stage1.py");
            return Files.isRegularFile(marker);
        }
        return PortableBundleSelfUpdater.hasDesktopInstallLayout(this, abs);
    }

    /** デスクトップ本体同期時にコピーするリーフファイル名（exe / bat / script / version 以外）。 */
    public List<String> desktopLeafFilesToSync() {
        List<String> leaves = new java.util.ArrayList<>();
        leaves.add(desktopExeName);
        leaves.add(launcherBatName);
        leaves.add(applyScriptName);
        leaves.add(AppPaths.VERSION_TXT_FILE_NAME);
        leaves.addAll(extraDesktopLeafFiles);
        return leaves;
    }
}
