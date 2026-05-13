package jp.co.pm.ai.planning.stage2.input;

import java.nio.file.Path;
import java.time.LocalTime;
import java.util.List;
import java.util.Optional;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

/**
 * 段階2 Java エンジンが入力解決したスナップショット（件数ログ・配台の入力に使用）。
 */
public record Stage2InputSnapshot(
        Path masterPath,
        List<String> memberDisplayNames,
        Optional<LocalTime> factoryStart,
        Optional<LocalTime> factoryEnd,
        int excludeRuleCount,
        Path planInputPath,
        String planSheetName,
        PlanInputTabularIo.TabularSheet planningTasksSheet) {}
