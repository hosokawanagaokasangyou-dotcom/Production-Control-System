package jp.co.pm.ai.planning.stage2.output;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttContractSheetTableBuilder;

class Stage2EquipmentGanttContractPlaceholderTest {

    @Test
    void placeholderParsesInEquipmentGanttContractBuilder(@TempDir Path tmp) throws Exception {
        Path p = tmp.resolve("planTest設.json");
        Stage2EquipmentGanttContractPlaceholder.write(p, LocalDate.of(2026, 5, 14));
        assertDoesNotThrow(() -> EquipmentGanttContractSheetTableBuilder.buildBundleFromContractPath(p));
        Files.deleteIfExists(p);
    }
}
