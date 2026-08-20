package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/**
 * 配台マスタ保存: 世代バックアップのあと master ブックへ書き戻し、JSON を更新する。
 */
public final class MasterDispatchSheetsSaveWriter {

    public record Result(Path jsonPath, Path workbookPath, Path jsonBackup, Path workbookBackup) {}

    private MasterDispatchSheetsSaveWriter() {}

    public static Result save(
            Path jsonPath,
            Path workbookPath,
            MasterDispatchSheetsDocument document,
            Map<String, String> ui)
            throws IOException {
        Objects.requireNonNull(jsonPath, "jsonPath");
        Objects.requireNonNull(workbookPath, "workbookPath");
        Objects.requireNonNull(document, "document");
        Path workbook = workbookPath.toAbsolutePath().normalize();
        if (!Files.isRegularFile(workbook)) {
            throw new IOException("master ブックが見つかりません: " + workbook);
        }
        Map<String, String> env = ui != null ? ui : Map.of();
        Optional<Path> jsonBackup = MasterDispatchSheetsGenerationBackup.backupExisting(jsonPath);
        Optional<Path> workbookBackup = MasterDispatchSheetsGenerationBackup.backupExisting(workbook);
        MasterDispatchSheetWorkbookExporter.writeBack(workbook, document, env);
        MasterDispatchSheetsJsonStore.write(jsonPath, document);
        return new Result(
                jsonPath.toAbsolutePath().normalize(),
                workbook,
                jsonBackup.orElse(null),
                workbookBackup.orElse(null));
    }
}
