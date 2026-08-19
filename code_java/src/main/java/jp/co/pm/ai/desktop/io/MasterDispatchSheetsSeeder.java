package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/**
 * 現在工場の JSON が無いときだけ、現在工場の master から吸い出す。他工場パスは受け取らない。
 */
public final class MasterDispatchSheetsSeeder {

    public enum Outcome {
        LOADED_EXISTING,
        IMPORTED,
        EMPTY_MISSING_SOURCE
    }

    public record Result(Outcome outcome, MasterDispatchSheetsDocument document) {
        public Result {
            Objects.requireNonNull(outcome, "outcome");
            Objects.requireNonNull(document, "document");
        }
    }

    private MasterDispatchSheetsSeeder() {}

    public static Result loadOrImport(Path jsonPath, Path sourceWorkbook, String factorySite)
            throws IOException {
        return loadOrImport(jsonPath, sourceWorkbook, factorySite, false);
    }

    public static Result loadOrImport(
            Path jsonPath, Path sourceWorkbook, String factorySite, boolean reimport)
            throws IOException {
        Objects.requireNonNull(jsonPath, "jsonPath");
        Objects.requireNonNull(sourceWorkbook, "sourceWorkbook");
        String site = factorySite != null ? factorySite : "";
        if (!reimport && Files.isRegularFile(jsonPath)) {
            return new Result(Outcome.LOADED_EXISTING, MasterDispatchSheetsJsonStore.read(jsonPath));
        }
        if (!Files.isRegularFile(sourceWorkbook)) {
            return new Result(Outcome.EMPTY_MISSING_SOURCE, MasterDispatchSheetsDocument.empty(site));
        }
        MasterDispatchSheetsDocument doc =
                MasterDispatchSheetWorkbookImporter.importWorkbook(sourceWorkbook, site);
        MasterDispatchSheetsJsonStore.write(jsonPath, doc);
        return new Result(Outcome.IMPORTED, doc);
    }
}
