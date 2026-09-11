package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

/** 検査表索引 CSV（工場別）。 */
public final class InspectionSheetIndexStore {

    public static final String HEADER =
            "irai_no,processing_date,processing_year_month,file_path,file_name,file_mtime_epoch,file_size,indexed_at";

    public record Row(
            String iraiNo,
            LocalDate processingDate,
            String processingYearMonth,
            String filePath,
            String fileName,
            long fileMtimeEpoch,
            long fileSize,
            String indexedAt) {}

    private InspectionSheetIndexStore() {}

    public static Path indexFile(FactorySite site) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        if (effective == FactorySite.RDP_LAUNCHER) {
            effective = FactorySite.KONAN;
        }
        return AppPaths.resolveDesktopAppHomeDir()
                .resolve("inspection-sheet-index")
                .resolve(effective.name() + ".csv");
    }

    public static List<Row> load(Path csv) throws IOException {
        if (csv == null || !Files.isRegularFile(csv)) {
            return List.of();
        }
        List<String> lines = Files.readAllLines(csv, StandardCharsets.UTF_8);
        List<Row> out = new ArrayList<>();
        boolean header = true;
        for (String line : lines) {
            if (header) {
                header = false;
                continue;
            }
            if (line == null || line.isBlank()) {
                continue;
            }
            List<String> cols = parseCsvLine(line);
            if (cols.size() < 8) {
                continue;
            }
            out.add(
                    new Row(
                            cols.get(0),
                            parseDate(cols.get(1)),
                            cols.get(2),
                            cols.get(3),
                            cols.get(4),
                            parseLong(cols.get(5)),
                            parseLong(cols.get(6)),
                            cols.get(7)));
        }
        return List.copyOf(out);
    }

    public static void save(Path csv, List<Row> rows) throws IOException {
        Files.createDirectories(csv.getParent());
        StringBuilder sb = new StringBuilder();
        sb.append(HEADER).append('\n');
        if (rows != null) {
            for (Row row : rows) {
                sb.append(toCsvLine(row)).append('\n');
            }
        }
        Files.writeString(csv, sb.toString(), StandardCharsets.UTF_8);
    }

    static String toCsvLine(Row row) {
        return String.join(
                ",",
                csv(row.iraiNo()),
                csv(row.processingDate() != null ? row.processingDate().toString() : ""),
                csv(row.processingYearMonth()),
                csv(row.filePath()),
                csv(row.fileName()),
                csv(Long.toString(row.fileMtimeEpoch())),
                csv(Long.toString(row.fileSize())),
                csv(row.indexedAt()));
    }

    static String csv(String value) {
        String v = value != null ? value : "";
        if (v.indexOf(',') >= 0 || v.indexOf('"') >= 0 || v.indexOf('\n') >= 0 || v.indexOf('\r') >= 0) {
            return '"' + v.replace("\"", "\"\"") + '"';
        }
        return v;
    }

    static List<String> parseCsvLine(String line) {
        List<String> out = new ArrayList<>();
        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);
            if (inQuotes) {
                if (c == '"') {
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') {
                        cur.append('"');
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    cur.append(c);
                }
            } else if (c == '"') {
                inQuotes = true;
            } else if (c == ',') {
                out.add(cur.toString());
                cur.setLength(0);
            } else {
                cur.append(c);
            }
        }
        out.add(cur.toString());
        return out;
    }

    private static LocalDate parseDate(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        try {
            return LocalDate.parse(raw.strip());
        } catch (RuntimeException ex) {
            return null;
        }
    }

    private static long parseLong(String raw) {
        try {
            return Long.parseLong(raw.strip());
        } catch (RuntimeException ex) {
            return 0L;
        }
    }
}
