package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * {@code code/} 直下の「キー,値」2 列 CSV（1 行目ヘッダ）を UTF-8 で読み書きする。キーに {@code ','} が含まれる場合は
 * 最後の {@code ','} を区切りとする（{@code planning_core} のロール長テーブルと同趣旨）。
 */
public final class CodeDispatchLookupTableIo {

    private CodeDispatchLookupTableIo() {}

    public record KeyValTable(String headerLine, LinkedHashMap<String, String> rows) {
        public KeyValTable {
            Objects.requireNonNull(headerLine, "headerLine");
            rows = rows != null ? new LinkedHashMap<>(rows) : new LinkedHashMap<>();
        }
    }

    public static KeyValTable readOrEmpty(Path path, String defaultHeaderLine) throws IOException {
        if (path == null || defaultHeaderLine == null || defaultHeaderLine.isBlank()) {
            throw new IOException("path/header");
        }
        if (!Files.isRegularFile(path)) {
            return new KeyValTable(defaultHeaderLine.strip(), new LinkedHashMap<>());
        }
        List<String> lines = Files.readAllLines(path, StandardCharsets.UTF_8);
        if (lines.isEmpty()) {
            return new KeyValTable(defaultHeaderLine.strip(), new LinkedHashMap<>());
        }
        String header = lines.getFirst().strip();
        if (header.isEmpty()) {
            header = defaultHeaderLine.strip();
        }
        LinkedHashMap<String, String> rows = new LinkedHashMap<>();
        for (int i = 1; i < lines.size(); i++) {
            String line = lines.get(i);
            if (line == null) {
                continue;
            }
            String t = line.strip();
            if (t.isEmpty()) {
                continue;
            }
            int c = t.lastIndexOf(',');
            if (c <= 0 || c >= t.length() - 1) {
                continue;
            }
            String k = t.substring(0, c).strip();
            String v = t.substring(c + 1).strip();
            if (k.isEmpty()) {
                continue;
            }
            rows.putIfAbsent(k, v);
        }
        return new KeyValTable(header, rows);
    }

    public static void write(Path path, KeyValTable table) throws IOException {
        Objects.requireNonNull(path, "path");
        Objects.requireNonNull(table, "table");
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        List<String> out = new ArrayList<>(1 + table.rows().size());
        out.add(table.headerLine().strip());
        for (Map.Entry<String, String> e : table.rows().entrySet()) {
            String k = e.getKey() != null ? e.getKey() : "";
            String v = e.getValue() != null ? e.getValue() : "";
            out.add(k + "," + v);
        }
        Files.write(path, out, StandardCharsets.UTF_8);
    }
}
