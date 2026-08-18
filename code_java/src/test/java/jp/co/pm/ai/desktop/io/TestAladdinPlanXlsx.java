package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** テスト用の最小アラジン加工計画 xlsx。 */
public final class TestAladdinPlanXlsx {

    private TestAladdinPlanXlsx() {}

    public static Path writeGrid(Path dir, String fileName, String[][] cells) throws IOException {
        Files.createDirectories(dir);
        Path out = dir.resolve(fileName);
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Sheet1");
            for (int r = 0; r < cells.length; r++) {
                Row row = sheet.createRow(r);
                for (int c = 0; c < cells[r].length; c++) {
                    row.createCell(c).setCellValue(cells[r][c] != null ? cells[r][c] : "");
                }
            }
            try (OutputStream os = Files.newOutputStream(out)) {
                wb.write(os);
            }
        }
        return out.toAbsolutePath().normalize();
    }

    public static Path writeMinimal(Path dir, String fileName) throws IOException {
        return writeGrid(
                dir,
                fileName,
                new String[][] {
                    {"列1", "列2", "列3", "列4"},
                    {"上段1", "", "", ""},
                    {"上段2", "", "", ""},
                    {"上段3", "", "", ""},
                    {"機械名", "依頼NO", "工程名", "2026/07/07"},
                    {"", "", "", ""},
                    {"M1", "T001", "工程A", "10"}
                });
    }

    public static Path writeWithQty(Path dir, String fileName, String qty) throws IOException {
        return writeGrid(
                dir,
                fileName,
                new String[][] {
                    {"列1", "列2", "列3", "列4"},
                    {"上段1", "", "", ""},
                    {"上段2", "", "", ""},
                    {"上段3", "", "", ""},
                    {"機械名", "依頼NO", "工程名", "2026/07/07"},
                    {"", "", "", ""},
                    {"M1", "T001", "工程A", qty}
                });
    }
}
