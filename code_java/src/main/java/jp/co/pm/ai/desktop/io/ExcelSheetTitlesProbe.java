package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/** Reads sheet names via Apache POI (openpyxl alignment path per plan). */
public final class ExcelSheetTitlesProbe {

    private ExcelSheetTitlesProbe() {}

    public static List<String> sheetNames(Path workbookPath) throws IOException {
        if (!Files.isRegularFile(workbookPath)) {
            throw new IOException("not a file: " + workbookPath);
        }
        try (Workbook wb = WorkbookFactory.create(workbookPath.toFile())) {
            int n = wb.getNumberOfSheets();
            List<String> names = new ArrayList<>(n);
            for (int i = 0; i < n; i++) {
                names.add(wb.getSheetName(i));
            }
            return names;
        }
    }
}
