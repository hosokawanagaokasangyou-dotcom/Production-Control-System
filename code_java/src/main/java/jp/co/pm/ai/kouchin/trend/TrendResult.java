package jp.co.pm.ai.kouchin.trend;

import java.nio.file.Path;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.kouchin.verify.DualWriteFiles;

/**
 * 月次トレンド実行結果。{@link #workbook()} は書込後も開いたままなので呼び出し側で close する。
 */
public record TrendResult(
        XSSFWorkbook workbook,
        DualWriteFiles.WriteOutcome xlsx,
        Path preferred,
        String stamp,
        TrendModel model,
        List<String> warnings) {

    public TrendResult {
        warnings = warnings == null ? List.of() : List.copyOf(warnings);
    }

    public Path preferredExcel() {
        return preferred;
    }
}
