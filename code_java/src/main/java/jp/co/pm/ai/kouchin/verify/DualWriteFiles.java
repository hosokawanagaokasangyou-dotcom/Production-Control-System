package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.io.PoiWorkbookFileWriter;

/**
 * 同一内容を複数出力先へ書く。片方失敗しても他方は継続する。
 */
public final class DualWriteFiles {

    private DualWriteFiles() {}

    public record WriteOutcome(List<Path> succeeded, List<String> failures) {
        public boolean allFailed() {
            return succeeded.isEmpty();
        }

        public boolean anyFailed() {
            return !failures.isEmpty();
        }
    }

    public static WriteOutcome writeWorkbook(
            Workbook workbook, List<Path> targets, Map<String, String> ui) {
        List<Path> ok = new ArrayList<>();
        List<String> fail = new ArrayList<>();
        for (Path target : targets) {
            try {
                Path parent = target.getParent();
                if (parent != null) {
                    Files.createDirectories(parent);
                }
                PoiWorkbookFileWriter.writeReplacing(target, workbook, ui);
                ok.add(target);
            } catch (IOException | RuntimeException ex) {
                fail.add(target + ": " + (ex.getMessage() == null ? ex.toString() : ex.getMessage()));
            }
        }
        return new WriteOutcome(List.copyOf(ok), List.copyOf(fail));
    }

    public static WriteOutcome writeBytes(byte[] bytes, List<Path> targets) {
        List<Path> ok = new ArrayList<>();
        List<String> fail = new ArrayList<>();
        for (Path target : targets) {
            try {
                Path parent = target.getParent();
                if (parent != null) {
                    Files.createDirectories(parent);
                }
                Files.write(target, bytes);
                ok.add(target);
            } catch (IOException | RuntimeException ex) {
                fail.add(target + ": " + (ex.getMessage() == null ? ex.toString() : ex.getMessage()));
            }
        }
        return new WriteOutcome(List.copyOf(ok), List.copyOf(fail));
    }

    public static WriteOutcome copyFile(Path source, List<Path> targets) {
        List<Path> ok = new ArrayList<>();
        List<String> fail = new ArrayList<>();
        for (Path target : targets) {
            try {
                Path parent = target.getParent();
                if (parent != null) {
                    Files.createDirectories(parent);
                }
                Files.copy(source, target, StandardCopyOption.REPLACE_EXISTING);
                ok.add(target);
            } catch (IOException | RuntimeException ex) {
                fail.add(target + ": " + (ex.getMessage() == null ? ex.toString() : ex.getMessage()));
            }
        }
        return new WriteOutcome(List.copyOf(ok), List.copyOf(fail));
    }
}
