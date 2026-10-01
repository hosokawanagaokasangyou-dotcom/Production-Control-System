package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * 書き出した検証結果ブックの行高を、Excel の「行の高さの自動調整」で内容に合わせる。
 */
public final class ExcelRowAutoFit {

    private ExcelRowAutoFit() {}

    /** Windows で Excel があるときだけ実行する。失敗しても検証結果ファイルは残す。 */
    public static boolean apply(List<Path> workbooks) {
        List<Path> files = existing(workbooks);
        if (files.isEmpty() || !isWindows()) {
            return false;
        }
        KouchinRunProgress.report("行高さを自動調整しています");
        Path script = null;
        try {
            script = Files.createTempFile("kouchin-row-autofit-", ".ps1");
            Files.write(script, withBom(scriptBody(files)));
            Process process = new ProcessBuilder(
                    "powershell.exe",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    script.toAbsolutePath().toString())
                    .redirectErrorStream(true)
                    .start();
            process.getInputStream().readAllBytes();
            boolean finished = process.waitFor(180, java.util.concurrent.TimeUnit.SECONDS);
            if (!finished) {
                process.destroyForcibly();
                return false;
            }
            return process.exitValue() == 0;
        } catch (IOException | InterruptedException ex) {
            if (ex instanceof InterruptedException) {
                Thread.currentThread().interrupt();
            }
            return false;
        } finally {
            if (script != null) {
                try {
                    Files.deleteIfExists(script);
                } catch (IOException ignored) {
                }
            }
        }
    }

    static boolean isWindows() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("windows");
    }

    /** 結合セルの行は列幅が狭い先頭列で伸びすぎるので、自動調整から外す。 */
    static String scriptBody(List<Path> files) {
        StringBuilder paths = new StringBuilder();
        for (Path file : files) {
            if (paths.length() > 0) {
                paths.append(", ");
            }
            paths.append('\'').append(escape(file.toAbsolutePath().toString())).append('\'');
        }
        return """
                $ErrorActionPreference = 'Stop'
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $excel.ScreenUpdating = $false
                $excel.AskToUpdateLinks = $false
                try {
                  foreach ($p in @(%s)) {
                    $wb = $excel.Workbooks.Open($p, 0, $false)
                    foreach ($sh in @($wb.Worksheets)) {
                      $used = $sh.UsedRange
                      if ($null -eq $used) { continue }
                      $merged = $used.MergeCells
                      if ($merged -eq $true -or $null -eq $merged) {
                        foreach ($row in @($used.Rows)) {
                          $wide = $false
                          foreach ($cell in @($row.Cells)) {
                            if ($cell.MergeCells -and $cell.MergeArea.Columns.Count -gt 1) {
                              $wide = $true
                              break
                            }
                          }
                          if (-not $wide) { $null = $row.AutoFit() }
                        }
                      } else {
                        $null = $used.EntireRow.AutoFit()
                      }
                    }
                    $wb.Save()
                    $wb.Close($true)
                  }
                } finally {
                  $excel.Quit()
                  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
                  [GC]::Collect()
                  [GC]::WaitForPendingFinalizers()
                }
                """.formatted(paths);
    }

    private static List<Path> existing(List<Path> workbooks) {
        List<Path> files = new ArrayList<>();
        if (workbooks == null) {
            return files;
        }
        for (Path file : workbooks) {
            if (file != null && Files.isRegularFile(file)) {
                files.add(file);
            }
        }
        return files;
    }

    private static String escape(String path) {
        return path.replace("'", "''");
    }

    private static byte[] withBom(String script) {
        byte[] body = script.getBytes(StandardCharsets.UTF_8);
        byte[] out = new byte[body.length + 3];
        out[0] = (byte) 0xEF;
        out[1] = (byte) 0xBB;
        out[2] = (byte) 0xBF;
        System.arraycopy(body, 0, out, 3, body.length);
        return out;
    }
}
