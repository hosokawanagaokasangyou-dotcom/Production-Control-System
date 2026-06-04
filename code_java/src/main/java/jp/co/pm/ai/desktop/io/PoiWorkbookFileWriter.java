package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * POI ワークブックを既存 Excel ファイルへ安全に上書きする。
 *
 * <p>ネットワーク上の xlsm へ {@link java.io.FileOutputStream} で直接書くと、保存失敗時に
 * 0 バイトへ切り詰められる。ローカル staging で完成させてから {@link Files#copy} で置換する。
 */
public final class PoiWorkbookFileWriter {

    private PoiWorkbookFileWriter() {}

    /**
     * {@code target} の内容を、完成した staging ファイルで置換する。
     *
     * @param target 上書き先（UNC 可）
     * @param workbook メモリ上のブック
     * @param ui {@link AppPaths#resolveRepoRoot} 解決用（staging ローカル退避）
     */
    public static void writeReplacing(Path target, Workbook workbook, Map<String, String> ui)
            throws IOException {
        Path normalized = target != null ? target.toAbsolutePath().normalize() : null;
        if (normalized == null) {
            throw new IOException("target path is null");
        }
        Path staging = allocateStagingFile(normalized, ui);
        IOException lastFailure = null;
        for (WriteMode mode : WriteMode.values()) {
            try {
                writeToStaging(staging, workbook, mode);
                validateStagingFile(staging);
                publishStagingToTarget(staging, normalized);
                return;
            } catch (Exception ex) {
                Files.deleteIfExists(staging);
                IOException failure =
                        ex instanceof IOException io ? io : new IOException(ex.getMessage(), ex);
                lastFailure = failure;
                if (mode == WriteMode.STANDARD && PoiWorkbookSaver.isPartNameFailure(ex)) {
                    continue;
                }
                throw failure;
            }
        }
        if (lastFailure != null) {
            throw lastFailure;
        }
        throw new IOException("failed to write workbook: " + normalized);
    }

    /**
     * バックアップファイル等を {@code target} へ安全にコピー置換する（staging 経由）。
     */
    public static void copyFileReplacing(Path source, Path target, Map<String, String> ui)
            throws IOException {
        Path normalizedSource = source != null ? source.toAbsolutePath().normalize() : null;
        Path normalizedTarget = target != null ? target.toAbsolutePath().normalize() : null;
        if (normalizedSource == null || !Files.isRegularFile(normalizedSource)) {
            throw new IOException("コピー元ファイルが見つかりません: " + source);
        }
        if (normalizedTarget == null) {
            throw new IOException("コピー先 path is null");
        }
        Path staging = allocateStagingFile(normalizedTarget, ui);
        try {
            Files.copy(normalizedSource, staging, StandardCopyOption.REPLACE_EXISTING);
            validateStagingFile(staging);
            publishStagingToTarget(staging, normalizedTarget);
        } catch (IOException ex) {
            Files.deleteIfExists(staging);
            throw ex;
        }
    }

    private enum WriteMode {
        STANDARD,
        LENIENT
    }

    static Path resolveStagingRoot(Map<String, String> ui) {
        String testRoot = System.getProperty("pm.ai.test.juchuWriteStagingRoot");
        if (testRoot != null && !testRoot.isBlank()) {
            return Path.of(testRoot).toAbsolutePath().normalize();
        }
        return AppPaths.resolveRepoRoot(ui != null ? ui : Map.of())
                .resolve(".pm-ai-cache")
                .resolve("juchu-write-staging")
                .toAbsolutePath()
                .normalize();
    }

    private static Path allocateStagingFile(Path target, Map<String, String> ui) throws IOException {
        Path stagingRoot = resolveStagingRoot(ui);
        Files.createDirectories(stagingRoot);
        String suffix = ".tmp";
        String baseName = target.getFileName() != null ? target.getFileName().toString() : "workbook.xlsm";
        if (baseName.length() > 80) {
            baseName = baseName.substring(0, 80);
        }
        return Files.createTempFile(stagingRoot, "juchu-", "-" + baseName + suffix).toAbsolutePath().normalize();
    }

    private static void writeToStaging(Path staging, Workbook workbook, WriteMode mode)
            throws IOException {
        try (OutputStream out = Files.newOutputStream(staging)) {
            if (mode == WriteMode.STANDARD) {
                PoiWorkbookSaver.write(workbook, out);
            } else {
                PoiWorkbookSaver.writeLenient(workbook, out);
            }
        }
    }

    private static void validateStagingFile(Path staging) throws IOException {
        if (!Files.isRegularFile(staging)) {
            throw new IOException("staging file missing: " + staging);
        }
        long size = Files.size(staging);
        if (size <= 0L) {
            throw new IOException("staging file is empty: " + staging);
        }
    }

    private static void publishStagingToTarget(Path staging, Path target) throws IOException {
        if (target.getParent() != null) {
            Files.createDirectories(target.getParent());
        }
        try {
            Files.move(
                    staging,
                    target,
                    StandardCopyOption.REPLACE_EXISTING,
                    StandardCopyOption.ATOMIC_MOVE);
        } catch (IOException moveEx) {
            Files.copy(staging, target, StandardCopyOption.REPLACE_EXISTING);
        } finally {
            Files.deleteIfExists(staging);
        }
    }
}
