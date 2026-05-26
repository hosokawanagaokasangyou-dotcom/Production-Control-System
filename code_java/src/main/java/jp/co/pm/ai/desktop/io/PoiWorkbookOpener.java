package jp.co.pm.ai.desktop.io;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Apache POI で xlsx/xlsm を開く際の Zip セキュリティ上限を緩和する。
 *
 * <p>マクロ付き依頼書など、ZIP 内部エントリ数が多いブックは既定 {@code MAX_FILE_COUNT=1000} で
 * 拒否されることがある。
 */
public final class PoiWorkbookOpener {

    /** 依頼書 xlsm 等を想定した上限（POI 既定 1000 より緩い）。 */
    private static final int MAX_ZIP_FILE_COUNT = 10_000;

    private static volatile boolean zipLimitsConfigured;

    static {
        configureZipSecureLimits();
    }

    private PoiWorkbookOpener() {}

    public static void configureZipSecureLimits() {
        if (zipLimitsConfigured) {
            return;
        }
        synchronized (PoiWorkbookOpener.class) {
            if (zipLimitsConfigured) {
                return;
            }
            ZipSecureFile.setMaxFileCount(MAX_ZIP_FILE_COUNT);
            zipLimitsConfigured = true;
        }
    }

    public static Workbook open(File file) throws IOException {
        configureZipSecureLimits();
        try (FileInputStream in = new FileInputStream(file)) {
            return WorkbookFactory.create(in);
        }
    }

    public static Workbook open(InputStream in) throws IOException {
        configureZipSecureLimits();
        return WorkbookFactory.create(in);
    }
}
