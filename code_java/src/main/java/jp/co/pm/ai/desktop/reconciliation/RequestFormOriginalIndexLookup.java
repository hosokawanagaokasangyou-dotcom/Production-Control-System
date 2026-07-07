package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.dispatch.DispatchAladdinEntrySheetBuilder;
import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;

/**
 * 依頼書原本フォルダ（{@link AppPaths#resolveRequestFormOriginalDir}）配下の各 xlsm の
 * 「目次」シートから、依頼NO 別の回答納期（K列「納期」）・契約NO（N列「契約No」）を読む。
 *
 * <p>原本フォルダは読み取り専用（一覧・読込のみ。書込・キャッシュ作成はしない）。
 */
public final class RequestFormOriginalIndexLookup {

    private RequestFormOriginalIndexLookup() {}

    /** 走査の進捗通知（{@code processed} / {@code total} ファイル）。 */
    @FunctionalInterface
    public interface ScanProgressListener {
        void onProgress(int processed, int total);
    }

    /**
     * 依頼NO 正規化キー → 目次情報。フォルダ未到達・目次なしのときは空 Map（警告を {@code warnings} へ追記）。
     *
     * @param warnings null 可（不要なら渡さない）
     */
    public static Map<String, DispatchAladdinEntrySheetBuilder.IndexInfo> loadByIraiNoKey(
            Map<String, String> ui, List<String> warnings) {
        return loadByIraiNoKey(ui, warnings, null);
    }

    /**
     * 進捗通知付き。{@code progress} はワーカースレッドから呼ばれる（UI 更新側で {@code Platform.runLater} すること）。
     */
    public static Map<String, DispatchAladdinEntrySheetBuilder.IndexInfo> loadByIraiNoKey(
            Map<String, String> ui, List<String> warnings, ScanProgressListener progress) {
        Map<String, DispatchAladdinEntrySheetBuilder.IndexInfo> out = new LinkedHashMap<>();
        Path dir = AppPaths.resolveRequestFormOriginalDir(ui);
        if (!NetworkSourceDirResolver.isRequestFormOriginalDirReachable(ui)) {
            warn(warnings, "依頼書原本フォルダにアクセスできません: " + dir);
            return out;
        }
        File[] files =
                dir.toFile()
                        .listFiles(
                                (d, name) ->
                                        name.endsWith(".xlsm")
                                                && !name.startsWith("~$")
                                                && !name.equals("加工依頼書入力.xlsm"));
        if (files == null || files.length == 0) {
            warn(warnings, "Excel 依頼書原本が見つかりません: " + dir);
            return out;
        }
        int processed = 0;
        for (File file : files) {
            try (Workbook wb = PoiWorkbookOpener.open(file)) {
                Sheet index = wb.getSheet(RequestFormOriginalIndexSheetLayout.SHEET_NAME);
                if (index == null) {
                    continue;
                }
                for (Map.Entry<String, RequestFormOriginalIndexSheetReader.IndexEntry> e :
                        RequestFormOriginalIndexSheetReader.read(index).entrySet()) {
                    RequestFormOriginalIndexSheetReader.IndexEntry entry = e.getValue();
                    out.putIfAbsent(
                            e.getKey(),
                            new DispatchAladdinEntrySheetBuilder.IndexInfo(
                                    entry.deliveryDate(), entry.contractNo()));
                }
            } catch (Exception ex) {
                warn(warnings, "原本目次読込エラー " + file.getName() + ": " + ex.getMessage());
            }
            processed++;
            if (progress != null) {
                progress.onProgress(processed, files.length);
            }
        }
        return out;
    }

    /** 依頼NO を目次キーと同じ正規化で揃える。 */
    public static String normalizeIraiNoKey(String iraiNo) {
        return JuchuTransferValueNormalizer.normalizeKey(iraiNo);
    }

    private static void warn(List<String> warnings, String message) {
        if (warnings != null) {
            warnings.add(message);
        }
    }
}
