package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;

/**
 * 依頼書原本フォルダ（読み取り専用）から、依頼NOに対応する xlsm とシート名を探す。
 */
public final class RequestFormOriginalFileFinder {

    /**
     * 見つかった原本。{@code sheetName} が空ならブック内に依頼シートが無い。
     * {@code fee} は配台の工程順で Gemini が計算した加工賃。計算できないときは null。
     */
    public record Found(String iraiNo, Path file, String sheetName, RequestFormOriginalFee.Result fee) {}

    private RequestFormOriginalFileFinder() {}

    /**
     * @param keys 正規化済み依頼NO。空なら何も開かない
     * @return キー → 原本。見つからなかったキーは含まれない
     */
    public static Map<String, Found> find(Path dir, Set<String> keys, List<String> warnings) {
        return find(dir, keys, warnings, null, Map.of());
    }

    /**
     * @param apiKey 復号済み Gemini API キー。空なら加工賃は計算しない
     */
    public static Map<String, Found> find(
            Path dir, Set<String> keys, List<String> warnings, String apiKey, Map<String, String> dispatchOrderByIrai) {
        if (dir == null || keys == null || keys.isEmpty() || !dir.toFile().isDirectory()) {
            return Map.of();
        }
        File[] files = dir.toFile().listFiles(
                (d, name) -> name.endsWith(".xlsm")
                        && !name.startsWith("~$")
                        && !name.equals("加工依頼書入力.xlsm"));
        if (files == null || files.length == 0) {
            warn(warnings, "Excel 依頼書原本が見つかりません: " + dir);
            return Map.of();
        }
        Map<String, Found> found = new LinkedHashMap<>();
        for (File file : files) {
            if (found.size() >= keys.size()) {
                break;
            }
            try (Workbook wb = PoiWorkbookOpener.open(file)) {
                Sheet index = wb.getSheet(RequestFormOriginalIndexSheetLayout.SHEET_NAME);
                if (index == null) {
                    continue;
                }
                Path path = file.toPath().toAbsolutePath().normalize();
                for (RequestFormOriginalIndexSheetReader.IndexEntry entry :
                        RequestFormOriginalIndexSheetReader.read(index).values()) {
                    String key = JuchuTransferValueNormalizer.normalizeKey(entry.iraiNo());
                    if (key.isEmpty() || !keys.contains(key) || found.containsKey(key)) {
                        continue;
                    }
                    String sheetName = matchingSheet(wb, key);
                    String order = dispatchOrderByIrai == null ? "" : dispatchOrderByIrai.getOrDefault(key, "");
                    RequestFormOriginalFee.Result fee =
                            feeOf(wb, sheetName, apiKey, order, entry.iraiNo(), warnings);
                    found.put(key, new Found(entry.iraiNo(), path, sheetName, fee));
                }
            } catch (Exception ex) {
                warn(warnings, "原本目次読込エラー " + file.getName() + ": " + ex.getMessage());
            }
        }
        return Map.copyOf(found);
    }

    private static RequestFormOriginalFee.Result feeOf(
            Workbook wb,
            String sheetName,
            String apiKey,
            String dispatchProcessOrder,
            String iraiNo,
            List<String> warnings) {
        if (sheetName == null || sheetName.isBlank() || apiKey == null || apiKey.isBlank()) {
            return null;
        }
        try {
            return RequestFormOriginalFee.fromSheet(wb.getSheet(sheetName), apiKey, dispatchProcessOrder);
        } catch (InterruptedException ex) {
            Thread.currentThread().interrupt();
            warn(warnings, iraiNo + " の加工賃計算を中断しました");
            return null;
        } catch (Exception ex) {
            String msg = ex.getMessage() == null ? ex.toString() : ex.getMessage();
            warn(warnings, iraiNo + " の加工賃を Gemini で計算できません: " + msg);
            return null;
        }
    }

    private static String matchingSheet(Workbook wb, String key) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            if (key.equals(JuchuTransferValueNormalizer.normalizeKey(name))) {
                return name;
            }
        }
        return "";
    }

    private static void warn(List<String> warnings, String message) {
        if (warnings != null) {
            warnings.add(message);
        }
    }
}
