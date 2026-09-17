package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 検出表用。③月次実績に①の対象月データがあるか。無ければ備考警告。
 */
public final class Source3TargetMonthCheck {

    public static final String NOTE_CSS = "pm-kouchin-note-warn";

    private record CacheEntry(long mtime, String ym, String customer, String warehouse, String note) {}

    private static final ConcurrentHashMap<String, CacheEntry> CACHE = new ConcurrentHashMap<>();

    private Source3TargetMonthCheck() {}

    public static String warningNote(YearMonthKey targetYm, String detail) {
        String ym = targetYm == null ? "(不明)" : targetYm.gatsudoLabel();
        String extra = detail == null || detail.isBlank() ? "" : "（" + detail + "）";
        return "対象月（" + ym + "）のデータがありません" + extra;
    }

    public static boolean isWarning(String note) {
        return note != null && note.contains("対象月") && note.contains("データがありません");
    }

    public static String noteIfUncovered(Path file, YearMonthKey targetYm, FactoryProfile profile) {
        String customer = profile == null ? "" : String.valueOf(profile.customer3());
        String warehouse = profile == null ? "" : String.valueOf(profile.warehouse3Contains());
        String ymKey = targetYm == null ? "" : targetYm.gatsudoLabel();
        if (file == null || !Files.isRegularFile(file)) {
            return warningNote(targetYm, "ファイルなし");
        }
        long mtime = 0L;
        try {
            mtime = Files.getLastModifiedTime(file).toMillis();
        } catch (Exception ignored) {
        }
        String cacheKey = file.toAbsolutePath().normalize().toString();
        CacheEntry cached = CACHE.get(cacheKey);
        if (cached != null
                && cached.mtime() == mtime
                && cached.ym().equals(ymKey)
                && cached.customer().equals(customer)
                && cached.warehouse().equals(warehouse)) {
            return cached.note();
        }
        String note = compute(file, targetYm, profile);
        CACHE.put(cacheKey, new CacheEntry(mtime, ymKey, customer, warehouse, note));
        return note;
    }

    private static String compute(Path file, YearMonthKey targetYm, FactoryProfile profile) {
        try {
            String customer = profile == null ? null : profile.customer3();
            String warehouse = profile == null ? null : profile.warehouse3Contains();
            AladdinData data = AladdinReader.read(file, customer, warehouse);
            if (AladdinData.coversTargetMonth(data, targetYm)) {
                return "";
            }
            String detail = data.count() == 0
                    ? "加工金額行が0件"
                    : "シート対象年月=" + (data.taishoYm() == null ? "(不明)" : data.taishoYm().ymLabel());
            return warningNote(targetYm, detail);
        } catch (RuntimeException e) {
            String msg = e.getMessage();
            if (msg != null && msg.contains("加工金額")) {
                return warningNote(targetYm, "加工金額行が0件");
            }
            return warningNote(targetYm, msg == null || msg.isBlank() ? "読み取り失敗" : msg);
        }
    }
}
