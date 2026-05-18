package jp.co.pm.ai.desktop.dispatch;

import java.util.List;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

/**
 * 納期管理ビュー・サマリ Excel の (段階3前) 表示用数量。
 *
 * <p>換算数量 &lt; (原反)ロール単位長さ のとき、アラジン再入力値は換算数量（配台の当日配台数量 m ではない）。
 * 配台タイムライン・加工時間算定は従来どおり原反ロール長ベースの数量を使う。
 */
public final class AladdinSystemDispatchDisplayQty {

    private static final double EPS = 1e-12;

    private AladdinSystemDispatchDisplayQty() {}

    /** 1 暦日分の表示 m（換算数量未満×原反長のとき {@code min(配台m, 換算数量)}）。 */
    public static double displayQtyForDay(
            double dispatchQtyM, double qtyConvM, double rawRollM) {
        double dq = Math.max(0.0, dispatchQtyM);
        if (dq <= EPS) {
            return 0.0;
        }
        double q = Math.max(0.0, qtyConvM);
        double r = Math.max(0.0, rawRollM);
        if (q > EPS && r > EPS && q + 1e-9 < r) {
            return Math.min(dq, q);
        }
        return dq;
    }

    /**
     * 依頼NO 行で暦日順に配分するときの表示 m と残り換算数量キャップ。
     *
     * @param remainingConvCap 換算数量の残り（null なら {@link #displayQtyForDay} と同趣旨でキャップなし）
     */
    public static DayDisplay allocateDay(
            double dispatchQtyM,
            double qtyConvM,
            double rawRollM,
            Double remainingConvCap) {
        double dq = Math.max(0.0, dispatchQtyM);
        if (dq <= EPS) {
            return new DayDisplay(0.0, remainingConvCap);
        }
        double q = Math.max(0.0, qtyConvM);
        double r = Math.max(0.0, rawRollM);
        if (!(q > EPS && r > EPS && q + 1e-9 < r)) {
            return new DayDisplay(dq, remainingConvCap);
        }
        if (remainingConvCap != null) {
            double cap = Math.max(0.0, remainingConvCap);
            double show = Math.min(dq, cap);
            return new DayDisplay(show, Math.max(0.0, cap - show));
        }
        return new DayDisplay(Math.min(dq, q), null);
    }

    public record DayDisplay(double displayM, Double remainingConvCap) {}

    /** 結果_配台表の行から (換算数量, 原反ロール長 m) を解決する。 */
    public static TaskQtyContext contextFromDispatchRow(
            List<String> headers, List<String> row, Stage2RollUnitLengthTables tables) {
        int iQty = headers.indexOf("換算数量");
        int iUsed = headers.indexOf("使用原反");
        double qtyConv =
                iQty >= 0 && row != null && iQty < row.size()
                        ? Stage2RollUnitLengthTables.parseFloatSafe(cellAt(row, iQty), 0.0)
                        : 0.0;
        String usedRaw =
                iUsed >= 0 && row != null && iUsed < row.size() ? cellAt(row, iUsed) : "";
        return new TaskQtyContext(qtyConv, resolveRawRollUnitM(usedRaw, tables));
    }

    public static double resolveRawRollUnitM(String usedRaw, Stage2RollUnitLengthTables tables) {
        if (tables != null) {
            Optional<Double> tab = tables.lookupByUsedRaw(usedRaw);
            if (tab.isPresent() && tab.get() > EPS) {
                return tab.get();
            }
        }
        double dim =
                Stage2RollUnitLengthTables.inferFromProductDimensions(
                        usedRaw != null ? usedRaw : "", 0.0);
        return dim > EPS ? dim : 0.0;
    }

    /** {@code mk + "\\t" + tid} キーで最初に見つかった行のコンテキスト。 */
    public static Map<String, TaskQtyContext> buildContextByMachineAndTaskId(
            List<String> headers, List<List<String>> rows, Stage2RollUnitLengthTables tables) {
        int iMk = headers.indexOf("機械名");
        int iTid = headers.indexOf("依頼NO");
        if (iMk < 0 || iTid < 0 || rows == null) {
            return Map.of();
        }
        java.util.LinkedHashMap<String, TaskQtyContext> out = new java.util.LinkedHashMap<>();
        for (List<String> row : rows) {
            String mk = normalizeEquipmentMatchKey(cellAt(row, iMk));
            String tid = cellAt(row, iTid).strip();
            if (mk.isEmpty() || tid.isEmpty()) {
                continue;
            }
            String key = mk + "\t" + tid;
            out.putIfAbsent(key, contextFromDispatchRow(headers, row, tables));
        }
        return Map.copyOf(out);
    }

    public record TaskQtyContext(double qtyConvM, double rawRollM) {
        public boolean usesConvertedQtyForAladdinDisplay() {
            return qtyConvM > EPS && rawRollM > EPS && qtyConvM + 1e-9 < rawRollM;
        }
    }

    private static String cellAt(List<String> row, int idx) {
        if (row == null || idx < 0 || idx >= row.size() || row.get(idx) == null) {
            return "";
        }
        return row.get(idx);
    }

    private static String normalizeEquipmentMatchKey(String val) {
        if (val == null || val.isBlank()) {
            return "";
        }
        String t = java.text.Normalizer.normalize(val, java.text.Normalizer.Form.NFKC);
        t = t.replace('\u00a0', ' ').replace('\u3000', ' ');
        t = t.replaceAll("[\u200b\u200c\u200d\ufeff]", "");
        return t.replaceAll("\\s+", " ").strip();
    }
}
