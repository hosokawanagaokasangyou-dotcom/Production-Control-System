package jp.co.pm.ai.planning.stage2.parity;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

/**
 * 配台計画タブの表（UI）と {@code PM_AI_PLAN_INPUT_PATH} の実体が、段階2が読む解釈で一致するか検証する。
 *
 * <p>見出しの並びがファイルと異なる場合でも、見出し名が一意なら列の並べ替えで照合する。
 */
public final class Stage2PlanInputUiParity {

    private Stage2PlanInputUiParity() {}

    public static Stage2ParityCheckResult compareUiToDisk(
            PlanInputTabularIo.TabularSheet ui, Path planInputPath, String requestedSheetName)
            throws IOException {
        Objects.requireNonNull(ui, "ui");
        Objects.requireNonNull(planInputPath, "planInputPath");
        if (!Files.isRegularFile(planInputPath)) {
            return new Stage2ParityCheckResult(
                    false, "タスク入力ファイルがありません: " + planInputPath);
        }
        PlanInputTabularIo.TabularRead disk =
                PlanInputTabularIo.readWithResolvedSheet(planInputPath, requestedSheetName);
        PlanInputTabularIo.TabularSheet fileTab = disk.tabular();
        if (!headersMultisetEqual(ui.headers(), fileTab.headers())) {
            return new Stage2ParityCheckResult(
                    false,
                    "配台計画タブの見出し集合とファイルが一致しません。\n"
                            + "（先に「保存」するか「再読み」してから同一検証してください）\n\n"
                            + "ファイル: "
                            + planInputPath);
        }
        PlanInputTabularIo.TabularSheet fileAligned = remapFileToUiHeaderOrder(ui.headers(), fileTab);
        if (fileAligned == null) {
            return new Stage2ParityCheckResult(
                    false,
                    "見出しに重複があるため、ファイルとタブの列順を自動照合できません。\n"
                            + "タブで「保存」してファイルの列順と一致させてから再実行してください。\n\n"
                            + "ファイル: "
                            + planInputPath);
        }
        if (!rowsEqual(ui.rows(), fileAligned.rows())) {
            return new Stage2ParityCheckResult(
                    false,
                    "配台計画タブのセル内容とタスク入力ファイルが一致しません（未保存の編集の可能性）。\n"
                            + "「保存」するか「再読み」後に再実行してください。\n\n"
                            + "ファイル: "
                            + planInputPath);
        }
        return new Stage2ParityCheckResult(
                true,
                "配台計画タブの表とタスク入力ファイルの内容が一致しました。\n\n"
                        + planInputPath
                        + "\n解決シート名: "
                        + (disk.resolvedSheetName().isEmpty() ? "（CSV）" : disk.resolvedSheetName()));
    }

    private static boolean headersMultisetEqual(List<String> a, List<String> b) {
        return multiset(normalizeHeaders(a)).equals(multiset(normalizeHeaders(b)));
    }

    private static Map<String, Integer> multiset(List<String> keys) {
        Map<String, Integer> m = new HashMap<>();
        for (String k : keys) {
            m.merge(k, 1, Integer::sum);
        }
        return m;
    }

    private static List<String> normalizeHeaders(List<String> h) {
        List<String> o = new ArrayList<>(h.size());
        for (String x : h) {
            o.add(norm(x));
        }
        return o;
    }

    /**
     * ファイル側の列を UI 見出し順に並べ替えた表を返す。見出し重複時は {@code null}。
     */
    private static PlanInputTabularIo.TabularSheet remapFileToUiHeaderOrder(
            List<String> uiHeaders, PlanInputTabularIo.TabularSheet file) {
        List<String> fh = normalizeHeaders(file.headers());
        List<String> uh = normalizeHeaders(uiHeaders);
        if (uh.size() != fh.size()) {
            return null;
        }
        if (new HashSet<>(uh).size() != uh.size()) {
            return null;
        }
        boolean[] used = new boolean[fh.size()];
        List<Integer> perm = new ArrayList<>(uh.size());
        for (String name : uh) {
            int j = -1;
            for (int i = 0; i < fh.size(); i++) {
                if (!used[i] && fh.get(i).equals(name)) {
                    j = i;
                    break;
                }
            }
            if (j < 0) {
                return null;
            }
            used[j] = true;
            perm.add(j);
        }
        List<List<String>> outRows = new ArrayList<>();
        for (List<String> row : file.rows()) {
            List<String> nr = new ArrayList<>(uiHeaders.size());
            for (int p : perm) {
                nr.add(p < row.size() ? norm(row.get(p)) : "");
            }
            outRows.add(nr);
        }
        return new PlanInputTabularIo.TabularSheet(new ArrayList<>(uiHeaders), outRows);
    }

    private static boolean rowsEqual(List<List<String>> a, List<List<String>> b) {
        if (a.size() != b.size()) {
            return false;
        }
        for (int i = 0; i < a.size(); i++) {
            if (!rowEqual(a.get(i), b.get(i))) {
                return false;
            }
        }
        return true;
    }

    private static boolean rowEqual(List<String> x, List<String> y) {
        int n = Math.max(x != null ? x.size() : 0, y != null ? y.size() : 0);
        for (int c = 0; c < n; c++) {
            String vx = x != null && c < x.size() ? norm(x.get(c)) : "";
            String vy = y != null && c < y.size() ? norm(y.get(c)) : "";
            if (!vx.equals(vy)) {
                return false;
            }
        }
        return true;
    }

    private static String norm(String s) {
        return s == null ? "" : s.strip();
    }
}
