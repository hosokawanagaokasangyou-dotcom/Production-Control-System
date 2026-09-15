package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;

import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.TextInputDialog;
import javafx.stage.DirectoryChooser;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

/** 後加工検査表フォルダ選択。DirectoryChooser の前に kensa 入力を必須にする。 */
public final class InspectionSheetDirPicker {

    public static final String KENSA_TOKEN = "kensa";

    /** 誤指定判定で見るファイル数の上限。見つかった時点で打ち切る。 */
    static final int MAX_PROBE_ENTRIES = 8_000;

    private static final int MAX_WALK_DEPTH = 16;

    private InspectionSheetDirPicker() {}

    public static boolean matchesKensa(String input) {
        return input != null && KENSA_TOKEN.equals(input.strip());
    }

    public static Optional<File> pick(Window owner, FactorySite factorySiteHint) {
        if (!confirmKensa(owner)) {
            return Optional.empty();
        }
        DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle("後加工検査表フォルダ");
        initialDirectory(factorySiteHint).ifPresent(p -> dc.setInitialDirectory(p.toFile()));
        while (true) {
            File selected = dc.showDialog(owner);
            if (selected == null) {
                return Optional.empty();
            }
            Optional<String> err = validateInspectionSheetDir(selected.toPath());
            if (err.isEmpty()) {
                return Optional.of(selected);
            }
            showRejectAlert(owner, err.get());
            dc.setInitialDirectory(selected);
        }
    }

    public static boolean confirmKensa(Window owner) {
        TextInputDialog dialog = new TextInputDialog();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle("後加工検査表フォルダ");
        dialog.setHeaderText(null);
        dialog.setContentText(
                "後加工検査表フォルダを選びます（依頼書原本フォルダではありません）。"
                        + "確認のため kensa と入力してください。キャンセルでスキップします。");
        Optional<String> ans = dialog.showAndWait();
        return ans.isPresent() && matchesKensa(ans.get());
    }

    /**
     * 空なら妥当。値が入っているときは、後加工検査表の Excel が1件以上あるディレクトリであること。
     */
    public static Optional<String> validateInspectionSheetDir(Path dir) {
        if (dir == null) {
            return Optional.of("フォルダが指定されていません。");
        }
        Path abs;
        try {
            abs = dir.toAbsolutePath().normalize();
        } catch (RuntimeException ex) {
            return Optional.of("パスが無効です。");
        }
        if (!Files.isDirectory(abs)) {
            return Optional.of("フォルダが存在しません:\n" + abs);
        }
        if (containsInspectionSheets(abs)) {
            return Optional.empty();
        }
        String leaf = abs.getFileName() != null ? abs.getFileName().toString() : abs.toString();
        return Optional.of(
                "「"
                        + leaf
                        + "」に後加工検査表の Excel が見つかりません。\n"
                        + "依頼書原本フォルダなど、検査表が入っていない場所は指定できません。");
    }

    public static boolean containsInspectionSheets(Path dir) {
        if (dir == null || !Files.isDirectory(dir)) {
            return false;
        }
        try (Stream<Path> walk = Files.walk(dir, MAX_WALK_DEPTH)) {
            int seen = 0;
            for (Path p : (Iterable<Path>) walk::iterator) {
                if (!Files.isRegularFile(p)) {
                    continue;
                }
                seen++;
                if (looksLikeInspectionSheetFile(p)) {
                    return true;
                }
                if (seen >= MAX_PROBE_ENTRIES) {
                    break;
                }
            }
            return false;
        } catch (IOException ex) {
            return false;
        }
    }

    static boolean looksLikeInspectionSheetFile(Path path) {
        if (!InspectionSheetIndexScanner.isInspectionExcel(path)) {
            return false;
        }
        Path namePath = path.getFileName();
        if (namePath == null) {
            return false;
        }
        String name = namePath.toString();
        String half = InspectionSheetIraiNo.toHalfWidth(name);
        if (half.contains("加工依頼書")) {
            return false;
        }
        if (InspectionSheetIraiNo.extractFromFileName(name).isPresent()) {
            return true;
        }
        if (half.contains("完了") || half.contains("検査") || half.contains("SEC済")) {
            return true;
        }
        return underInspectionSheetTree(path);
    }

    private static boolean underInspectionSheetTree(Path path) {
        Path cur = path.getParent();
        int up = 0;
        while (cur != null && up < 8) {
            Path fn = cur.getFileName();
            if (fn != null) {
                String n = InspectionSheetIraiNo.toHalfWidth(fn.toString());
                if (n.contains("後加工検査表") || n.contains("検査表")) {
                    return true;
                }
            }
            cur = cur.getParent();
            up++;
        }
        return false;
    }

    static void showRejectAlert(Window owner, String message) {
        Alert a = new Alert(AlertType.WARNING);
        if (owner != null) {
            a.initOwner(owner);
        }
        a.setTitle("後加工検査表フォルダ");
        a.setHeaderText(null);
        a.setContentText(message);
        a.showAndWait();
    }

    public static Optional<Path> initialDirectory(FactorySite factorySiteHint) {
        for (Path candidate : initialDirectoryCandidates(factorySiteHint)) {
            if (Files.isDirectory(candidate)) {
                return Optional.of(candidate.toAbsolutePath().normalize());
            }
        }
        return Optional.empty();
    }

    public static List<Path> initialDirectoryCandidates(FactorySite factorySiteHint) {
        List<Path> out = new ArrayList<>();
        FactorySite site =
                factorySiteHint != null && factorySiteHint != FactorySite.RDP_LAUNCHER
                        ? factorySiteHint
                        : FactorySite.KONAN;
        String factoryDefault = AppPaths.defaultInspectionSheetDirForFactory(site);
        if (!factoryDefault.isEmpty()) {
            out.add(Path.of(factoryDefault));
        }
        Path box = Path.of(System.getProperty("user.home"), "Box");
        Path nagaoka = box.resolve("長岡産業").resolve("後加工検査表");
        String leaf = AppPaths.inspectionSheetDirLeafForFactory(site);
        if (!leaf.isEmpty()) {
            Path boxFactory = nagaoka.resolve(leaf);
            if (out.stream().noneMatch(p -> p.equals(boxFactory))) {
                out.add(boxFactory);
            }
        }
        out.add(nagaoka);
        out.add(box);
        return List.copyOf(out);
    }
}
