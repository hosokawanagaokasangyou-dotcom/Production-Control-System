package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import javafx.scene.control.TextInputDialog;
import javafx.stage.DirectoryChooser;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

/** 後加工検査表フォルダ選択。DirectoryChooser の前に kensa 入力を必須にする。 */
public final class InspectionSheetDirPicker {

    public static final String KENSA_TOKEN = "kensa";

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
        File selected = dc.showDialog(owner);
        return Optional.ofNullable(selected);
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

    public static Optional<Path> initialDirectory(FactorySite factorySiteHint) {
        for (Path candidate : initialDirectoryCandidates(factorySiteHint)) {
            if (Files.isDirectory(candidate)) {
                return Optional.of(candidate.toAbsolutePath().normalize());
            }
        }
        return Optional.empty();
    }

    public static List<Path> initialDirectoryCandidates(FactorySite factorySiteHint) {
        Path box = Path.of(System.getProperty("user.home"), "Box");
        Path nagaoka = box.resolve("長岡産業").resolve("後加工検査表");
        List<Path> out = new ArrayList<>();
        if (factorySiteHint != null && factorySiteHint != FactorySite.RDP_LAUNCHER) {
            String leaf = AppPaths.inspectionSheetDirLeafForFactory(factorySiteHint);
            if (!leaf.isEmpty()) {
                out.add(nagaoka.resolve(leaf));
            }
        }
        out.add(nagaoka);
        out.add(box);
        return List.copyOf(out);
    }
}
