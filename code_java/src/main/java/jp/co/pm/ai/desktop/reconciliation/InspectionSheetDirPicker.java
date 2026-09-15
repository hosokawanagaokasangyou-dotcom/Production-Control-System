package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;

import javafx.event.ActionEvent;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonBar.ButtonData;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

/** 後加工検査表フォルダ選択。DirectoryChooser の前に kensa 入力を必須にする。 */
public final class InspectionSheetDirPicker {

    public static final String KENSA_TOKEN = "kensa";

    public static final String KENSA_CONFIRM_TITLE = "後加工検査表フォルダ — 確認";

    public static final String KENSA_CONFIRM_HEADER = "後加工検査表のフォルダを選びます";

    public static final String KENSA_CONFIRM_WARNING =
            "依頼書原本フォルダではありません。\n"
                    + "Box の「後加工検査表」配下（国分／湖南）を指定してください。\n"
                    + "原本・依頼書のフォルダを選ぶと索引が壊れます。";

    public static final String KENSA_CONFIRM_TOKEN_HINT =
            "確認のため、下に表示した kensa を入力してください。";

    public static final String KENSA_CONFIRM_MISMATCH =
            "「kensa」と一致しません。依頼書原本フォルダではありません。";

    public static final int KENSA_TOKEN_FONT_SIZE_PX = 36;

    public static final int KENSA_WARNING_FONT_SIZE_PX = 16;

    /** クリーム背景向け。親ウィンドウのダークテーマ文字色は使わない。 */
    public static final String KENSA_CONFIRM_TEXT_FILL = "#1a1a1a";

    public static final String KENSA_CONFIRM_WARNING_FILL = "#6b140c";

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
        Dialog<Boolean> dialog = new Dialog<>();
        if (owner != null) {
            dialog.initOwner(owner);
        }
        dialog.setTitle(KENSA_CONFIRM_TITLE);
        dialog.setHeaderText(null);

        Label header = new Label(KENSA_CONFIRM_HEADER);
        header.setWrapText(true);
        header.setStyle(kensaLabelStyle(20, true, KENSA_CONFIRM_TEXT_FILL));

        Alert iconSource = new Alert(AlertType.WARNING);
        Node graphic = iconSource.getGraphic();
        HBox headerRow = new HBox(12);
        headerRow.setAlignment(Pos.CENTER_LEFT);
        if (graphic != null) {
            headerRow.getChildren().add(graphic);
        }
        headerRow.getChildren().add(header);

        Label warning = new Label(KENSA_CONFIRM_WARNING);
        warning.setWrapText(true);
        warning.setStyle(
                kensaLabelStyle(KENSA_WARNING_FONT_SIZE_PX, true, KENSA_CONFIRM_WARNING_FILL));

        Label tokenHint = new Label(KENSA_CONFIRM_TOKEN_HINT);
        tokenHint.setWrapText(true);
        tokenHint.setStyle(kensaLabelStyle(15, true, KENSA_CONFIRM_TEXT_FILL));

        Label token = new Label(KENSA_TOKEN);
        token.setStyle(
                kensaLabelStyle(KENSA_TOKEN_FONT_SIZE_PX, true, KENSA_CONFIRM_TEXT_FILL)
                        + " -fx-font-family: Consolas, 'Courier New', monospace;");

        TextField input = new TextField();
        input.setPromptText(KENSA_TOKEN);
        input.setStyle(
                "-fx-font-size: 22px; -fx-font-weight: bold; -fx-text-fill: "
                        + KENSA_CONFIRM_TEXT_FILL
                        + " !important; -fx-prompt-text-fill: #4a4a4a !important;"
                        + " -fx-control-inner-background: #ffffff; -fx-background-color: #ffffff;");
        input.setPrefColumnCount(12);

        Label mismatch = new Label();
        mismatch.setWrapText(true);
        mismatch.setStyle(kensaLabelStyle(14, true, "#b00020"));

        VBox body = new VBox(12, headerRow, warning, tokenHint, token, input, mismatch);
        body.setPadding(new Insets(4, 0, 0, 0));

        dialog.getDialogPane().setContent(body);
        dialog.getDialogPane().setPrefWidth(580);
        dialog.getDialogPane().setStyle("-fx-background-color: #fff8e1;");

        ButtonType proceed = new ButtonType("フォルダを選ぶ", ButtonData.OK_DONE);
        ButtonType cancel = new ButtonType("キャンセル", ButtonData.CANCEL_CLOSE);
        dialog.getDialogPane().getButtonTypes().setAll(proceed, cancel);

        Node okBtn = dialog.getDialogPane().lookupButton(proceed);
        okBtn.setDisable(true);
        input.textProperty()
                .addListener(
                        (obs, old, now) -> {
                            boolean ok = matchesKensa(now);
                            okBtn.setDisable(!ok);
                            if (ok) {
                                mismatch.setText("");
                            }
                        });
        okBtn.addEventFilter(
                ActionEvent.ACTION,
                e -> {
                    if (!matchesKensa(input.getText())) {
                        e.consume();
                        mismatch.setText(KENSA_CONFIRM_MISMATCH);
                        input.requestFocus();
                    }
                });

        dialog.setOnShown(e -> input.requestFocus());
        dialog.setResultConverter(bt -> bt == proceed);
        return dialog.showAndWait().orElse(false);
    }

    private static String kensaLabelStyle(int fontPx, boolean bold, String fill) {
        return "-fx-font-size: "
                + fontPx
                + "px;"
                + (bold ? " -fx-font-weight: bold;" : "")
                + " -fx-text-fill: "
                + fill
                + " !important;";
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
