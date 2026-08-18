package jp.co.pm.ai.desktop.ui;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Consumer;

import javafx.scene.control.Alert;
import javafx.scene.control.ButtonBar;
import javafx.scene.control.ButtonType;
import javafx.stage.Window;

/**
 * 拡張子／ファイル名不一致のソースファイルを、ユーザー確認後に削除する。
 */
public final class SourceExtensionMismatchDeletion {

    private SourceExtensionMismatchDeletion() {}

    /**
     * 確認ダイアログを出し、承認時のみ通常ファイルを削除する。
     *
     * @param owner ダイアログ親（null 可）
     * @param paths 削除候補（ディレクトリはスキップ）
     * @param onDeleted 削除に成功したパス一覧（1件以上）
     * @param onCancelledOrEmpty キャンセル／対象なし時（null 可）
     * @param onError 失敗メッセージ（null 可）
     */
    public static void confirmAndDelete(
            Window owner,
            List<Path> paths,
            Consumer<List<Path>> onDeleted,
            Runnable onCancelledOrEmpty,
            Consumer<String> onError) {
        List<Path> targets = regularFilesOnly(paths);
        if (targets.isEmpty()) {
            if (onCancelledOrEmpty != null) {
                onCancelledOrEmpty.run();
            }
            return;
        }
        Alert alert = new Alert(Alert.AlertType.CONFIRMATION);
        alert.setTitle("不正拡張子ファイルの削除");
        alert.setHeaderText("次のファイルを削除します。よろしいですか？");
        alert.setContentText(formatPathList(targets) + "\n\nこの操作は取り消せません。");
        if (owner != null) {
            alert.initOwner(owner);
        }
        ButtonType delete = new ButtonType("削除する", ButtonBar.ButtonData.OK_DONE);
        ButtonType cancel = new ButtonType("キャンセル", ButtonBar.ButtonData.CANCEL_CLOSE);
        alert.getButtonTypes().setAll(delete, cancel);
        Optional<ButtonType> choice = alert.showAndWait();
        if (choice.isEmpty() || choice.get() != delete) {
            if (onCancelledOrEmpty != null) {
                onCancelledOrEmpty.run();
            }
            return;
        }
        List<Path> deleted = new ArrayList<>();
        List<String> errors = new ArrayList<>();
        for (Path path : targets) {
            try {
                if (!Files.isRegularFile(path)) {
                    errors.add("ファイルがありません: " + path);
                    continue;
                }
                Files.delete(path);
                deleted.add(path);
            } catch (IOException ex) {
                String msg =
                        ex.getMessage() != null && !ex.getMessage().isBlank()
                                ? ex.getMessage()
                                : ex.toString();
                errors.add(path + " → " + msg);
            }
        }
        if (!errors.isEmpty() && onError != null) {
            onError.accept(String.join("\n", errors));
        }
        if (!deleted.isEmpty() && onDeleted != null) {
            onDeleted.accept(List.copyOf(deleted));
        } else if (deleted.isEmpty() && onCancelledOrEmpty != null) {
            onCancelledOrEmpty.run();
        }
    }

    private static List<Path> regularFilesOnly(List<Path> paths) {
        if (paths == null || paths.isEmpty()) {
            return List.of();
        }
        List<Path> out = new ArrayList<>();
        for (Path raw : paths) {
            if (raw == null) {
                continue;
            }
            Path abs = raw.toAbsolutePath().normalize();
            try {
                if (Files.isRegularFile(abs) && !out.contains(abs)) {
                    out.add(abs);
                }
            } catch (Exception ignored) {
                // UNC 等で判定失敗したパスは候補に含めない
            }
        }
        return List.copyOf(out);
    }

    private static String formatPathList(List<Path> paths) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < paths.size(); i++) {
            if (i > 0) {
                sb.append('\n');
            }
            sb.append("・").append(Objects.toString(paths.get(i), ""));
        }
        return sb.toString();
    }
}
