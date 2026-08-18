package jp.co.pm.ai.desktop.ui;

import java.nio.file.Path;
import java.util.List;
import java.util.function.Consumer;

import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Window;

/** ソース拡張子不正時に表を暗転し、エラーメッセージと削除操作を重ねる。 */
public final class SourceExtensionErrorOverlay {

    public static final String STYLE_CLASS = "source-extension-error-overlay";

    private SourceExtensionErrorOverlay() {}

    public static void show(StackPane host, String message) {
        show(host, message, List.of(), null, null);
    }

    /**
     * @param deletablePaths ユーザー確認後に削除可能な不正ファイル（空なら削除ボタンなし）
     * @param owner 確認ダイアログの親
     * @param onDeleted 削除成功後（1件以上）
     */
    public static void show(
            StackPane host,
            String message,
            List<Path> deletablePaths,
            Window owner,
            Consumer<List<Path>> onDeleted) {
        clear(host);
        if (host == null) {
            return;
        }
        Label label = new Label(message != null ? message : "ソースファイルの拡張子が不正です。");
        label.setWrapText(true);
        label.setMaxWidth(520);
        label.getStyleClass().add("source-extension-error-overlay-label");
        VBox box = new VBox(12);
        box.setAlignment(Pos.CENTER);
        box.getStyleClass().add(STYLE_CLASS);
        box.setMouseTransparent(false);
        box.getChildren().add(label);
        List<Path> targets = deletablePaths != null ? List.copyOf(deletablePaths) : List.of();
        if (!targets.isEmpty()) {
            Button deleteButton = new Button("不正拡張子ファイルを削除…");
            deleteButton.getStyleClass().add("source-extension-error-overlay-delete");
            deleteButton.setOnAction(
                    e ->
                            SourceExtensionMismatchDeletion.confirmAndDelete(
                                    owner,
                                    targets,
                                    deleted -> {
                                        if (onDeleted != null) {
                                            onDeleted.accept(deleted);
                                        }
                                    },
                                    null,
                                    err -> showDeleteError(owner, err)));
            box.getChildren().add(deleteButton);
        }
        StackPane.setAlignment(box, Pos.CENTER);
        host.getChildren().add(box);
        box.toFront();
    }

    private static void showDeleteError(Window owner, String message) {
        javafx.scene.control.Alert alert =
                new javafx.scene.control.Alert(
                        javafx.scene.control.Alert.AlertType.ERROR,
                        message != null ? message : "削除に失敗しました。");
        alert.setTitle("不正拡張子ファイルの削除");
        alert.setHeaderText(null);
        if (owner != null) {
            alert.initOwner(owner);
        }
        alert.showAndWait();
    }

    /** 同一ホスト上で stale 等の後ろに隠れた場合に最前面へ。 */
    public static void toFrontIfShowing(StackPane host) {
        if (host == null) {
            return;
        }
        host.getChildren().stream()
                .filter(
                        n ->
                                n.getStyleClass() != null
                                        && n.getStyleClass().contains(STYLE_CLASS))
                .forEach(javafx.scene.Node::toFront);
    }

    public static void clear(StackPane host) {
        if (host == null) {
            return;
        }
        host.getChildren()
                .removeIf(
                        n ->
                                n.getStyleClass() != null
                                        && n.getStyleClass().contains(STYLE_CLASS));
    }

    public static boolean isShowing(StackPane host) {
        if (host == null) {
            return false;
        }
        return host.getChildren().stream()
                .anyMatch(
                        n ->
                                n.getStyleClass() != null
                                        && n.getStyleClass().contains(STYLE_CLASS));
    }
}
