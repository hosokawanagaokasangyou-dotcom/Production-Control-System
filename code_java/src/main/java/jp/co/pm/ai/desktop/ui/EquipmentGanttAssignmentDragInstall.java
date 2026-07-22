package jp.co.pm.ai.desktop.ui;

import javafx.scene.Cursor;
import javafx.scene.Node;
import javafx.scene.control.ContextMenu;
import javafx.scene.control.MenuItem;
import javafx.scene.input.ClipboardContent;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.MouseEvent;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;

import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentDragPayload;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentDropHandler;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentDropTarget;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentEditActions;
import jp.co.pm.ai.desktop.io.gantt.EquipmentGanttAssignmentInteraction;
/** 設備ガント担当割当編集のドラッグ＆ドロップ UI 配線。 */
final class EquipmentGanttAssignmentDragInstall {

    private EquipmentGanttAssignmentDragInstall() {}

    static void installBadgeSource(
            StackPane badge,
            String barId,
            String memberKey,
            EquipmentGanttAssignmentInteraction interaction) {
        if (badge == null
                || barId == null
                || barId.isBlank()
                || memberKey == null
                || memberKey.isBlank()
                || interaction == null
                || !interaction.active()) {
            return;
        }
        EquipmentGanttAssignmentDropHandler dropHandler = interaction.dropHandler();
        badge.setMouseTransparent(false);
        applyOpenHandCursor(badge);
        Runnable startDrag =
                () -> {
                    Dragboard db = badge.startDragAndDrop(TransferMode.MOVE);
                    ClipboardContent content = new ClipboardContent();
                    content.putString(
                            new EquipmentGanttAssignmentDragPayload(barId, memberKey).encode());
                    db.setContent(content);
                };
        badge.setOnDragDetected(
                e -> {
                    startDrag.run();
                    e.consume();
                });
        for (Node child : badge.getChildrenUnmodifiable()) {
            child.addEventHandler(
                    MouseEvent.DRAG_DETECTED,
                    e -> {
                        startDrag.run();
                        e.consume();
                    });
        }
        installDropTarget(badge, barId, memberKey, dropHandler);
        installBadgeContextMenu(badge, barId, memberKey, interaction.editActions());
    }

    static void installBarBodyDropTarget(
            Region zone,
            String barId,
            EquipmentGanttAssignmentInteraction interaction) {
        if (zone == null || barId == null || barId.isBlank() || interaction == null) {
            return;
        }
        installDropTarget(zone, barId, "", interaction.dropHandler());
        installBarContextMenu(zone, barId, interaction.editActions());
    }

    private static void installBadgeContextMenu(
            StackPane badge,
            String barId,
            String memberKey,
            EquipmentGanttAssignmentEditActions editActions) {
        if (editActions == null) {
            return;
        }
        ContextMenu menu = new ContextMenu();
        MenuItem remove = new MenuItem("担当を削除");
        remove.setOnAction(
                e ->
                        editActions.onRemovePersonRequested(
                                barId,
                                memberKey,
                                menuAnchorX(badge),
                                menuAnchorY(badge)));
        menu.getItems().add(remove);
        badge.addEventHandler(
                javafx.scene.input.ContextMenuEvent.CONTEXT_MENU_REQUESTED,
                e -> {
                    menu.show(badge, e.getScreenX(), e.getScreenY());
                    e.consume();
                });
    }

    private static void installBarContextMenu(
            Region zone, String barId, EquipmentGanttAssignmentEditActions editActions) {
        if (editActions == null) {
            return;
        }
        ContextMenu menu = new ContextMenu();
        MenuItem add = new MenuItem("担当を追加…");
        add.setOnAction(
                e ->
                        editActions.onAddPersonRequested(
                                barId, menuAnchorX(zone), menuAnchorY(zone)));
        menu.getItems().add(add);
        zone.addEventHandler(
                javafx.scene.input.ContextMenuEvent.CONTEXT_MENU_REQUESTED,
                e -> {
                    menu.show(zone, e.getScreenX(), e.getScreenY());
                    e.consume();
                });
    }

    private static double menuAnchorX(Node node) {
        return node.localToScreen(node.getBoundsInLocal()).getMinX();
    }

    private static double menuAnchorY(Node node) {
        return node.localToScreen(node.getBoundsInLocal()).getMinY();
    }

    private static void installDropTarget(
            Region node,
            String barId,
            String memberKey,
            EquipmentGanttAssignmentDropHandler dropHandler) {
        if (dropHandler == null) {
            return;
        }
        node.setOnDragOver(
                e -> {
                    if (canAccept(e)) {
                        e.acceptTransferModes(TransferMode.MOVE);
                    }
                    e.consume();
                });
        node.setOnDragEntered(
                e -> {
                    if (canAccept(e)) {
                        node.setOpacity(0.82);
                    }
                    e.consume();
                });
        node.setOnDragExited(
                e -> {
                    node.setOpacity(1.0);
                    e.consume();
                });
        node.setOnDragDropped(
                e -> {
                    node.setOpacity(1.0);
                    EquipmentGanttAssignmentDragPayload source = readSource(e);
                    if (source == null) {
                        e.setDropCompleted(false);
                        e.consume();
                        return;
                    }
                    boolean ok =
                            dropHandler.onDrop(
                                    source, new EquipmentGanttAssignmentDropTarget(barId, memberKey));
                    e.setDropCompleted(ok);
                    e.consume();
                });
    }

    private static boolean canAccept(DragEvent e) {
        EquipmentGanttAssignmentDragPayload source = readSource(e);
        return source != null && e.getDragboard().hasString();
    }

    private static EquipmentGanttAssignmentDragPayload readSource(DragEvent e) {
        if (e == null || e.getDragboard() == null || !e.getDragboard().hasString()) {
            return null;
        }
        return EquipmentGanttAssignmentDragPayload.decode(e.getDragboard().getString());
    }

    /** ピル内 {@link javafx.scene.control.Label} 上でも OPEN_HAND になるよう子へ伝播する。 */
    static void applyOpenHandCursor(StackPane badge) {
        if (badge == null) {
            return;
        }
        badge.setCursor(Cursor.OPEN_HAND);
        badge.setOnMouseEntered(e -> badge.setCursor(Cursor.OPEN_HAND));
        for (Node child : badge.getChildrenUnmodifiable()) {
            child.setMouseTransparent(false);
            child.setCursor(Cursor.OPEN_HAND);
            child.setOnMouseEntered(ev -> child.setCursor(Cursor.OPEN_HAND));
        }
    }
}
