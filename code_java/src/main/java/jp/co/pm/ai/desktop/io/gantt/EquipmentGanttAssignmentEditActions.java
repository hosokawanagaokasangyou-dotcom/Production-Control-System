package jp.co.pm.ai.desktop.io.gantt;

/** 担当割当編集モードでの追加・削除 UI コールバック。 */
public interface EquipmentGanttAssignmentEditActions {

    void onAddPersonRequested(String barId, double screenX, double screenY);

    void onRemovePersonRequested(String barId, String memberKey, double screenX, double screenY);
}
