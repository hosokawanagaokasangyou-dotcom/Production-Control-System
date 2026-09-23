package jp.co.pm.ai.desktop.dispatch;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.HBox;

import jp.co.pm.ai.desktop.MainShellController;

/**
 * 各タブに置く同じ選択。表示の正はシェルの {@link DispatchResultSelection} 一つ。
 */
public final class DispatchResultSelectorBar extends HBox {

    private final MainShellController shell;
    private final ComboBox<String> operatorCombo = new ComboBox<>();
    private final ComboBox<String> generationCombo = new ComboBox<>();
    private final Label badge = new Label();
    private final Label catalogNote = new Label();
    private final Button removeButton = new Button("共有から外す");
    private boolean suppress;
    private List<DispatchSnapshotStore.SnapshotRef> catalog = List.of();
    private String catalogError = "";
    private final Map<String, String> operatorDirByLabel = new LinkedHashMap<>();
    private final Map<String, DispatchSnapshotStore.SnapshotRef> refByLabel = new LinkedHashMap<>();

    public DispatchResultSelectorBar(MainShellController shell) {
        super(8);
        this.shell = shell;
        setAlignment(Pos.CENTER_LEFT);
        Label caption = new Label("配台結果");
        operatorCombo.setPrefWidth(160);
        generationCombo.setPrefWidth(280);
        badge.getStyleClass().add("pm-dispatch-result-badge");
        badge.setVisible(false);
        badge.setManaged(false);
        catalogNote.getStyleClass().add("pm-dispatch-result-badge");
        catalogNote.setVisible(false);
        catalogNote.setManaged(false);
        Button refresh = new Button("一覧を更新");
        refresh.setOnAction(e -> shell.refreshDispatchSnapshotCatalogAsync());
        removeButton.setOnAction(e -> shell.deleteSelectedOwnDispatchSnapshotAsync());
        operatorCombo.setOnAction(e -> onOperatorChosen());
        generationCombo.setOnAction(e -> onGenerationChosen());
        operatorCombo.setOnShowing(e -> shell.refreshDispatchSnapshotCatalogAsync());
        generationCombo.setOnShowing(e -> shell.refreshDispatchSnapshotCatalogAsync());
        getChildren().addAll(caption, operatorCombo, generationCombo, refresh, removeButton, badge, catalogNote);
        syncFromSelection();
    }

    public void setCatalog(List<DispatchSnapshotStore.SnapshotRef> refs) {
        setCatalog(refs, catalogError);
    }

    public void setCatalog(List<DispatchSnapshotStore.SnapshotRef> refs, String error) {
        catalog = refs == null ? List.of() : List.copyOf(refs);
        catalogError = error == null ? "" : error.strip();
        syncFromSelection();
    }

    public void syncFromSelection() {
        suppress = true;
        try {
            rebuildOperatorItems();
            DispatchResultSelection selection = shell.dispatchResultSelection();
            if (selection.isLocalLatest()) {
                operatorCombo.getSelectionModel().select(localLabel());
                generationCombo.getItems().setAll("（ローカル最新）");
                generationCombo.getSelectionModel().select(0);
                generationCombo.setDisable(true);
                badge.setVisible(false);
                badge.setManaged(false);
                removeButton.setDisable(true);
            } else {
                String opLabel = operatorLabel(selection.operatorDir());
                if (!operatorCombo.getItems().contains(opLabel)) {
                    operatorCombo.getItems().add(opLabel);
                    operatorDirByLabel.put(opLabel, selection.operatorDir());
                }
                operatorCombo.getSelectionModel().select(opLabel);
                fillGenerations(selection.operatorDir());
                generationCombo.setDisable(false);
                String genLabel = generationLabel(selection);
                if (!generationCombo.getItems().contains(genLabel)) {
                    generationCombo.getItems().add(genLabel);
                }
                generationCombo.getSelectionModel().select(genLabel);
                String text = selection.badgeText();
                badge.setText(text);
                badge.setTooltip(new Tooltip(selection.detailText()));
                badge.setVisible(!text.isBlank());
                badge.setManaged(!text.isBlank());
                removeButton.setDisable(!selection.ownPast());
            }
            boolean failed = !catalogError.isBlank();
            catalogNote.setText(failed ? "一覧失敗: " + catalogError : "");
            catalogNote.setVisible(failed);
            catalogNote.setManaged(failed);
        } catch (RuntimeException ex) {
            shell.appendLog("[dispatch-snapshot] 選択表示を更新できません: " + ex.getMessage());
        } finally {
            suppress = false;
        }
    }

    private void onOperatorChosen() {
        if (suppress) {
            return;
        }
        String label = operatorCombo.getValue();
        if (label == null || localLabel().equals(label)) {
            shell.useLocalLatestDispatchResult();
            return;
        }
        String dir = operatorDirByLabel.get(label);
        DispatchSnapshotStore.SnapshotRef newest = newestOf(dir);
        if (newest == null) {
            syncFromSelection();
            return;
        }
        shell.selectDispatchSnapshot(newest);
    }

    private void onGenerationChosen() {
        if (suppress || generationCombo.isDisable()) {
            return;
        }
        String label = generationCombo.getValue();
        DispatchSnapshotStore.SnapshotRef ref = refByLabel.get(label);
        if (ref == null) {
            return;
        }
        shell.selectDispatchSnapshot(ref);
    }

    private void rebuildOperatorItems() {
        operatorDirByLabel.clear();
        List<String> labels = new ArrayList<>();
        labels.add(localLabel());
        for (DispatchSnapshotStore.SnapshotRef ref : catalog) {
            String label = operatorLabel(ref.operatorDir());
            operatorDirByLabel.putIfAbsent(label, ref.operatorDir());
            if (!labels.contains(label)) {
                labels.add(label);
            }
        }
        operatorCombo.getItems().setAll(labels);
    }

    private void fillGenerations(String operatorDir) {
        refByLabel.clear();
        List<String> labels = new ArrayList<>();
        for (DispatchSnapshotStore.SnapshotRef ref : catalog) {
            if (!ref.operatorDir().equals(operatorDir)) {
                continue;
            }
            String label = generationLabel(ref);
            refByLabel.put(label, ref);
            labels.add(label);
        }
        generationCombo.getItems().setAll(labels);
    }

    private DispatchSnapshotStore.SnapshotRef newestOf(String operatorDir) {
        DispatchSnapshotStore.SnapshotRef best = null;
        for (DispatchSnapshotStore.SnapshotRef ref : catalog) {
            if (!ref.operatorDir().equals(operatorDir)) {
                continue;
            }
            if (best == null || ref.generationDir().compareTo(best.generationDir()) > 0) {
                best = ref;
            }
        }
        return best;
    }

    private static String localLabel() {
        return "自分のローカル最新";
    }

    private static String operatorLabel(String operatorDir) {
        return operatorDir == null || operatorDir.isBlank() ? "（無名）" : operatorDir;
    }

    private static String generationLabel(DispatchResultSelection selection) {
        String when = selection.detailText();
        return when == null || when.isBlank() ? selection.generationDir() : when;
    }

    private static String generationLabel(DispatchSnapshotStore.SnapshotRef ref) {
        StringBuilder sb = new StringBuilder();
        if (ref.savedAt() != null && !ref.savedAt().isBlank()) {
            sb.append(ref.savedAt());
        } else {
            sb.append(ref.generationDir());
        }
        if (ref.stage() != null && !ref.stage().isBlank()) {
            sb.append(' ').append(ref.stage());
        }
        if (ref.host() != null && !ref.host().isBlank()) {
            sb.append(' ').append(ref.host());
        }
        if (ref.missing() != null && !ref.missing().isEmpty()) {
            sb.append(" 欠落");
        }
        return sb.toString();
    }
}
