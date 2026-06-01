package jp.co.pm.ai.desktop;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.beans.property.ReadOnlyStringWrapper;

import jp.co.pm.ai.desktop.io.PlanInputTabularIo;

/**
 * 「配台計画_タスク入力3.0」タブ。段階3.0 前処理（入力3表生成）で書き出した枝番タスクを
 * 表示し、段階3.0/3.1/3.2 の実行起点となる。表は {@link PlanInputTabularIo} で読み出した
 * 読み取り専用ビュー（編集は段階2.0 入力タブと異なり当面行わない）。
 */
public class PlanInputStage3TabController {

    /** 入力3表シート名（Python 側 PLAN_INPUT_STAGE3_SHEET_NAME の既定と一致）。 */
    public static final String STAGE3_SHEET_NAME = "\u914d\u53f0\u8a08\u753b_\u30bf\u30b9\u30af\u5165\u529b3.0";

    @FXML private Button stage30RunButton;
    @FXML private Button stage31RunButton;
    @FXML private Button stage32RunButton;
    @FXML private Button reloadButton;
    @FXML private Label statusLabel;
    @FXML private TableView<List<String>> tableView;

    private MainShellController shell;

    public void bindShell(MainShellController shell) {
        this.shell = shell;
    }

    @FXML
    private void onStage30RunButtonAction() {
        if (shell != null) {
            shell.triggerStage30();
        }
    }

    @FXML
    private void onStage31RunButtonAction() {
        if (shell != null) {
            shell.triggerStage31();
        }
    }

    @FXML
    private void onStage32RunButtonAction() {
        if (shell != null) {
            shell.triggerStage32();
        }
    }

    @FXML
    private void onReloadButtonAction() {
        reloadFromDisk();
    }

    /** 入力3表シートをディスクから読み込み、表に反映する。 */
    public void reloadFromDisk() {
        Path workbook = resolveWorkbookPath();
        if (workbook == null || !Files.isRegularFile(workbook)) {
            setStatus("入力3表の元ブックが見つかりません。配台計画手動修正タブで「入力3表を生成」してください。");
            setRows(List.of(), List.of());
            return;
        }
        try {
            PlanInputTabularIo.TabularSheet sheet =
                    PlanInputTabularIo.read(workbook, STAGE3_SHEET_NAME);
            setRows(sheet.headers(), sheet.rows());
            setStatus("入力3表: " + sheet.rows().size() + " 行（" + workbook + "）");
        } catch (Exception ex) {
            setRows(List.of(), List.of());
            setStatus(
                    "入力3表シート「" + STAGE3_SHEET_NAME + "」を読み込めません。"
                            + "段階3.0 前処理（入力3表生成）が未実行の可能性があります。詳細: "
                            + ex.getMessage());
        }
    }

    private Path resolveWorkbookPath() {
        if (shell == null) {
            return null;
        }
        String p = shell.stage3PlanInputWorkbookPath();
        if (p == null || p.isBlank()) {
            return null;
        }
        try {
            return Path.of(p.trim());
        } catch (Exception ex) {
            return null;
        }
    }

    private void setRows(List<String> headers, List<List<String>> rows) {
        Runnable apply =
                () -> {
                    tableView.getColumns().clear();
                    for (int i = 0; i < headers.size(); i++) {
                        final int col = i;
                        TableColumn<List<String>, String> tc = new TableColumn<>(headers.get(i));
                        tc.setCellValueFactory(
                                cd -> {
                                    List<String> r = cd.getValue();
                                    String v = (r != null && col < r.size()) ? r.get(col) : "";
                                    return new ReadOnlyStringWrapper(v);
                                });
                        tableView.getColumns().add(tc);
                    }
                    ObservableList<List<String>> data = FXCollections.observableArrayList(rows);
                    tableView.setItems(data);
                };
        if (Platform.isFxApplicationThread()) {
            apply.run();
        } else {
            Platform.runLater(apply);
        }
    }

    private void setStatus(String msg) {
        Runnable apply = () -> statusLabel.setText(msg);
        if (Platform.isFxApplicationThread()) {
            apply.run();
        } else {
            Platform.runLater(apply);
        }
    }
}
