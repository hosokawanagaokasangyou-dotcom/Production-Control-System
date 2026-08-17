package jp.co.pm.ai.desktop.ui;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicReference;

import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.ListCell;
import javafx.scene.control.ListView;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;

/**
 * アラジン入力用配台計画の世代ファイル（操作者別フォルダ配下の xlsx）を選ぶダイアログ。
 *
 * <p>「世代を開く」は Excel を開く。「同一化チェック」は比較対象 xlsx を返す（任意パスの参照可）。
 * index ファイルは持たず、共有の配台計画フォルダ配下のディレクトリ走査のみで一覧化する。
 */
public final class DispatchAladdinEntryGenerationDialog {

    private static final DateTimeFormatter TS =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm").withZone(ZoneId.systemDefault());

    private static final String ALL_OPERATORS = "（すべて）";

    private DispatchAladdinEntryGenerationDialog() {}

    private record GenerationFile(Path path, String operator, long lastModifiedMillis) {

        String displayLabel() {
            return TS.format(Instant.ofEpochMilli(lastModifiedMillis))
                    + "  ["
                    + operator
                    + "]  "
                    + path.getFileName();
        }
    }

    /** 世代一覧のルートフォルダ（共有側）。 */
    static Path generationRoot(Map<String, String> ui) {
        return AppPaths.aladdinEntryDispatchPlanDir(ui);
    }

    /**
     * ダイアログを表示する。
     *
     * @param defaultOperator 既定選択の操作者名（自分）。一覧に無いときは「すべて」
     */
    public static void show(Window owner, Map<String, String> ui, String defaultOperator) {
        showInternal(owner, ui, defaultOperator, false);
    }

    /**
     * 同一化チェック用に世代 xlsx（または参照した任意 xlsx）を選ぶ。
     * キャンセル時は empty。
     */
    public static Optional<Path> chooseGenerationFile(
            Window owner, Map<String, String> ui, String defaultOperator) {
        return showInternal(owner, ui, defaultOperator, true);
    }

    private static Optional<Path> showInternal(
            Window owner, Map<String, String> ui, String defaultOperator, boolean chooseForCheck) {
        Path root = generationRoot(ui);
        AtomicReference<Path> chosen = new AtomicReference<>();

        Stage stage = new Stage();
        stage.initModality(Modality.WINDOW_MODAL);
        if (owner != null) {
            stage.initOwner(owner);
        }
        stage.setTitle(
                chooseForCheck
                        ? "同一化チェック — 比較するExcelを選択"
                        : "アラジン入力用配台計画 — 世代を開く");

        Label hintLabel =
                new Label(
                        "比較する配台計画 Excel を選び、「このファイルでチェック」を押してください。"
                                + " 一覧に無いファイルは「参照…」から指定できます。");
        hintLabel.setWrapText(true);
        hintLabel.setVisible(chooseForCheck);
        hintLabel.setManaged(chooseForCheck);

        Label rootLabel = new Label("世代フォルダ: " + root);
        rootLabel.setWrapText(true);

        ComboBox<String> operatorCombo = new ComboBox<>();
        ListView<GenerationFile> listView = new ListView<>();
        listView.setCellFactory(
                lv ->
                        new ListCell<>() {
                            @Override
                            protected void updateItem(GenerationFile item, boolean empty) {
                                super.updateItem(item, empty);
                                setText(empty || item == null ? null : item.displayLabel());
                            }
                        });

        Runnable applyList =
                () -> {
                    String selected = operatorCombo.getValue();
                    String operatorFilter =
                            selected == null || ALL_OPERATORS.equals(selected) ? null : selected;
                    listView.setItems(
                            FXCollections.observableArrayList(
                                    listCandidates(root, operatorFilter, ui, chooseForCheck)));
                    if (!listView.getItems().isEmpty()
                            && listView.getSelectionModel().getSelectedItem() == null) {
                        listView.getSelectionModel().selectFirst();
                    }
                };

        Runnable refresh =
                () -> {
                    String selected = operatorCombo.getValue();
                    List<String> operators = listOperators(root);
                    List<String> comboItems = new ArrayList<>();
                    comboItems.add(ALL_OPERATORS);
                    comboItems.addAll(operators);
                    operatorCombo.setItems(FXCollections.observableArrayList(comboItems));
                    String next =
                            selected != null && comboItems.contains(selected)
                                    ? selected
                                    : (defaultOperator != null && comboItems.contains(defaultOperator)
                                            ? defaultOperator
                                            : ALL_OPERATORS);
                    operatorCombo.setValue(next);
                    applyList.run();
                };

        operatorCombo.valueProperty()
                .addListener(
                        (obs, a, b) -> {
                            if (b == null) {
                                return;
                            }
                            applyList.run();
                        });

        Button primaryButton = new Button(chooseForCheck ? "このファイルでチェック" : "開く");
        primaryButton.setDefaultButton(true);
        primaryButton.setOnAction(
                e -> {
                    if (chooseForCheck) {
                        acceptSelected(listView, stage, chosen);
                    } else {
                        openSelected(listView, stage);
                    }
                });
        listView.setOnMouseClicked(
                e -> {
                    if (e.getClickCount() >= 2) {
                        if (chooseForCheck) {
                            acceptSelected(listView, stage, chosen);
                        } else {
                            openSelected(listView, stage);
                        }
                    }
                });

        Button browseButton = new Button("参照…");
        browseButton.setVisible(chooseForCheck);
        browseButton.setManaged(chooseForCheck);
        browseButton.setOnAction(e -> browseAndChoose(stage, ui, chosen));

        Button openFolderButton = new Button("フォルダを開く");
        openFolderButton.setOnAction(
                e -> {
                    try {
                        Files.createDirectories(root);
                        java.awt.Desktop.getDesktop().open(root.toFile());
                    } catch (Exception ex) {
                        showError(stage, "フォルダを開けませんでした。\n" + ex.getMessage());
                    }
                });

        Button refreshButton = new Button("再読み");
        refreshButton.setOnAction(e -> refresh.run());

        Button closeButton = new Button(chooseForCheck ? "キャンセル" : "閉じる");
        closeButton.setCancelButton(true);
        closeButton.setOnAction(e -> stage.close());

        HBox topRow = new HBox(8, new Label("操作者"), operatorCombo, refreshButton);
        topRow.setAlignment(Pos.CENTER_LEFT);
        HBox buttonRow =
                chooseForCheck
                        ? new HBox(8, primaryButton, browseButton, openFolderButton, closeButton)
                        : new HBox(8, primaryButton, openFolderButton, closeButton);
        buttonRow.setAlignment(Pos.CENTER_RIGHT);

        VBox rootBox = new VBox(8, hintLabel, rootLabel, topRow, listView, buttonRow);
        rootBox.setPadding(new Insets(12));
        VBox.setVgrow(listView, Priority.ALWAYS);

        stage.setScene(new Scene(rootBox, 640, 480));
        refresh.run();
        if (chooseForCheck) {
            stage.showAndWait();
            return Optional.ofNullable(chosen.get());
        }
        stage.show();
        return Optional.empty();
    }

    private static void acceptSelected(
            ListView<GenerationFile> listView, Stage stage, AtomicReference<Path> chosen) {
        GenerationFile selected = listView.getSelectionModel().getSelectedItem();
        if (selected == null) {
            showError(stage, "比較する配台計画 Excel を選択してください。");
            return;
        }
        chosen.set(selected.path().toAbsolutePath().normalize());
        stage.close();
    }

    private static void browseAndChoose(
            Stage stage, Map<String, String> ui, AtomicReference<Path> chosen) {
        FileChooser chooser = new FileChooser();
        chooser.setTitle("比較する配台計画 Excel");
        chooser.getExtensionFilters()
                .add(new FileChooser.ExtensionFilter("Excel", "*.xlsx"));
        Path localDir = AppPaths.aladdinEntryDispatchPlanLocalDir(ui);
        Path genRoot = generationRoot(ui);
        Path initial =
                Files.isDirectory(localDir)
                        ? localDir
                        : (Files.isDirectory(genRoot) ? genRoot : null);
        if (initial != null) {
            chooser.setInitialDirectory(initial.toFile());
        }
        File file = chooser.showOpenDialog(stage);
        if (file == null) {
            return;
        }
        chosen.set(file.toPath().toAbsolutePath().normalize());
        stage.close();
    }

    private static void openSelected(ListView<GenerationFile> listView, Stage stage) {
        GenerationFile selected = listView.getSelectionModel().getSelectedItem();
        if (selected == null) {
            showError(stage, "開く世代ファイルを選択してください。");
            return;
        }
        try {
            DesktopFileOpener.openFileReadOnly(selected.path());
        } catch (IOException ex) {
            showError(stage, "ファイルを開けませんでした。\n" + selected.path() + "\n" + ex.getMessage());
        }
    }

    /** ルート直下の操作者フォルダ名（昇順）。 */
    private static List<String> listOperators(Path root) {
        if (!Files.isDirectory(root)) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        try (var stream = Files.list(root)) {
            stream.filter(Files::isDirectory)
                    .map(p -> p.getFileName().toString())
                    .sorted(String.CASE_INSENSITIVE_ORDER)
                    .forEach(out::add);
        } catch (IOException ignored) {
            return List.of();
        }
        return out;
    }

    /** 同一化チェック候補（ローカル最新＋世代）。テスト用。 */
    static List<Path> listCandidatePaths(
            Map<String, String> ui, String operator, boolean includeLocalLatest) {
        return listCandidates(generationRoot(ui), operator, ui, includeLocalLatest).stream()
                .map(GenerationFile::path)
                .toList();
    }

    /** 同一化チェック時はローカル最新を先頭に足す。 */
    private static List<GenerationFile> listCandidates(
            Path root, String operator, Map<String, String> ui, boolean includeLocalLatest) {
        List<GenerationFile> out = new ArrayList<>();
        if (includeLocalLatest) {
            Path local = AppPaths.aladdinEntryDispatchPlanLocalXlsxPath(ui);
            if (Files.isRegularFile(local)) {
                out.add(new GenerationFile(local, "ローカル最新", lastModifiedMillis(local)));
            }
        }
        out.addAll(listGenerations(root, operator));
        return out;
    }

    /** 世代 xlsx を新しい順に列挙する（{@code operator} が null のとき全操作者）。 */
    private static List<GenerationFile> listGenerations(Path root, String operator) {
        List<String> operators = operator != null ? List.of(operator) : listOperators(root);
        List<GenerationFile> out = new ArrayList<>();
        for (String op : operators) {
            Path dir = root.resolve(op);
            if (!Files.isDirectory(dir)) {
                continue;
            }
            try (var stream = Files.list(dir)) {
                stream.filter(Files::isRegularFile)
                        .filter(p -> p.getFileName().toString().endsWith(".xlsx"))
                        .forEach(
                                p ->
                                        out.add(
                                                new GenerationFile(
                                                        p, op, lastModifiedMillis(p))));
            } catch (IOException ignored) {
                // 個別フォルダの列挙失敗はスキップ
            }
        }
        out.sort(Comparator.comparingLong(GenerationFile::lastModifiedMillis).reversed());
        return out;
    }

    private static long lastModifiedMillis(Path p) {
        try {
            return Files.getLastModifiedTime(p).toMillis();
        } catch (IOException e) {
            return 0L;
        }
    }

    private static void showError(Stage owner, String message) {
        Alert alert = new Alert(Alert.AlertType.WARNING, message);
        alert.initOwner(owner);
        alert.setHeaderText(null);
        alert.showAndWait();
    }
}
