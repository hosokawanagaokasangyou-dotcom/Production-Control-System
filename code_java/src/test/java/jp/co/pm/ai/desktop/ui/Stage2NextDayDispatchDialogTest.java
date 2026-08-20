package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.planning.stage2.core.Stage2PlanRowDispatchQtyMetrics;

class Stage2NextDayDispatchDialogTest {

    private static final Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM UNIT_3045 =
            new Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM(3045, 3045, 3045, true);

    @Test
    void aladdinRowUsesNextDayRollCountAndConvertsItToLegacyExcludedMeters() {
        var row =
                new Stage2AladdinTodayExcludeNextDayDispatchDialog.Row(
                        "T-AL", "スリット", "スリット機1　湖南", 6090, 10660, UNIT_3045);

        assertEquals("1", row.rollCountProperty().get());
        assertEquals(6090.0, row.toEntryFromNextDayInput().excludeNextDayM(), 1e-9);

        row.rollCountProperty().set("0");
        assertEquals(9135.0, row.toEntryFromNextDayInput().excludeNextDayM(), 1e-9);
    }

    @Test
    void unifiedResultContainsBothKindsOfLegacyEntries() {
        var inProgress =
                new Stage2InProgressNextDayDispatchDialog.Row(
                        "T-IN",
                        "スリット",
                        "スリット機1　湖南",
                        2870,
                        13530,
                        10660,
                        0,
                        10660,
                        UNIT_3045);
        var aladdin =
                new Stage2AladdinTodayExcludeNextDayDispatchDialog.Row(
                        "T-AL", "スリット", "スリット機1　湖南", 6090, 10660, UNIT_3045);

        Stage2NextDayDispatchDialog.Result result =
                Stage2NextDayDispatchDialog.collectResult(
                        List.of(inProgress), List.of(aladdin));

        assertEquals(1, result.inProgressEntries().size());
        assertEquals(1, result.aladdinExcludeEntries().size());
        assertEquals(9135.0, result.inProgressEntries().get(0).nextDayDispatchM(), 1e-9);
        assertEquals(6090.0, result.aladdinExcludeEntries().get(0).excludeNextDayM(), 1e-9);
    }

    @Test
    void planInputOptionsUseTheSameNextDayDispatchMeaning() throws Exception {
        var resource =
                Stage2NextDayDispatchDialogTest.class.getResourceAsStream(
                        "/jp/co/pm/ai/desktop/fxml/PlanInputTab.fxml");
        assertTrue(resource != null);
        String fxml;
        try (resource) {
            fxml = new String(resource.readAllBytes(), StandardCharsets.UTF_8);
        }

        assertTrue(fxml.contains("アラジン当日対象行の翌日配台"));
        assertTrue(fxml.contains("①と②をまとめて設定"));
        assertFalse(fxml.contains("翌日除外量を設定"));
    }

    @Test
    void processColumnShowsTheRowProcessAndIsReadOnly() {
        var row =
                new Stage2InProgressNextDayDispatchDialog.Row(
                        "T-IN",
                        "スリット",
                        "スリット機1　湖南",
                        2870,
                        13530,
                        10660,
                        0,
                        10660,
                        UNIT_3045);

        var column = Stage2NextDayRollDispatchDialogSupport.createProcessColumn();
        var cellValue =
                column.getCellValueFactory()
                        .call(new javafx.scene.control.TableColumn.CellDataFeatures<>(
                                null, column, row));

        assertEquals("工程名", column.getText());
        assertEquals("スリット", cellValue.getValue());
        assertFalse(column.isEditable());
    }

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void rollCountChoicesAreZeroThroughMaxInclusive() {
        assertEquals(List.of("0"), Stage2NextDayRollDispatchDialogSupport.rollCountChoices(0));
        assertEquals(
                List.of("0", "1", "2", "3"),
                Stage2NextDayRollDispatchDialogSupport.rollCountChoices(3));
        assertEquals(List.of("0"), Stage2NextDayRollDispatchDialogSupport.rollCountChoices(-1));
    }

    @Test
    void clampRollCountChoiceStaysWithinMax() {
        assertEquals("0", Stage2NextDayRollDispatchDialogSupport.clampRollCountChoice("", 3));
        assertEquals("2", Stage2NextDayRollDispatchDialogSupport.clampRollCountChoice("2", 3));
        assertEquals("3", Stage2NextDayRollDispatchDialogSupport.clampRollCountChoice("9", 3));
        assertEquals("0", Stage2NextDayRollDispatchDialogSupport.clampRollCountChoice("x", 3));
    }

    @Test
    void rollCountColumnUsesNonEditableComboBox() throws Exception {
        CountDownLatch done = new CountDownLatch(1);
        AtomicReference<ComboBox<?>> comboRef = new AtomicReference<>();
        AtomicReference<String> columnText = new AtomicReference<>();
        AtomicReference<Throwable> error = new AtomicReference<>();
        Platform.runLater(
                () -> {
                    try {
                        TableColumn<Stage2NextDayRollDispatchDialogSupport.RowModel, String>
                                column =
                                        Stage2NextDayRollDispatchDialogSupport
                                                .createRollCountColumn("翌日配台(ロール)");
                        columnText.set(column.getText());
                        TableCell<Stage2NextDayRollDispatchDialogSupport.RowModel, String> cell =
                                column.getCellFactory().call(column);
                        assertInstanceOf(ComboBox.class, cell.getGraphic());
                        comboRef.set((ComboBox<?>) cell.getGraphic());
                    } catch (Throwable t) {
                        error.set(t);
                    } finally {
                        done.countDown();
                    }
                });
        assertTrue(done.await(5, TimeUnit.SECONDS));
        if (error.get() != null) {
            throw new AssertionError(error.get());
        }
        assertEquals("翌日配台(ロール)", columnText.get());
        ComboBox<?> combo = comboRef.get();
        assertFalse(combo.isEditable());
    }
}
