package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.stream.Collectors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TableCell;
import javafx.scene.control.TableColumn;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

import jp.co.pm.ai.desktop.config.AppPaths;
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

    @Test
    void clipboardHtmlIncludesTitleHeaderHintTableAndCurrentRollCount() {
        var row =
                new Stage2InProgressNextDayDispatchDialog.Row(
                        "W9-1",
                        "検査",
                        "熱融着機 湖南",
                        3900,
                        4500,
                        4500,
                        0,
                        600,
                        new Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM(
                                300, 300, 300, true));
        row.rollCountProperty().set("2");

        String html =
                Stage2NextDayRollDispatchDialogSupport.toClipboardHtml(
                        Stage2NextDayDispatchDialog.THEME, List.of(row));

        assertTrue(html.contains("<h2>"), html);
        assertTrue(html.contains("段階2 — 翌日の配台量"), html);
        assertTrue(html.contains("対象行について、翌日に配台するロール数を指定してください。"), html);
        assertTrue(html.contains("ロール数はコンボボックスから選びます。"), html);
        assertTrue(html.contains("<table"), html);
        assertTrue(html.contains("依頼NO"), html);
        assertTrue(html.contains("機械名"), html);
        assertTrue(html.contains("工程名"), html);
        assertTrue(html.contains("対象理由"), html);
        assertTrue(html.contains("実加工"), html);
        assertTrue(html.contains("換算数量"), html);
        assertTrue(html.contains("配台数量"), html);
        assertTrue(html.contains("アラジン当日"), html);
        assertTrue(html.contains("残量"), html);
        assertTrue(html.contains("1ロール"), html);
        assertTrue(html.contains("翌日配台(ロール)"), html);
        assertTrue(html.contains("換算(m)"), html);
        assertTrue(html.contains("W9-1"), html);
        assertTrue(html.contains("熱融着機 湖南"), html);
        assertTrue(html.contains("検査"), html);
        assertTrue(html.contains("加工途中"), html);
        assertTrue(html.contains(">2<") || html.contains(">2</td>"), html);
        assertTrue(html.contains("600 m"), html);
    }

    @Test
    void clipboardHtmlUsesTheRollCountCurrentlySelectedInTheCombo() {
        var row =
                new Stage2InProgressNextDayDispatchDialog.Row(
                        "W9-2",
                        "EC",
                        "EC機 湖南",
                        2100,
                        4500,
                        4500,
                        0,
                        2400,
                        new Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM(
                                300, 300, 300, true));
        row.rollCountProperty().set("8");

        String html =
                Stage2NextDayRollDispatchDialogSupport.toClipboardHtml(
                        Stage2NextDayDispatchDialog.THEME, List.of(row));

        assertTrue(html.contains(">8</td>"), html);
        assertTrue(html.contains("2400 m"), html);
    }

    @Test
    void clipboardHtmlEscapesSpecialCharacters() {
        var row =
                new Stage2InProgressNextDayDispatchDialog.Row(
                        "A&B<C>",
                        "検査",
                        "機\"械",
                        0,
                        0,
                        0,
                        0,
                        0,
                        UNIT_3045);

        String html =
                Stage2NextDayRollDispatchDialogSupport.toClipboardHtml(
                        Stage2NextDayDispatchDialog.THEME, List.of(row));

        assertTrue(html.contains("A&amp;B&lt;C&gt;"), html);
        assertTrue(html.contains("機&quot;械"), html);
        assertFalse(html.contains("A&B<C>"), html);
    }

    @Test
    void clipboardHtmlOmitsOptionalColumnsWhenThemeHidesThem() {
        var row =
                new Stage2AladdinTodayExcludeNextDayDispatchDialog.Row(
                        "T9-1",
                        "エンボス",
                        "エンボス 湖南",
                        2800,
                        2000,
                        new Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM(
                                400, 400, 400, true));
        row.rollCountProperty().set("5");

        var theme =
                new Stage2NextDayRollDispatchDialogSupport.Theme(
                        "アラジン当日配台 — 翌日の配台量",
                        "ヘッダ",
                        "ヒント",
                        "実加工",
                        "翌日配台(ロール)",
                        "",
                        "",
                        true,
                        false);

        String html =
                Stage2NextDayRollDispatchDialogSupport.toClipboardHtml(theme, List.of(row));

        assertTrue(html.contains("アラジン当日"), html);
        assertTrue(html.contains("対象理由"), html);
        assertFalse(html.contains("換算数量"), html);
        assertFalse(html.contains("配台数量"), html);
        assertTrue(html.contains(">5</td>"), html);
    }

    @Test
    void copyHtmlButtonTypeDoesNotCloseTheDialogAsOk() {
        assertEquals("HTMLコピー", Stage2NextDayRollDispatchDialogSupport.COPY_HTML.getText());
        assertTrue(
                Stage2NextDayRollDispatchDialogSupport.COPY_HTML.getButtonData().isCancelButton()
                        == false);
        assertTrue(
                Stage2NextDayRollDispatchDialogSupport.COPY_HTML.getButtonData().isDefaultButton()
                        == false);
    }

    @Test
    void writeA4LandscapeDocx_usesA4LandscapePageAndCurrentRollCount() throws Exception {
        var row =
                new Stage2InProgressNextDayDispatchDialog.Row(
                        "W9-1",
                        "検査",
                        "熱融着機 湖南",
                        3900,
                        4500,
                        4500,
                        0,
                        600,
                        new Stage2PlanRowDispatchQtyMetrics.DispatchSimulatorUnitM(
                                300, 300, 300, true));
        row.rollCountProperty().set("2");

        Path dest = Files.createTempFile("stage2-next-day-", ".docx");
        try {
            Stage2NextDayRollDispatchDialogSupport.writeA4LandscapeDocx(
                    Stage2NextDayDispatchDialog.THEME, List.of(row), dest);

            try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(dest))) {
                CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
                assertNotNull(sectPr);
                CTPageSz pageSz = sectPr.getPgSz();
                assertNotNull(pageSz);
                assertEquals(STPageOrientation.LANDSCAPE, pageSz.getOrient());
                assertEquals(
                        Stage2NextDayRollDispatchDialogSupport.A4_LANDSCAPE_WIDTH_TWIPS,
                        Long.parseLong(String.valueOf(pageSz.getW())));
                assertEquals(
                        Stage2NextDayRollDispatchDialogSupport.A4_LANDSCAPE_HEIGHT_TWIPS,
                        Long.parseLong(String.valueOf(pageSz.getH())));

                String paras =
                        doc.getParagraphs().stream()
                                .map(XWPFParagraph::getText)
                                .collect(Collectors.joining("\n"));
                assertTrue(paras.contains("段階2 — 翌日の配台量"), paras);
                assertTrue(paras.contains("対象行について、翌日に配台するロール数を指定してください。"), paras);

                assertEquals(1, doc.getTables().size());
                XWPFTable table = doc.getTables().get(0);
                List<String> headers =
                        table.getRow(0).getTableCells().stream()
                                .map(XWPFTableCell::getText)
                                .toList();
                assertTrue(headers.contains("依頼NO"), headers.toString());
                assertTrue(headers.contains("翌日配台(ロール)"), headers.toString());
                assertTrue(headers.contains("換算(m)"), headers.toString());

                assertEquals(2, table.getNumberOfRows());
                List<String> cells =
                        table.getRow(1).getTableCells().stream()
                                .map(XWPFTableCell::getText)
                                .toList();
                assertEquals("W9-1", cells.get(headers.indexOf("依頼NO")));
                assertEquals("2", cells.get(headers.indexOf("翌日配台(ロール)")));
                assertEquals("600 m", cells.get(headers.indexOf("換算(m)")));
            }
        } finally {
            Files.deleteIfExists(dest);
        }
    }

    @Test
    void exportWordButtonTypeDoesNotCloseTheDialogAsOk() {
        assertEquals("Word出力", Stage2NextDayRollDispatchDialogSupport.EXPORT_WORD.getText());
        assertTrue(
                Stage2NextDayRollDispatchDialogSupport.EXPORT_WORD.getButtonData().isCancelButton()
                        == false);
        assertTrue(
                Stage2NextDayRollDispatchDialogSupport.EXPORT_WORD.getButtonData().isDefaultButton()
                        == false);
    }

    @Test
    void resolveExportWordPath_isUnderDesktopAppHomeTempWordSubfolder() {
        Path dest =
                Stage2NextDayRollDispatchDialogSupport.resolveExportWordPath(
                        Stage2NextDayDispatchDialog.THEME);
        Path home = AppPaths.resolveDesktopAppHomeDir();
        Path expectedDir = home.resolve(AppPaths.TEMP_WORD_DIR);
        assertEquals(expectedDir, dest.getParent());
        assertEquals(
                AppPaths.resolveSessionStateStorePath().getParent(),
                dest.getParent().getParent());
        assertEquals("段階2 — 翌日の配台量.docx", dest.getFileName().toString());
    }
}
