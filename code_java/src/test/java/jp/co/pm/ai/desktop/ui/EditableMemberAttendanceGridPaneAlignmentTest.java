package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.time.YearMonth;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import javafx.application.Platform;
import javafx.scene.Scene;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledOnOs;
import org.junit.jupiter.api.condition.OS;

@EnabledOnOs(OS.WINDOWS)
class EditableMemberAttendanceGridPaneAlignmentTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void nameRowsStayAlignedWithDateCells_whenTodayColumnIsHighlighted() throws Exception {
        AtomicReference<EditableMemberAttendanceGridPane> paneRef = new AtomicReference<>();
        AtomicReference<Throwable> fxError = new AtomicReference<>();
        CountDownLatch done = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    try {
                        EditableMemberAttendanceGridPane pane =
                                new EditableMemberAttendanceGridPane();
                        pane.loadFromMemberGridJson(sampleMonthJson(14));
                        Scene scene = new Scene(pane, 1600, 720);
                        var css =
                                EditableMemberAttendanceGridPane.class.getResource(
                                        "/jp/co/pm/ai/desktop/css/pm-ai-desktop.css");
                        assertNotNull(css, "pm-ai-desktop.css");
                        scene.getStylesheets().add(css.toExternalForm());
                        pane.applyCss();
                        pane.autosize();
                        pane.layout();
                        pane.forceLayoutForTests();
                        paneRef.set(pane);
                    } catch (Throwable t) {
                        fxError.set(t);
                    } finally {
                        done.countDown();
                    }
                });
        assertTrue(done.await(15, TimeUnit.SECONDS), "FX レイアウトが完了しない");
        if (fxError.get() != null) {
            throw new AssertionError(fxError.get());
        }
        EditableMemberAttendanceGridPane pane = paneRef.get();
        assertNotNull(pane);
        assertEquals(14, pane.memberRowCountForTest());
        assertTrue(pane.leftBodyPrefHeightForTest() > 40, "左グリッドがレイアウトされている");
        assertEquals(
                pane.leftBodyPrefHeightForTest(),
                pane.bodyDatePrefHeightForTest(),
                0.5,
                "氏名列と日付セルの総高さが一致すること");
        assertTrue(
                pane.maxNameCellRowLayoutYDelta() < 1.0,
                "各行の氏名と日付セルの Y がずれている: " + pane.maxNameCellRowLayoutYDelta());
    }

    private static ObjectNode sampleMonthJson(int memberCount) {
        LocalDate today = LocalDate.now();
        YearMonth ym = YearMonth.from(today);
        ObjectNode root = MAPPER.createObjectNode();
        root.put("year", ym.getYear());
        root.put("month", ym.getMonthValue());
        ArrayNode dates = root.putArray("dates");
        for (int day = 1; day <= ym.lengthOfMonth(); day++) {
            dates.add(ym.atDay(day).toString());
        }
        ArrayNode roster = root.putArray("member_roster");
        List<String> names =
                List.of(
                        "細川 守",
                        "砂田 奈美",
                        "古家 淳子",
                        "宮島 剛",
                        "岡司 智子",
                        "冨田 裕子",
                        "春樹 真由美",
                        "竹内 正美",
                        "菅沼 あぐみ",
                        "森下 誠",
                        "小川 達也",
                        "西田 憲史",
                        "近藤 清高",
                        "東出 繁利");
        for (int i = 0; i < memberCount; i++) {
            ObjectNode ent = roster.addObject();
            ent.put("name", names.get(i % names.size()) + (i >= names.size() ? i : ""));
            ent.put("primary_role", i % 5 == 0 ? "物流" : "後加工");
        }
        root.set("cells", MAPPER.createObjectNode());
        return root;
    }
}
