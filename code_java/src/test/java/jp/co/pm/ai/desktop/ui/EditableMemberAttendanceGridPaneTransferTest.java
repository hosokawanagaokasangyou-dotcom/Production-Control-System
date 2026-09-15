package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;
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
class EditableMemberAttendanceGridPaneTransferTest {

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
    void exportKeepsTransferredMemberWhenHiddenFromMonth() throws Exception {
        EditableMemberAttendanceGridPane pane = runOnFx(() -> {
            EditableMemberAttendanceGridPane p = new EditableMemberAttendanceGridPane();
            p.loadFromMemberGridJson(octoberJson());
            return p;
        });
        assertEquals(1, pane.memberRowCountForTest());
        assertEquals(2, pane.rosterCountForTest());
        Map<String, Object> patch = pane.exportPatchJson();
        @SuppressWarnings("unchecked")
        List<Map<String, Object>> roster = (List<Map<String, Object>>) patch.get("member_roster");
        assertEquals(2, roster.size());
        Map<String, Object> transferred =
                roster.stream()
                        .filter(e -> "異動".equals(e.get("name")))
                        .findFirst()
                        .orElseThrow();
        assertEquals("2026-09-15", transferred.get("inactive_from"));
    }

    @Test
    void transferDateLocksCellsOnAndAfterInactiveFrom() throws Exception {
        EditableMemberAttendanceGridPane pane = runOnFx(() -> {
            EditableMemberAttendanceGridPane p = new EditableMemberAttendanceGridPane();
            p.loadFromMemberGridJson(septemberJson());
            p.setMemberInactiveFrom("異動", LocalDate.of(2026, 9, 15));
            return p;
        });
        assertEquals(2, pane.memberRowCountForTest());
        assertFalse(pane.cellInactiveForTest(LocalDate.of(2026, 9, 14), "異動"));
        assertTrue(pane.cellInactiveForTest(LocalDate.of(2026, 9, 15), "異動"));
        assertFalse(pane.cellInactiveForTest(LocalDate.of(2026, 9, 15), "在籍"));
    }

    private static EditableMemberAttendanceGridPane runOnFx(
            java.util.function.Supplier<EditableMemberAttendanceGridPane> factory)
            throws Exception {
        AtomicReference<EditableMemberAttendanceGridPane> paneRef = new AtomicReference<>();
        AtomicReference<Throwable> fxError = new AtomicReference<>();
        CountDownLatch done = new CountDownLatch(1);
        Platform.runLater(
                () -> {
                    try {
                        EditableMemberAttendanceGridPane pane = factory.get();
                        Scene scene = new Scene(pane, 1200, 480);
                        pane.applyCss();
                        pane.autosize();
                        pane.layout();
                        paneRef.set(pane);
                    } catch (Throwable t) {
                        fxError.set(t);
                    } finally {
                        done.countDown();
                    }
                });
        assertTrue(done.await(15, TimeUnit.SECONDS), "FX が完了しない");
        if (fxError.get() != null) {
            throw new AssertionError(fxError.get());
        }
        assertNotNull(paneRef.get());
        return paneRef.get();
    }

    private static ObjectNode octoberJson() {
        return monthJson(2026, 10, true);
    }

    private static ObjectNode septemberJson() {
        return monthJson(2026, 9, false);
    }

    private static ObjectNode monthJson(int year, int month, boolean hideTransferredInMembers) {
        ObjectNode root = MAPPER.createObjectNode();
        root.put("year", year);
        root.put("month", month);
        ArrayNode dates = root.putArray("dates");
        LocalDate start = LocalDate.of(year, month, 1);
        int len = start.lengthOfMonth();
        for (int day = 1; day <= len; day++) {
            dates.add(start.withDayOfMonth(day).toString());
        }
        ArrayNode members = root.putArray("members");
        members.add("在籍");
        if (!hideTransferredInMembers) {
            members.add("異動");
        }
        ArrayNode roster = root.putArray("member_roster");
        ObjectNode stay = roster.addObject();
        stay.put("name", "在籍");
        stay.put("primary_role", "後加工");
        ObjectNode leave = roster.addObject();
        leave.put("name", "異動");
        leave.put("primary_role", "物流");
        leave.put("inactive_from", "2026-09-15");
        root.set("cells", MAPPER.createObjectNode());
        return root;
    }
}
