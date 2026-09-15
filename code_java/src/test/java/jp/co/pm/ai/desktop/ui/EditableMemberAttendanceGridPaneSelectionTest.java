package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Supplier;

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
class EditableMemberAttendanceGridPaneSelectionTest {

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
    void selectedMemberMarksNameRoleAndDateRow() throws Exception {
        EditableMemberAttendanceGridPane pane =
                runOnFx(
                        () -> {
                            EditableMemberAttendanceGridPane p =
                                    new EditableMemberAttendanceGridPane();
                            p.loadFromMemberGridJson(twoMemberJson());
                            p.selectMemberForTest("吉岡 廣海");
                            return p;
                        });
        assertEquals("吉岡 廣海", pane.selectedMemberName());
        assertTrue(pane.memberNameHasSelectedStyleForTest("吉岡 廣海"));
        assertTrue(pane.memberRoleHasSelectedStyleForTest("吉岡 廣海"));
        assertTrue(pane.memberRowBandHasSelectedStyleForTest("吉岡 廣海"));
        assertFalse(pane.memberNameHasSelectedStyleForTest("河合 直樹"));
        assertFalse(pane.memberRoleHasSelectedStyleForTest("河合 直樹"));
        assertFalse(pane.memberRowBandHasSelectedStyleForTest("河合 直樹"));
    }

    private static EditableMemberAttendanceGridPane runOnFx(
            Supplier<EditableMemberAttendanceGridPane> factory) throws Exception {
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

    private static ObjectNode twoMemberJson() {
        ObjectNode root = MAPPER.createObjectNode();
        root.put("year", 2026);
        root.put("month", 9);
        ArrayNode dates = root.putArray("dates");
        LocalDate start = LocalDate.of(2026, 9, 1);
        for (int day = 1; day <= 3; day++) {
            dates.add(start.withDayOfMonth(day).toString());
        }
        ArrayNode members = root.putArray("members");
        members.add("河合 直樹");
        members.add("吉岡 廣海");
        ArrayNode roster = root.putArray("member_roster");
        ObjectNode a = roster.addObject();
        a.put("name", "河合 直樹");
        a.put("primary_role", "後加工");
        ObjectNode b = roster.addObject();
        b.put("name", "吉岡 廣海");
        b.put("primary_role", "後加工");
        root.set("cells", MAPPER.createObjectNode());
        return root;
    }
}
