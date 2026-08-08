package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class MainShellRunTabGatingTest {

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void keepsGroupContainingRunEnabledAndDisablesBlockedSibling() throws Exception {
        CountDownLatch completed = new CountDownLatch(1);
        AtomicReference<AssertionError> failure = new AtomicReference<>();
        Platform.runLater(
                () -> {
                    try {
                        Tab run = new Tab("run");
                        Tab blocked = new Tab("blocked");
                        TabPane inner = new TabPane(run, blocked);
                        Tab group = new Tab("group", inner);
                        Tab remote = new Tab("remote");
                        TabPane outer = new TabPane(group, remote);

                        MainShellRunTabGating.apply(
                                outer,
                                true,
                                tab -> tab == run || tab == remote,
                                run);

                        assertFalse(group.isDisable());
                        assertFalse(run.isDisable());
                        assertTrue(blocked.isDisable());
                        assertFalse(remote.isDisable());
                    } catch (AssertionError error) {
                        failure.set(error);
                    } finally {
                        completed.countDown();
                    }
                });

        assertTrue(completed.await(5, TimeUnit.SECONDS));
        if (failure.get() != null) {
            throw failure.get();
        }
    }

    @Test
    void afterDisableStillSelectsPreferredLeafWhenRemoteWouldStayOperable() throws Exception {
        CountDownLatch completed = new CountDownLatch(1);
        AtomicReference<AssertionError> failure = new AtomicReference<>();
        Platform.runLater(
                () -> {
                    try {
                        Tab run = new Tab("run");
                        Tab remote = new Tab("remote");
                        Tab blocked = new Tab("blocked");
                        TabPane outer = new TabPane(blocked, remote, run);
                        outer.getSelectionModel().select(blocked);

                        MainShellRunTabGating.apply(
                                outer,
                                true,
                                tab -> tab == run || tab == remote,
                                run);

                        assertEquals(run, outer.getSelectionModel().getSelectedItem());
                    } catch (AssertionError error) {
                        failure.set(error);
                    } finally {
                        completed.countDown();
                    }
                });

        assertTrue(completed.await(5, TimeUnit.SECONDS));
        if (failure.get() != null) {
            throw failure.get();
        }
    }

    @Test
    void envInitPendingKeepsEnvLeafOperableWhenBlockedSiblingWasSelected() throws Exception {
        CountDownLatch completed = new CountDownLatch(1);
        AtomicReference<AssertionError> failure = new AtomicReference<>();
        Platform.runLater(
                () -> {
                    try {
                        Tab env = new Tab("env");
                        env.setContent(new javafx.scene.layout.Pane());
                        Tab blocked = new Tab("blocked");
                        blocked.setContent(new javafx.scene.layout.Pane());
                        TabPane inner = new TabPane(blocked, env);
                        inner.getSelectionModel().select(blocked);
                        Tab group = new Tab("group", inner);
                        Tab other = new Tab("other");
                        TabPane outer = new TabPane(group, other);

                        MainShellRunTabGating.applyEnvInitPending(outer, env);

                        assertFalse(group.isDisable());
                        assertFalse(env.isDisable());
                        assertFalse(env.getContent().isDisable());
                        assertFalse(blocked.isDisable());
                        assertFalse(other.isDisable());
                        assertEquals(env, inner.getSelectionModel().getSelectedItem());
                        assertEquals(group, outer.getSelectionModel().getSelectedItem());
                    } catch (AssertionError error) {
                        failure.set(error);
                    } finally {
                        completed.countDown();
                    }
                });

        assertTrue(completed.await(5, TimeUnit.SECONDS));
        if (failure.get() != null) {
            throw failure.get();
        }
    }

    @Test
    void guestSessionOnlyKeepsRunLeafOperableWhenBlockedSiblingWasSelected() throws Exception {
        CountDownLatch completed = new CountDownLatch(1);
        AtomicReference<AssertionError> failure = new AtomicReference<>();
        Platform.runLater(
                () -> {
                    try {
                        Tab run = new Tab("run");
                        run.setContent(new javafx.scene.layout.Pane());
                        Tab blocked = new Tab("blocked");
                        blocked.setContent(new javafx.scene.layout.Pane());
                        TabPane inner = new TabPane(blocked, run);
                        inner.getSelectionModel().select(blocked);
                        Tab group = new Tab("group", inner);
                        Tab other = new Tab("other");
                        TabPane outer = new TabPane(group, other);

                        MainShellRunTabGating.applyGuestSessionOnly(outer, run);

                        assertFalse(group.isDisable());
                        assertFalse(run.isDisable());
                        assertFalse(run.getContent().isDisable());
                        assertFalse(blocked.isDisable());
                        assertFalse(other.isDisable());
                        assertEquals(run, inner.getSelectionModel().getSelectedItem());
                        assertEquals(group, outer.getSelectionModel().getSelectedItem());
                    } catch (AssertionError error) {
                        failure.set(error);
                    } finally {
                        completed.countDown();
                    }
                });

        assertTrue(completed.await(5, TimeUnit.SECONDS));
        if (failure.get() != null) {
            throw failure.get();
        }
    }
}
