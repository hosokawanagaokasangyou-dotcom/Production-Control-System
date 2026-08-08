package jp.co.pm.ai.desktop.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

class SingleInstanceGuardTest {

    private SingleInstanceGuard guard;
    private int port;

    @BeforeEach
    void setUp() throws Exception {
        port = SingleInstanceGuard.findFreePort();
        System.setProperty(SingleInstanceGuard.PROP_ENABLED, "true");
        System.setProperty(SingleInstanceGuard.PROP_PORT, Integer.toString(port));
        guard = new SingleInstanceGuard();
    }

    @AfterEach
    void tearDown() {
        if (guard != null) {
            guard.close();
        }
        System.clearProperty(SingleInstanceGuard.PROP_ENABLED);
        System.clearProperty(SingleInstanceGuard.PROP_PORT);
    }

    @Test
    void primaryAcceptsActivateAndInvokesCallbackOnce() throws Exception {
        AtomicInteger activations = new AtomicInteger();
        CountDownLatch latch = new CountDownLatch(1);
        guard.setOnActivateRequest(
                () -> {
                    activations.incrementAndGet();
                    latch.countDown();
                });

        assertEquals(SingleInstanceGuard.Role.PRIMARY, guard.tryAcquire());

        assertTrue(SingleInstanceGuard.sendActivate(port, 500));
        assertTrue(latch.await(2, TimeUnit.SECONDS));
        assertEquals(1, activations.get());
    }

    @Test
    void secondAcquireBecomesSecondaryWhenPrimaryListening() throws Exception {
        assertEquals(SingleInstanceGuard.Role.PRIMARY, guard.tryAcquire());

        SingleInstanceGuard second = new SingleInstanceGuard();
        try {
            assertEquals(SingleInstanceGuard.Role.SECONDARY, second.tryAcquire());
        } finally {
            second.close();
        }
    }

    @Test
    void disabledPropertySkipsGuard() throws Exception {
        System.setProperty(SingleInstanceGuard.PROP_ENABLED, "false");
        assertEquals(SingleInstanceGuard.Role.DISABLED, guard.tryAcquire());
        assertFalse(SingleInstanceGuard.sendActivate(port, 200));
    }
}
