package jp.co.pm.ai.planning.stage2;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.AppPaths;

class Stage2EnvParsingTest {

    @Test
    void javaDelegatesPythonDispatch_defaultOff() {
        assertFalse(Stage2EnvParsing.javaDelegatesPythonDispatch(Map.of()));
        assertFalse(Stage2EnvParsing.javaDelegatesPythonDispatch(null));
    }

    @Test
    void javaDelegatesPythonDispatch_truthy() {
        Map<String, String> m = new HashMap<>();
        m.put(AppPaths.KEY_PM_AI_STAGE2_JAVA_DELEGATE_PYTHON_DISPATCH, "1");
        assertTrue(Stage2EnvParsing.javaDelegatesPythonDispatch(m));
    }

    @Test
    void dispatchCoreExplicitJava_recognized() {
        Map<String, String> m = new HashMap<>();
        m.put(AppPaths.KEY_PM_AI_DISPATCH_ENGINE, "JAVA");
        assertTrue(Stage2EnvParsing.dispatchCoreExplicitJava(m));
        assertFalse(Stage2EnvParsing.dispatchCoreExplicitPython(m));
    }

    @Test
    void dispatchCoreExplicitPython_recognized() {
        Map<String, String> m = new HashMap<>();
        m.put(AppPaths.KEY_PM_AI_DISPATCH_ENGINE, "Python");
        assertTrue(Stage2EnvParsing.dispatchCoreExplicitPython(m));
        assertFalse(Stage2EnvParsing.dispatchCoreExplicitJava(m));
    }

    @Test
    void dispatchCoreUnset_bothFalse() {
        assertFalse(Stage2EnvParsing.dispatchCoreExplicitJava(Map.of()));
        assertFalse(Stage2EnvParsing.dispatchCoreExplicitPython(Map.of()));
    }
}
