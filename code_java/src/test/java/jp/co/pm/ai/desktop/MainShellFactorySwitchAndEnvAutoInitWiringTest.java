package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;

class MainShellFactorySwitchAndEnvAutoInitWiringTest {

    @Test
    void switchActiveFactorySite_promptsBeforeBusyWhenNotStartup() throws Exception {
        String body = methodBody("private void switchActiveFactorySite(FactorySite newSite, boolean startup)");
        assertTrue(
                body.contains("FactorySiteSwitchConfirmation.shouldPromptUser(startup)"),
                "ユーザー操作の工場切替は確認ダイアログを出す");
        int prompt = body.indexOf("shouldPromptUser");
        int busy = body.indexOf("beginFactorySiteSwitchBusy");
        assertTrue(prompt >= 0 && busy > prompt, "確認後に進捗モーダルを出す");
        assertTrue(body.contains("revertFactorySiteCombosToPersistedSite"), "キャンセル時はコンボを戻す");
    }

    @Test
    void completeEnvVarsStartupCheck_autoInitializesWithoutButton() throws Exception {
        String body = methodBody("private void completeEnvVarsStartupCheck(boolean schedulePostStartupWork)");
        assertTrue(
                body.contains("autoInitializeEnvVarsIfNeeded"),
                "環境変数初期化が必要なときはボタン待ちせず自動実行する");
        int evaluate = body.indexOf("evaluateEnvVarsDifferFromInitialAtStartup");
        int autoInit = body.indexOf("autoInitializeEnvVarsIfNeeded");
        assertTrue(evaluate >= 0 && autoInit > evaluate, "照合のあとに自動初期化する");
    }

    private static String methodBody(String signature) throws Exception {
        Path java = Path.of("src/main/java/jp/co/pm/ai/desktop/MainShellController.java");
        String text = Files.readString(java, StandardCharsets.UTF_8);
        int start = text.indexOf(signature);
        assertTrue(start >= 0, signature + " が見つからない");
        int end = text.indexOf("\n    private ", start + signature.length());
        if (end < 0) {
            end = Math.min(text.length(), start + 2500);
        }
        return text.substring(start, end);
    }
}
