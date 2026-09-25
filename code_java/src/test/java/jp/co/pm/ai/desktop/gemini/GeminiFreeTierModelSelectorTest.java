package jp.co.pm.ai.desktop.gemini;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.benchmark.GeminiModelsListRestClient.ListedModel;

class GeminiFreeTierModelSelectorTest {

    @Test
    void selectsFlashLiteWithGenerateContent_sortedNewestFirst() {
        List<ListedModel> listed =
                List.of(
                        model("models/gemini-3.1-flash-lite-preview", "generateContent"),
                        model("models/gemini-3.1-flash-lite", "generateContent"),
                        model("models/gemini-3.5-flash-lite", "generateContent"),
                        model("models/gemini-3.8-flash", "generateContent"),
                        model("models/gemini-3.7-flash", "generateContent"),
                        model("models/gemini-3.8-flash-tts", "generateContent"),
                        model("models/gemini-3.1-flash-lite-image", "generateContent"),
                        model("models/gemini-2.5-flash", "generateContent"),
                        model("models/text-embedding-004", "embedContent"));
        List<String> out = GeminiFreeTierModelSelector.selectFlashLiteGenerateContentModels(listed);
        assertEquals(
                List.of(
                        "gemini-3.8-flash",
                        "gemini-3.7-flash",
                        "gemini-3.5-flash-lite",
                        "gemini-3.1-flash-lite",
                        "gemini-3.1-flash-lite-preview"),
                out);
    }

    @Test
    void dropsGenerationsWithoutFreeTierAllocation() {
        List<ListedModel> listed =
                List.of(
                        model("models/gemini-2.0-flash-lite", "generateContent"),
                        model("models/gemini-2.5-flash-lite", "generateContent"),
                        model("models/gemini-1.5-flash-lite", "generateContent"),
                        model("models/gemini-3.1-flash-lite", "generateContent"));
        List<String> out = GeminiFreeTierModelSelector.selectFlashLiteGenerateContentModels(listed);
        assertEquals(List.of("gemini-3.1-flash-lite"), out);
    }

    @Test
    void keepsStableFlashWithoutLiteSuffix() {
        List<String> out =
                GeminiFreeTierModelSelector.selectFlashLiteGenerateContentModels(
                        List.of(model("models/gemini-3.8-flash", "generateContent")));
        assertEquals(List.of("gemini-3.8-flash"), out);
    }

    @Test
    void emptyWhenNoTextFlash() {
        List<String> out =
                GeminiFreeTierModelSelector.selectFlashLiteGenerateContentModels(
                        List.of(
                                model("models/gemini-3.5-pro", "generateContent"),
                                model("models/gemini-3.8-flash-tts", "generateContent")));
        assertTrue(out.isEmpty());
    }

    private static ListedModel model(String name, String... methods) {
        return new ListedModel(name, List.of(methods));
    }
}
