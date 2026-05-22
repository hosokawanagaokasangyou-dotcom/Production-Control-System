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
                        model("models/gemini-2.0-flash-lite", "generateContent"),
                        model("models/gemini-3.1-flash-lite-preview", "generateContent"),
                        model("models/gemini-3.1-flash-lite", "generateContent"),
                        model("models/gemini-2.5-flash-lite", "generateContent"),
                        model("models/gemini-2.5-flash", "generateContent"),
                        model("models/text-embedding-004", "embedContent"));
        List<String> out = GeminiFreeTierModelSelector.selectFlashLiteGenerateContentModels(listed);
        assertEquals(
                List.of(
                        "gemini-3.1-flash-lite",
                        "gemini-3.1-flash-lite-preview",
                        "gemini-2.5-flash-lite",
                        "gemini-2.0-flash-lite"),
                out);
    }

    @Test
    void emptyWhenNoFlashLite() {
        List<String> out =
                GeminiFreeTierModelSelector.selectFlashLiteGenerateContentModels(
                        List.of(model("models/gemini-2.5-flash", "generateContent")));
        assertTrue(out.isEmpty());
    }

    private static ListedModel model(String name, String... methods) {
        return new ListedModel(name, List.of(methods));
    }
}
