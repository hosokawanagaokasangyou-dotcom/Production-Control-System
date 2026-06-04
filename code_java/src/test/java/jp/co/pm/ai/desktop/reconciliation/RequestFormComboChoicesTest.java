package jp.co.pm.ai.desktop.reconciliation;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RequestFormComboChoicesTest {

    private static final ObjectMapper JSON = new ObjectMapper();

    @Test
    void fromJson_roundTrip() throws Exception {
        RequestFormComboChoices original =
                RequestFormComboChoices.of(
                        Map.of(
                                RequestFormComboChoices.KEY_USER,
                                List.of("A社", "B社"),
                                RequestFormComboChoices.KEY_YOTO,
                                List.of("輸出", "国内")));
        ObjectNode root = JSON.createObjectNode();
        original.writeToObjectNode(root);

        RequestFormComboChoices loaded = RequestFormComboChoices.fromJson(root);
        assertEquals(List.of("A社", "B社"), loaded.optionsFor(RequestFormComboChoices.KEY_USER));
        assertEquals(List.of("輸出", "国内"), loaded.optionsFor(RequestFormComboChoices.KEY_YOTO));
        assertEquals(
                RequestFormComboChoices.bundledDefaults().optionsFor(RequestFormComboChoices.KEY_INPUT_KBN),
                loaded.optionsFor(RequestFormComboChoices.KEY_INPUT_KBN));
    }

    @Test
    void fieldDefaults_roundTripAndEffectiveDefault() throws Exception {
        RequestFormComboChoices original =
                RequestFormComboChoices.of(
                        Map.of(),
                        Map.of(
                                RequestFormComboChoices.KEY_INPUT_KBN,
                                "例外入力",
                                RequestFormComboChoices.KEY_KAKO_KBN,
                                "TPI"));
        ObjectNode root = JSON.createObjectNode();
        original.writeToObjectNode(root);

        RequestFormComboChoices loaded = RequestFormComboChoices.fromJson(root);
        assertEquals("例外入力", loaded.defaultFor(RequestFormComboChoices.KEY_INPUT_KBN));
        assertEquals("TPI", loaded.defaultFor(RequestFormComboChoices.KEY_KAKO_KBN));
        assertEquals(
                "例外入力",
                loaded.effectiveDefaultFor(RequestFormComboChoices.KEY_INPUT_KBN));
        assertEquals("TPI", loaded.effectiveDefaultFor(RequestFormComboChoices.KEY_KAKO_KBN));
    }

    @Test
    void bundledFieldDefaults_matchWorkInstructionSpec() {
        assertEquals(
                "通常入力",
                RequestFormComboChoices.bundledDefaults()
                        .effectiveDefaultFor(RequestFormComboChoices.KEY_INPUT_KBN));
        assertEquals(
                "後加工",
                RequestFormComboChoices.bundledDefaults()
                        .effectiveDefaultFor(RequestFormComboChoices.KEY_KAKO_KBN));
    }

    @Test
    void empty_fallsBackToBundledDefaults() {
        RequestFormComboChoices empty = RequestFormComboChoices.empty();
        assertTrue(empty.isEmpty());
        assertFalse(
                RequestFormComboChoices.bundledDefaults()
                        .optionsFor(RequestFormComboChoices.KEY_USER)
                        .isEmpty());
        assertEquals(
                RequestFormComboChoices.bundledDefaults().optionsFor(RequestFormComboChoices.KEY_USER),
                empty.optionsFor(RequestFormComboChoices.KEY_USER));
    }

    @Test
    void masterCandidatePrefixFilters_roundTrip() throws Exception {
        RequestFormComboChoices original =
                RequestFormComboChoices.of(
                        Map.of(
                                RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_PRODUCT,
                                List.of("A2", "B1"),
                                RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_RAW,
                                List.of("R1")));
        ObjectNode root = JSON.createObjectNode();
        original.writeToObjectNode(root);

        RequestFormComboChoices loaded = RequestFormComboChoices.fromJson(root);
        assertEquals(
                List.of("A2", "B1"),
                loaded.optionsFor(RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_PRODUCT));
        assertEquals(
                List.of("R1"),
                loaded.optionsFor(RequestFormComboChoices.KEY_MASTER_CANDIDATE_PREFIX_RAW));
    }
}
