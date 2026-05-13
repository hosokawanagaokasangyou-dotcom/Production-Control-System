package jp.co.pm.ai.planning.stage2.parity;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.databind.ObjectMapper;

class Stage2PlanJsonGoldenNormalizerTest {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    @Test
    void stripTopLevelVolatile_removesSourceXlsx() throws Exception {
        var root = MAPPER.readTree("{\"format_version\":2,\"source_xlsx\":\"計画999.json\",\"sheets\":{}}");
        var n = Stage2PlanJsonGoldenNormalizer.stripTopLevelVolatile(root);
        assertTrue(n.has("format_version"));
        assertTrue(n.has("sheets"));
        assertFalse(n.has("source_xlsx"));
    }
}
