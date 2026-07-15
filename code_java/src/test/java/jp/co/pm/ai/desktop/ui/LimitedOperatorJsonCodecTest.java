package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;

import org.junit.jupiter.api.Test;

class LimitedOperatorJsonCodecTest {

    @Test
    void decodeAndEncode_preserveSelectionOrder() {
        List<String> names = LimitedOperatorJsonCodec.decode("[\"山田\",\"佐藤\"]");

        assertEquals(List.of("山田", "佐藤"), names);
        assertEquals("[\"山田\",\"佐藤\"]", LimitedOperatorJsonCodec.encode(names));
    }

    @Test
    void decode_emptyCellIsEmptySelectionAndEncodeEmptyIsEmptyCell() {
        assertEquals(List.of(), LimitedOperatorJsonCodec.decode(""));
        assertEquals("", LimitedOperatorJsonCodec.encode(List.of()));
    }

    @Test
    void decode_rejectsMalformedJsonNonArrayNonStringAndDuplicates() {
        assertThrows(IllegalArgumentException.class, () -> LimitedOperatorJsonCodec.decode("["));
        assertThrows(IllegalArgumentException.class, () -> LimitedOperatorJsonCodec.decode("{\"name\":\"山田\"}"));
        assertThrows(IllegalArgumentException.class, () -> LimitedOperatorJsonCodec.decode("[\"山田\",1]"));
        assertThrows(
                IllegalArgumentException.class,
                () -> LimitedOperatorJsonCodec.decode("[\"山田\",\"山田\"]"));
    }

    @Test
    void decode_rejectsTrailingJsonTokens() {
        assertThrows(
                IllegalArgumentException.class,
                () -> LimitedOperatorJsonCodec.decode("[\"山田\"] true"));
    }
}
