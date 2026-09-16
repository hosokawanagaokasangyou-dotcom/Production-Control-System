package jp.co.pm.ai.kouchin.verify;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class TorayCsvReaderTest {

    private static final String HEADER =
            "入庫場所,生産場所,発注No.,入庫日,品番,品名カナ,数量,品名,単価,金額,備考1,備考2";

    @TempDir
    Path tempDir;

    /** A010 の3行（うち1行は取消行）と A010P の1行、末尾に符号込みのT小計。 */
    private Path sampleCsv() throws IOException {
        String csv = String.join("\n",
                HEADER,
                "*****,,,,,,,,,,,",
                "A010,N122B,191-352R,250705,X1,ｶﾅ,10,,1000,10000,,",
                ",,191-353R,250706,X2,ｶﾅ,5,,1000,5000,,",
                ",,190-765Y,250710,X3,ｶﾅ,8,-,2320,18560,,",
                "A010P,N122F,191-999X,250712,X4,ｶﾅ,4,,5000,20000,,",
                ",,T,,,,,,,16440,,",
                "");
        return write("RVSHEET202607.csv", csv);
    }

    private Path write(String name, String csv) throws IOException {
        Path path = tempDir.resolve(name);
        Files.write(path, csv.getBytes(Charset.forName("windows-31j")));
        return path;
    }

    @Test
    @DisplayName("入庫場所を前方補完し、対象工場の契約NOだけを集計する")
    void aggregatesTargetBashoOnly() throws IOException {
        TorayCsvData data = TorayCsvReader.read(sampleCsv(), "A010");

        assertEquals(3, data.byKeiyaku().size());
        assertEquals(10000.0, data.byKeiyaku().get("191352R"), 0.001);
        assertEquals(5000.0, data.byKeiyaku().get("191353R"), 0.001);
        assertFalse(data.byKeiyaku().containsKey("191999X"), "A010P の行は対象外");
        assertEquals(-3560.0, data.total(), 0.001);
    }

    @Test
    @DisplayName("H列が「-」の行は符号を反転して集計する")
    void flipsSignOnMinusRow() throws IOException {
        TorayCsvData data = TorayCsvReader.read(sampleCsv(), "A010");

        assertEquals(-18560.0, data.byKeiyaku().get("190765Y"), 0.001);
        assertEquals(1, data.minusRowCount());
        assertEquals(-18560.0, data.minusTotal(), 0.001);
        List<TorayCsvData.MinusRow> rows = data.minusRows().get("190765Y");
        assertEquals(1, rows.size());
        assertEquals(5, rows.get(0).rowNo(), "ヘッダー・飾り行を含む1始まりの行番号");
        assertEquals("250710", rows.get(0).nyukoDate());
    }

    @Test
    @DisplayName("入庫場所別の合計と入庫月日を保持する")
    void keepsBashoTotalsAndDates() throws IOException {
        TorayCsvData data = TorayCsvReader.read(sampleCsv(), "A010");

        assertEquals(-3560.0, data.byBasho().get("A010"), 0.001);
        assertEquals(20000.0, data.byBasho().get("A010P"), 0.001);
        assertEquals(java.util.Set.of("250705"), data.nyukoDates().get("191352R"));
    }

    @Test
    @DisplayName("入庫場所を A010P にすると湖南工場分だけを集計する")
    void readsKonanBasho() throws IOException {
        TorayCsvData data = TorayCsvReader.read(sampleCsv(), "A010P");

        assertEquals(1, data.byKeiyaku().size());
        assertEquals(20000.0, data.byKeiyaku().get("191999X"), 0.001);
    }

    @Test
    @DisplayName("符号を反映した合計がT行と一致すれば検算警告は出ない")
    void noSubtotalErrorWhenBalanced() throws IOException {
        TorayCsvData data = TorayCsvReader.read(sampleCsv(), "A010");

        assertTrue(data.subtotalErrors().isEmpty());
        assertTrue(data.warnings().isEmpty());
    }

    @Test
    @DisplayName("T小計と合わないブロックを検出し、差額と同額の契約NOを候補に挙げる")
    void detectsSubtotalMismatch() throws IOException {
        String csv = String.join("\n",
                HEADER,
                "A010,N122B,191-352R,250705,X1,ｶﾅ,10,,1000,10000,,",
                ",,191-353R,250706,X2,ｶﾅ,5,,1000,5000,,",
                ",,T,,,,,,,10000,,",
                "");
        TorayCsvData data = TorayCsvReader.read(write("RVSHEET202607.csv", csv), "A010");

        assertEquals(1, data.subtotalErrors().size());
        TorayCsvData.SubtotalError error = data.subtotalErrors().get(0);
        assertEquals(5000.0, error.diff(), 0.001);
        assertEquals("A010", error.bashos());
        assertEquals(1, error.candidates().size());
        assertEquals("191353R", error.candidates().get(0).keiyaku());
    }

    @Test
    @DisplayName("発注No.形式でない行（小計・飾り行）は取り込まない")
    void ignoresNonDataRows() throws IOException {
        String csv = String.join("\n",
                HEADER,
                "A010,N122B,191-352R,250705,X1,ｶﾅ,10,,1000,10000,,",
                ",,GT,,,,,,,10000,,",
                ",,#VALUE!,,,,,,,999,,",
                "-------,,,,,,,,,,,",
                "");
        TorayCsvData data = TorayCsvReader.read(write("RVSHEET202607.csv", csv), "A010");

        assertEquals(1, data.byKeiyaku().size());
        assertEquals(10000.0, data.total(), 0.001);
    }

    @Test
    @DisplayName("H列に想定外の値があればプラス集計のうえ警告する")
    void warnsOnUnexpectedSign() throws IOException {
        String csv = String.join("\n",
                HEADER,
                "A010,N122B,191-352R,250705,X1,ｶﾅ,10,＊,1000,10000,,",
                "");
        TorayCsvData data = TorayCsvReader.read(write("RVSHEET202607.csv", csv), "A010");

        assertEquals(10000.0, data.total(), 0.001);
        assertEquals(1, data.warnings().size());
        assertTrue(data.warnings().get(0).contains("H列"));
    }

    @Test
    @DisplayName("金額列に「金額」ヘッダーが無ければエラーで停止する")
    void failsWhenHeaderIsMissing() throws IOException {
        String csv = String.join("\n",
                "入庫場所,生産場所,発注No.,入庫日,品番,品名カナ,数量,品名,単価,合計,備考1,備考2",
                "A010,N122B,191-352R,250705,X1,ｶﾅ,10,,1000,10000,,",
                "");
        Path path = write("RVSHEET202607.csv", csv);

        VerifyException e = assertThrows(VerifyException.class, () -> TorayCsvReader.read(path, "A010"));
        assertTrue(e.getMessage().contains("金額"));
    }

    @Test
    @DisplayName("対象入庫場所のデータ行が1件も無ければエラーで停止する")
    void failsWhenNoRowsForBasho() throws IOException {
        VerifyException e = assertThrows(VerifyException.class,
                () -> TorayCsvReader.read(sampleCsv(), "A999"));
        assertTrue(e.getMessage().contains("A999"));
    }

    @Test
    @DisplayName("引用符付きセル・セル内改行を正しく解釈する")
    void parsesQuotedFields() {
        List<List<String>> rows = TorayCsvReader.parseCsv("a,\"b,c\",\"d\ne\"\r\nf,g,h\r\n");

        assertEquals(2, rows.size());
        assertEquals(List.of("a", "b,c", "d\ne"), rows.get(0));
        assertEquals(List.of("f", "g", "h"), rows.get(1));
    }

    @Test
    @DisplayName("cp932 の日本語を復号できる")
    void decodesCp932() throws IOException {
        Path path = write("RVSHEET202607.csv", HEADER + "\n");

        assertTrue(TorayCsvReader.decodeCp932(path).contains("入庫場所"));
    }
}
