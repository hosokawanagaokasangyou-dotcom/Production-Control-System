package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.Instant;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class KouchinOutlookDropSupportTest {

    @TempDir
    Path tmp;

    @Test
    @DisplayName("OutlookがTEMPに書いた直後のRVSHEET.csvを拾う")
    void picksRecentTempRvsheetCsv() throws Exception {
        Path csv = tmp.resolve("RVSHEET.csv");
        Files.writeString(csv, "outlook", StandardCharsets.UTF_8);
        Files.setLastModifiedTime(csv, FileTime.from(Instant.parse("2026-09-17T01:58:00Z")));
        List<Path> found = KouchinOutlookDropSupport.recentTempCsvFiles(
                tmp, Instant.parse("2026-09-17T01:58:10Z"));
        assertEquals(1, found.size());
        assertEquals(csv.toAbsolutePath().normalize(), found.get(0).toAbsolutePath().normalize());
    }

    @Test
    @DisplayName("古いTEMPのCSVは使わない")
    void ignoresStaleTempCsv() throws Exception {
        Path csv = tmp.resolve("RVSHEET.csv");
        Files.writeString(csv, "old", StandardCharsets.UTF_8);
        Files.setLastModifiedTime(csv, FileTime.from(Instant.parse("2026-09-17T01:00:00Z")));
        List<Path> found = KouchinOutlookDropSupport.recentTempCsvFiles(
                tmp, Instant.parse("2026-09-17T01:58:00Z"));
        assertTrue(found.isEmpty());
    }

    @Test
    @DisplayName("FileGroupDescriptorWから元ファイル名を読む")
    void parsesOriginalNameFromFileGroupDescriptorW() {
        byte[] bytes = fileGroupDescriptorW("RVSHEET202608.csv");
        assertEquals(List.of("RVSHEET202608.csv"), KouchinOutlookDropSupport.fileNamesFromDescriptor(bytes));
    }

    @Test
    @DisplayName("TEMPのRVSHEET.csvは元のRVSHEETyyyymm名でコピーする")
    void copiesBareTempAsDatedName() throws Exception {
        Path src = tmp.resolve("RVSHEET.csv");
        Files.writeString(src, "body", StandardCharsets.UTF_8);
        Path named = KouchinOutlookDropSupport.withOriginalFileName(src, "RVSHEET202608.csv");
        assertEquals("RVSHEET202608.csv", named.getFileName().toString());
        assertEquals("body", Files.readString(named, StandardCharsets.UTF_8));
    }

    @Test
    @DisplayName("入庫日YYMMDDからRVSHEETyyyyMM.csvを推定する")
    void infersDatedNameFromNyukoYymmdd() throws Exception {
        Path src = tmp.resolve("RVSHEET.csv");
        Files.writeString(src, "A010,x,191-352R,260821,1\nA010,x,191-352R,260824,1\n", StandardCharsets.UTF_8);
        assertEquals("RVSHEET202608.csv", KouchinOutlookDropSupport.datedRvsheetNameFromCsv(src));
        List<Path> named = KouchinOutlookDropSupport.withInferredDatedNames(List.of(src));
        assertEquals(1, named.size());
        assertEquals("RVSHEET202608.csv", named.get(0).getFileName().toString());
        assertEquals("A010,x,191-352R,260821,1\nA010,x,191-352R,260824,1\n",
                Files.readString(named.get(0), StandardCharsets.UTF_8));
    }

    @Test
    @DisplayName("Outlookのexternal-bodyから添付名を読む")
    void parsesOutlookExternalBodyName() {
        assertEquals(List.of("RVSHEET.csv"), KouchinOutlookDropSupport.fileNamesFromContentTypeIds(List.of(
                "message/external-body;access-type=clipboard;index=0;name=\"RVSHEET.csv\"")));
        assertEquals(List.of("RVSHEET.csv"), KouchinOutlookDropSupport.fileNamesFromContentTypeIds(List.of(
                "[message/external-body;access-type=clipboard;index=0;name=\"RVSHEET.csv\"]")));
    }

    @Test
    @DisplayName("Outlookが指名したTEMPのCSVは古くても使う")
    void usesStaleTempWhenNamedByOutlook() throws Exception {
        Path csv = tmp.resolve("RVSHEET.csv");
        Files.writeString(csv, "stale", StandardCharsets.UTF_8);
        Files.setLastModifiedTime(csv, FileTime.from(Instant.parse("2026-09-17T01:00:00Z")));
        assertTrue(KouchinOutlookDropSupport.recentTempCsvFiles(
                tmp, Instant.parse("2026-09-17T01:58:00Z")).isEmpty());
        List<Path> found = KouchinOutlookDropSupport.tempFilesNamed(tmp, List.of("RVSHEET.csv"));
        assertEquals(1, found.size());
        assertEquals(csv.toAbsolutePath().normalize(), found.get(0).toAbsolutePath().normalize());
    }

    @Test
    @DisplayName("RVSHEET (1).csv があればより新しい方を使う")
    void prefersNewerWindowsCopyName() throws Exception {
        Path orig = tmp.resolve("RVSHEET.csv");
        Path copy = tmp.resolve("RVSHEET (1).csv");
        Files.writeString(orig, "old", StandardCharsets.UTF_8);
        Files.writeString(copy, "new", StandardCharsets.UTF_8);
        Files.setLastModifiedTime(orig, FileTime.from(Instant.parse("2026-09-17T01:00:00Z")));
        Files.setLastModifiedTime(copy, FileTime.from(Instant.parse("2026-09-17T01:58:00Z")));
        Path newest = KouchinOutlookDropSupport.newestTempRvsheetCsv(tmp);
        assertEquals(copy.toAbsolutePath().normalize(), newest.toAbsolutePath().normalize());
    }

    private static byte[] fileGroupDescriptorW(String name) {
        int struct = 592;
        ByteBuffer buf = ByteBuffer.allocate(4 + struct).order(ByteOrder.LITTLE_ENDIAN);
        buf.putInt(1);
        buf.position(4 + 72);
        byte[] utf16 = name.getBytes(StandardCharsets.UTF_16LE);
        buf.put(utf16);
        buf.putShort((short) 0);
        return buf.array();
    }
}
