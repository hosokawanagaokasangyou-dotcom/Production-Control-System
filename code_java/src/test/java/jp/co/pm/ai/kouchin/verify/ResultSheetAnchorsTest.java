package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class ResultSheetAnchorsTest {

    @Test
    @DisplayName("長岡明細の行番号は東レまとめのその行へリンクする")
    void linksNagaokaRow() {
        List<ResultSheetAnchors.Hit> hits = ResultSheetAnchors.find(
                "②長岡明細 413行目 C列が契約NO形式でないが金額あり");
        assertEquals(1, hits.size());
        assertEquals("'②東レまとめ(当月)'!C413", hits.get(0).address());
    }

    @Test
    @DisplayName("シート番地と①CSVと引用シートをそれぞれリンクにする")
    void linksSheetAddresses() {
        List<ResultSheetAnchors.Hit> bang = ResultSheetAnchors.find(
                "東レまとめ!225 ↔ 東レV.C!C83");
        assertEquals(2, bang.size(), bang.toString());
        assertEquals("'②東レまとめ(当月)'!A225", bang.get(0).address());
        assertEquals("'②東レV.C(当月)'!C83", bang.get(1).address());
        assertEquals("'②東レT(当月)'!G6",
                ResultSheetAnchors.find("参照先 =東レT!G6").get(0).address());
        assertEquals("'①東レCSV原本'!A12",
                ResultSheetAnchors.find("①CSV 12〜20行目ブロック").get(0).address());
        assertEquals("'②東レT.V.C(当月)'!C9",
                ResultSheetAnchors.find("②加工賃試算 「東レT.V.C」9行目 C列").get(0).address());
        assertTrue(ResultSheetAnchors.find("データは6行目以降").isEmpty());
    }

    @Test
    @DisplayName("折り返しコメントの行だけ高さを伸ばす")
    void growsWrappedCommentOnly() {
        try (XSSFWorkbook book = new XSSFWorkbook()) {
            XSSFSheet sheet = book.createSheet("検証A");
            sheet.setColumnWidth(0, 12 * 256);
            Row shortRow = sheet.createRow(0);
            Cell shortCell = shortRow.createCell(0);
            shortCell.setCellValue("一致");
            CellStyle plain = book.createCellStyle();
            plain.setWrapText(true);
            shortCell.setCellStyle(plain);
            Row longRow = sheet.createRow(1);
            Cell longCell = longRow.createCell(0);
            longCell.setCellValue("②長岡明細 413行目 C列が契約NO形式でないが金額あり。差額の内訳を確認する。");
            CellStyle wrap = book.createCellStyle();
            wrap.setWrapText(true);
            longCell.setCellStyle(wrap);
            ResultSheetAnchors.link(longCell);
            WrappedRowHeight.fit(sheet);
            assertTrue(longRow.getHeightInPoints() > shortRow.getHeightInPoints() + 10f,
                    "long=" + longRow.getHeightInPoints() + " short=" + shortRow.getHeightInPoints());
            assertEquals("'②東レまとめ(当月)'!C413", longCell.getHyperlink().getAddress());
        } catch (Exception ex) {
            throw new RuntimeException(ex);
        }
    }

    @Test
    @DisplayName("全角は列幅の2として行数を数える")
    void countsFullWidthAsTwo() {
        assertTrue(WrappedRowHeight.lines("あいうえおかきくけこ", 8 * 256) >= 3);
        assertEquals(1, WrappedRowHeight.lines("短い", 40 * 256));
    }
}
