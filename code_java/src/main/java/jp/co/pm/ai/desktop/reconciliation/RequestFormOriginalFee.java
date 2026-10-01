package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.time.Duration;
import java.util.Locale;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.benchmark.GeminiGenerateContentRestClient;
import jp.co.pm.ai.desktop.config.GeminiDispatchModelTryOrderDefaults;

/**
 * 依頼書原本シートの加工賃（円）。
 * 計算式は帳票ごとに違うため、シートのセルを Gemini {@code generateContent} に渡して金額を決める。
 */
public final class RequestFormOriginalFee {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final DataFormatter FORMATTER = new DataFormatter(Locale.JAPAN);
    private static final int MAX_ROWS = 40;
    private static final int MAX_COLS = 50;
    private static final Duration TIMEOUT = Duration.ofSeconds(45);

    /**
     * @param meters 加工賃の計算に使った加工後数量（m）。不明なら 0
     * @param yenPerMeter モデルが読んだ単価合計（円/m）。不明なら 0
     * @param amountYen 加工賃（円）
     * @param reason モデルが使った計算の説明
     */
    public record Result(double meters, double yenPerMeter, double amountYen, String reason) {}

    private RequestFormOriginalFee() {}

    /**
     * シート内容を Gemini に解釈させる。キーが空、または金額が取れないときは null。
     */
    public static Result fromSheet(Sheet sheet, String apiKey) throws IOException, InterruptedException {
        if (sheet == null || apiKey == null || apiKey.isBlank()) {
            return null;
        }
        String grid = sheetText(sheet);
        if (grid.isBlank()) {
            return null;
        }
        String prompt = prompt(sheet.getSheetName(), grid);
        IOException last = null;
        for (String modelId : GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_FALLBACK_TRY_ORDER) {
            GeminiGenerateContentRestClient.FullBodyResult res =
                    GeminiGenerateContentRestClient.generateContentFullBody(
                            apiKey, modelId, prompt, 512, TIMEOUT);
            if (!res.errorSummary().isEmpty()) {
                int status = res.httpStatus();
                if (status == 404 || status == 429) {
                    last = new IOException(res.errorSummary());
                    continue;
                }
                throw new IOException(res.errorSummary());
            }
            Optional<String> text = GeminiGenerateContentRestClient.extractFirstCandidateText(res.body());
            if (text.isEmpty()) {
                last = new IOException("候補テキストが空です（model=" + modelId + "）");
                continue;
            }
            Result parsed = parseModelJson(text.get());
            if (parsed != null) {
                return parsed;
            }
            last = new IOException("加工賃の金額を解釈できません（model=" + modelId + "）");
        }
        if (last != null) {
            throw last;
        }
        return null;
    }

    /** 非空セルを「座標=値」で並べる。Gemini への入力。 */
    static String sheetText(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        int lastRow = Math.min(sheet.getLastRowNum(), MAX_ROWS - 1);
        for (int r = 0; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            int lastCell = Math.min(Math.max(row.getLastCellNum(), 0), MAX_COLS);
            for (int c = 0; c < lastCell; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) {
                    continue;
                }
                String text;
                try {
                    text = FORMATTER.formatCellValue(cell);
                } catch (RuntimeException ex) {
                    continue;
                }
                if (text == null) {
                    continue;
                }
                text = text.replace('\r', ' ').replace('\n', ' ').strip();
                if (text.isEmpty()) {
                    continue;
                }
                if (sb.length() > 0) {
                    sb.append('\n');
                }
                sb.append(cell.getAddress().formatAsString()).append('=').append(text);
            }
        }
        return sb.toString();
    }

    /** モデル応答の JSON から加工賃を読む。金額が無いときは null。 */
    public static Result parseModelJson(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String s = raw.strip();
        int fence = s.indexOf('{');
        int end = s.lastIndexOf('}');
        if (fence < 0 || end <= fence) {
            return null;
        }
        JsonNode node;
        try {
            node = JSON.readTree(s.substring(fence, end + 1));
        } catch (IOException ex) {
            return null;
        }
        if (node == null || !node.has("amountYen") || node.get("amountYen").isNull()) {
            return null;
        }
        double amount = node.get("amountYen").asDouble(Double.NaN);
        if (Double.isNaN(amount) || amount <= 0.0) {
            return null;
        }
        double meters = node.path("meters").asDouble(0.0);
        double rate = node.path("yenPerMeter").asDouble(0.0);
        if (Double.isNaN(meters) || meters < 0.0) {
            meters = 0.0;
        }
        if (Double.isNaN(rate) || rate < 0.0) {
            rate = 0.0;
        }
        String reason = node.path("reason").asText("").replace('\r', ' ').replace('\n', ' ').strip();
        return new Result(meters, rate, amount, reason);
    }

    /** 加工後数量の決め方。テストからも参照する。 */
    static String quantityRules() {
        return "加工賃（円）は、加工後の数量（m）に単価（円/m）を掛けて求める。"
                + "加工前の数量から加工後の数量へは、工程ごとに次を適用する。\n"
                + "・スライス: 厚みを1/2にする。加工後の数量は2倍。\n"
                + "・スリットで端部をトリミングする場合: 加工後の数量は変わらない。\n"
                + "・二つにスリット、三つにスリットなど分割する場合: 加工後の数量は分割数の整数倍（2倍、3倍）。\n"
                + "工程が複数あるときは、シートの並びと注記に沿って数量の変化を積み上げる。"
                + "注記で不要とされた工程の単価は掛けない。\n";
    }

    private static String prompt(String sheetName, String grid) {
        return "加工依頼書の1シートです。加工賃（円）の計算方法は帳票によって違います。"
                + "見出し・注記・数値から、このシートが意図する加工賃の金額を求めてください。"
                + "列位置を決め打ちせず、書かれている内容に従ってください。\n"
                + quantityRules()
                + "meters には加工賃の計算に使った加工後数量（m）を入れる。\n"
                + "出力は JSON オブジェクト1つのみ。説明文やコードフェンスは禁止。\n"
                + "{\"amountYen\":数値,\"meters\":数値またはnull,\"yenPerMeter\":数値またはnull,\"reason\":\"計算の短い説明\"}\n"
                + "金額が判断できないときは {\"amountYen\":null,\"meters\":null,\"yenPerMeter\":null,\"reason\":\"理由\"}\n\n"
                + "【シート名】" + (sheetName == null ? "" : sheetName) + "\n"
                + "【セル】\n"
                + grid
                + "\n";
    }
}
