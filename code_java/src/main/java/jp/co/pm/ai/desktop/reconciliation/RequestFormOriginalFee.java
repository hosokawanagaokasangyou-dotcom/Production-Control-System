package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
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
     * 契約NO 1件分の加工賃。
     *
     * @param keiyaku 契約NO
     * @param meters その契約の数量（m）
     * @param yenPerMeter その契約の単価（円/m）
     * @param amountYen その契約の加工賃（円）
     * @param reason その契約の計算説明
     */
    public record ContractFee(String keiyaku, double meters, double yenPerMeter, double amountYen, String reason) {}

    /**
     * @param meters 加工賃の計算に使った加工後数量（m）。不明なら 0
     * @param yenPerMeter モデルが読んだ単価合計（円/m）。不明なら 0
     * @param amountYen 依頼書全体の加工賃（円）。契約が複数なら {@code contracts} の合計
     * @param reason モデルが使った計算の説明
     * @param contracts 契約NOごとの加工賃。モデルが分けなかったときは空
     */
    public record Result(
            double meters,
            double yenPerMeter,
            double amountYen,
            String reason,
            List<ContractFee> contracts) {

        public Result(double meters, double yenPerMeter, double amountYen, String reason) {
            this(meters, yenPerMeter, amountYen, reason, List.of());
        }

        public Result {
            if (reason == null) {
                reason = "";
            }
            contracts = contracts == null ? List.of() : List.copyOf(contracts);
        }
    }

    private RequestFormOriginalFee() {}

    /**
     * シート内容を Gemini に解釈させる。キーが空、または金額が取れないときは null。
     */
    public static Result fromSheet(Sheet sheet, String apiKey, String dispatchProcessOrder)
            throws IOException, InterruptedException {
        if (sheet == null || apiKey == null || apiKey.isBlank()) {
            return null;
        }
        String grid = sheetText(sheet);
        if (grid.isBlank()) {
            return null;
        }
        String prompt = prompt(sheet.getSheetName(), grid, dispatchProcessOrder);
        IOException last = null;
        for (String modelId : GeminiDispatchModelTryOrderDefaults.PLANNING_CORE_FALLBACK_TRY_ORDER) {
            GeminiGenerateContentRestClient.FullBodyResult res =
                    GeminiGenerateContentRestClient.generateContentFullBody(
                            apiKey, modelId, prompt, 1024, TIMEOUT);
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
        if (node == null) {
            return null;
        }
        List<ContractFee> contracts = contractFees(node.get("lines"));
        double amount = node.path("amountYen").asDouble(Double.NaN);
        if (Double.isNaN(amount) || amount <= 0.0) {
            amount = sumContracts(contracts);
        }
        if (Double.isNaN(amount) || amount <= 0.0) {
            return null;
        }
        double meters = node.path("meters").asDouble(0.0);
        double rate = node.path("yenPerMeter").asDouble(0.0);
        if (Double.isNaN(meters) || meters < 0.0) {
            meters = sumMeters(contracts);
        }
        if (Double.isNaN(rate) || rate < 0.0) {
            rate = 0.0;
        }
        String reason = node.path("reason").asText("").replace('\r', ' ').replace('\n', ' ').strip();
        return new Result(meters, rate, amount, reason, contracts);
    }

    private static List<ContractFee> contractFees(JsonNode lines) {
        if (lines == null || !lines.isArray()) {
            return List.of();
        }
        List<ContractFee> out = new ArrayList<>();
        for (JsonNode line : lines) {
            if (line == null || !line.isObject()) {
                continue;
            }
            String keiyaku = line.path("keiyaku").asText("").replace('\r', ' ').replace('\n', ' ').strip();
            double yen = line.path("amountYen").asDouble(Double.NaN);
            if (keiyaku.isEmpty() || Double.isNaN(yen) || yen <= 0.0) {
                continue;
            }
            double meters = line.path("meters").asDouble(0.0);
            double rate = line.path("yenPerMeter").asDouble(0.0);
            if (Double.isNaN(meters) || meters < 0.0) {
                meters = 0.0;
            }
            if (Double.isNaN(rate) || rate < 0.0) {
                rate = 0.0;
            }
            String reason = line.path("reason").asText("").replace('\r', ' ').replace('\n', ' ').strip();
            out.add(new ContractFee(keiyaku, meters, rate, yen, reason));
        }
        return List.copyOf(out);
    }

    private static double sumContracts(List<ContractFee> contracts) {
        double sum = 0.0;
        for (ContractFee line : contracts) {
            sum += line.amountYen();
        }
        return sum;
    }

    private static double sumMeters(List<ContractFee> contracts) {
        double sum = 0.0;
        for (ContractFee line : contracts) {
            sum += line.meters();
        }
        return sum;
    }

    /** 加工後数量の決め方。テストからも参照する。 */
    static String quantityRules() {
        return "加工賃（円）は、製品行ごとの数量(m)に、その行にかかる単価（円/m）を掛けて合計する。"
                + "シートに印字済みの加工賃合計（例: 35.00 と 15.00）は、スライスを半額にする前の足し算になっていることがある。それを正としない。\n"
                + "・スライスの円/mは必ず記載の半額にする。22円なら11円。厚みは1/2になる。"
                + "製品行に数量(m)が書いてあるときはその数量を使う。書いてある数量をさらに一律で2倍しない。\n"
                + "・スリットで端部をトリミングする場合: 加工後の数量は変わらない。\n"
                + "・二つにスリット、三つにスリットなど分割する場合: 加工後の数量は分割数の整数倍（2倍、3倍）。\n"
                + "製品行ごとに、どの工程を含めるかは原反と製品のタイプ・長さで決める。\n"
                + "・原反と同じタイプで、長さだけ半分になっている行（例: 原反AG00・長さ200m → 製品AG00・長さ100m）は分割のみ。\n"
                + "・ECがあるとタイプが変わる（例: 原反AG00 → 製品KG00）。"
                + "タイプが変わった行は、スライス（半額）・EC・分割を含める。\n"
                + "例: スライス22・EC18・分割15、製品200m（KG00）と100m（AG00）、原反AG00長さ200。"
                + "正は (11+18+15)×200 + 15×100 = 10300円。"
                + "200×35 + 100×15 = 8500円は、35円がスライスを半額にしていないので誤り。\n"
                + "工程が複数あるときの順番は、依頼書原本の加工1・加工2の並びではない。"
                + "配台システムの順番（受注ファイルの加工内容。カンマ区切りの左が先）で数量の変化を積み上げる。"
                + "単価は依頼書の各工程の円/mを使い、掛ける順番だけ配台順に従う。"
                + "注記で不要とされた工程の単価は掛けない。"
                + "配台順が無いときは、原本の並びで代用せず金額は出さない。\n"
                + "加工賃は契約NOごとに分ける。1枚の依頼書に契約NOが複数あるときは、"
                + "製品行の契約Ｎｏごとに金額を出し、依頼NOの合計を各契約へコピーしない。\n"
                + "例の10300円で200mの行と100mの行が別契約なら、lines は8800円と1500円の2件、amountYen は10300。\n";
    }

    private static String prompt(String sheetName, String grid, String dispatchProcessOrder) {
        String order = dispatchProcessOrder == null ? "" : dispatchProcessOrder.strip();
        String orderLine = order.isEmpty()
                ? "【配台の工程順】不明。依頼書原本の行順を工程順に使わず、amountYen は null。\n"
                : "【配台の工程順】左が先。受注ファイルの加工内容。依頼書原本の並びは使わない。\n" + order + "\n";
        return "加工依頼書の1シートです。加工賃（円）の計算方法は帳票によって違います。"
                + "見出し・注記・数値から、このシートが意図する加工賃の金額を求めてください。"
                + "列位置を決め打ちせず、書かれている内容に従ってください。\n"
                + quantityRules()
                + orderLine
                + "meters には加工賃の計算に使った加工後数量（m）を入れる。\n"
                + "出力は JSON オブジェクト1つのみ。説明文やコードフェンスは禁止。\n"
                + "{\"amountYen\":数値,\"meters\":数値またはnull,\"yenPerMeter\":数値またはnull,\"reason\":\"計算の短い説明\","
                + "\"lines\":[{\"keiyaku\":\"契約NO\",\"amountYen\":数値,\"meters\":数値またはnull,\"yenPerMeter\":数値またはnull,\"reason\":\"その契約の説明\"}]}\n"
                + "lines はシートにある契約NOを1件ずつ入れる。契約が1件でも lines は1件。\n"
                + "金額が判断できないときは {\"amountYen\":null,\"meters\":null,\"yenPerMeter\":null,\"reason\":\"理由\",\"lines\":[]}\n\n"
                + "【シート名】" + (sheetName == null ? "" : sheetName) + "\n"
                + "【セル】\n"
                + grid
                + "\n";
    }
}
