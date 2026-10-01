package jp.co.pm.ai.kouchin.verify;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import jp.co.pm.ai.desktop.reconciliation.JuchuTransferValueNormalizer;
import jp.co.pm.ai.desktop.reconciliation.RequestFormOriginalFee;

/**
 * 依頼書原本の演算結果を見て、長岡産業側の②を直すかどうかを契約NOごとに書く。
 */
public final class NagaokaExcelFixAdvice {

    public static final String SHEET = "長岡側の直し方";

    /**
     * @param action 長岡側Excelへの直し方。直さないときはその理由
     */
    public record Line(
            String irai,
            String keiyaku,
            Double amount1,
            Double amount2,
            Double originalYen,
            String judge,
            String action) {}

    private NagaokaExcelFixAdvice() {}

    public static List<Line> lines(VerifyResult result, RequestFormOriginalAttacher.Plan originals) {
        if (result == null) {
            return List.of();
        }
        double tol = tolerance(result);
        List<Line> out = new ArrayList<>();
        Set<String> covered = new LinkedHashSet<>();
        if (result.recordsA() != null) {
            for (RecordA rec : result.recordsA()) {
                if (rec == null) {
                    continue;
                }
                covered.add(Norm.keiyaku(rec.keiyaku()));
                List<String> irais = RequestFormOriginalAttacher.splitIrai(rec.iraiNo());
                Bucket bucket = new Bucket(displayIrai(rec.iraiNo(), irais));
                bucket.add(rec);
                bucket.fixes.addAll(fixesFor(result, rec.keiyaku(), rec.iraiNo()));
                Line line = bucket.toLine(result.profile(), feeFor(result, originals, rec.iraiNo(), rec.keiyaku()), tol);
                if (line != null) {
                    out.add(line);
                }
            }
        }
        out.addAll(linesOnlyD(result, originals, covered, tol));
        return List.copyOf(out);
    }

    public static String principle(FactoryProfile profile) {
        boolean kokubu = profile != null && profile.id() == FactoryId.KOKUBU;
        String target = kokubu
                ? "長岡産業側で直すのは②後加工工賃明細です。"
                        + "「東レまとめ」のAAやGを直接書き換えない。"
                        + "単価は元シート（東レT、東レV.C、東レY、東レW.E）のH〜Xを直し、まとめは参照の再計算に任せる。"
                : "長岡産業側で直すのは②加工賃試算です。"
                        + "加工賃の合計セルを値で上書きせず、種類別の単価を直す。";
        return target
                + "正とする金額は依頼書原本の演算結果（配台の加工内容順。左が先。"
                + "スライス単価は半額。原反と同じタイプで長さだけ半分の行は分割のみ。ECでタイプが変わった行はスライス半額・EC・分割。端部トリミングのスリットは数量据え置き、分割スリットは整数倍）。"
                + "依頼書の加工1・加工2の並びは工程順ではない。"
                + "比較と直し額は契約NOごと。同じ依頼NOの契約を合算しない。"
                + "②が原本と違うときだけ長岡側を直す。"
                + "②が原本と一致し①だけ違うときはExcelを直さず、差額を東レへ報告する。";
    }

    /** 同じ依頼NOにぶら下がる契約NOの数。検証Aと検証Dの両方を見る。 */
    public static int contractCount(VerifyResult result, List<String> irais) {
        if (result == null || irais == null || irais.isEmpty()) {
            return 0;
        }
        Set<String> want = iraiKeys(irais);
        Set<String> contracts = new LinkedHashSet<>();
        if (result.recordsA() != null) {
            for (RecordA rec : result.recordsA()) {
                if (rec == null || !sharesIrai(rec.iraiNo(), want)) {
                    continue;
                }
                String key = Norm.keiyaku(rec.keiyaku());
                if (!key.isEmpty()) {
                    contracts.add(key);
                }
            }
        }
        MatomeCheckResult d = result.checkD();
        if (d != null && d.rows() != null) {
            for (MatomeCheckResult.MatomeRow row : d.rows()) {
                if (row == null || !sharesIrai(row.irai(), want)) {
                    continue;
                }
                String key = Norm.keiyaku(row.keiyaku());
                if (!key.isEmpty()) {
                    contracts.add(key);
                }
            }
        }
        return contracts.size();
    }

    private static RequestFormOriginalFee.Result feeFor(
            VerifyResult result,
            RequestFormOriginalAttacher.Plan originals,
            String iraiRaw,
            String keiyaku) {
        if (originals == null) {
            return null;
        }
        List<String> irais = RequestFormOriginalAttacher.splitIrai(iraiRaw);
        Double yen = originals.amountFor(irais, keiyaku, contractCount(result, irais));
        if (yen == null) {
            return null;
        }
        return new RequestFormOriginalFee.Result(0, 0, yen, originals.reasonFor(irais, keiyaku));
    }

    private static List<String> fixesFor(VerifyResult result, String keiyaku, String iraiRaw) {
        MatomeCheckResult d = result.checkD();
        if (d == null || d.isSkipped() || d.rows() == null) {
            return List.of();
        }
        String want = Norm.keiyaku(keiyaku);
        Set<String> irais = iraiKeys(RequestFormOriginalAttacher.splitIrai(iraiRaw));
        List<String> fixes = new ArrayList<>();
        for (MatomeCheckResult.MatomeRow row : d.rows()) {
            if (row == null || !sharesIrai(row.irai(), irais)) {
                continue;
            }
            String key = Norm.keiyaku(row.keiyaku());
            if (!want.isEmpty() && !want.equals(key)) {
                continue;
            }
            fixes.add(FileDiscovery.joinNonEmpty(" ", row.place(), row.judge(), row.detail()));
        }
        return fixes;
    }

    private static List<Line> linesOnlyD(
            VerifyResult result,
            RequestFormOriginalAttacher.Plan originals,
            Set<String> covered,
            double tol) {
        MatomeCheckResult d = result.checkD();
        if (d == null || d.isSkipped() || d.rows() == null) {
            return List.of();
        }
        Map<String, Bucket> buckets = new LinkedHashMap<>();
        for (MatomeCheckResult.MatomeRow row : d.rows()) {
            if (row == null) {
                continue;
            }
            String key = Norm.keiyaku(row.keiyaku());
            if (!key.isEmpty() && covered.contains(key)) {
                continue;
            }
            List<String> irais = RequestFormOriginalAttacher.splitIrai(row.irai());
            String id = key.isEmpty()
                    ? "依頼:" + String.join(",", iraiKeys(irais))
                    : key;
            Bucket bucket = buckets.computeIfAbsent(id, k -> new Bucket(displayIrai(row.irai(), irais)));
            if (row.keiyaku() != null && !row.keiyaku().isBlank() && !bucket.keiyaku.contains(row.keiyaku())) {
                bucket.keiyaku.add(row.keiyaku());
            }
            bucket.fixes.add(FileDiscovery.joinNonEmpty(" ", row.place(), row.judge(), row.detail()));
            for (String irai : irais) {
                String iraiKey = JuchuTransferValueNormalizer.normalizeKey(irai);
                if (!iraiKey.isEmpty() && !bucket.keys.contains(iraiKey)) {
                    bucket.keys.add(iraiKey);
                }
            }
        }
        List<Line> out = new ArrayList<>();
        for (Bucket bucket : buckets.values()) {
            String keiyaku = bucket.keiyaku.isEmpty() ? "" : bucket.keiyaku.get(0);
            Line line = bucket.toLine(
                    result.profile(), feeFor(result, originals, bucket.irai, keiyaku), tol);
            if (line != null) {
                out.add(line);
            }
        }
        return out;
    }

    private static Set<String> iraiKeys(List<String> irais) {
        Set<String> want = new LinkedHashSet<>();
        if (irais == null) {
            return want;
        }
        for (String irai : irais) {
            String key = JuchuTransferValueNormalizer.normalizeKey(irai);
            if (!key.isEmpty()) {
                want.add(key);
            }
        }
        return want;
    }

    private static boolean sharesIrai(String raw, Set<String> want) {
        if (want == null || want.isEmpty()) {
            return false;
        }
        for (String irai : RequestFormOriginalAttacher.splitIrai(raw)) {
            if (want.contains(JuchuTransferValueNormalizer.normalizeKey(irai))) {
                return true;
            }
        }
        return false;
    }

    private static String displayIrai(String raw, List<String> irais) {
        if (irais.size() == 1) {
            return irais.get(0);
        }
        return raw == null ? "" : raw;
    }

    private static double tolerance(VerifyResult result) {
        if (result.info() != null && result.info().containsKey("許容差")) {
            return result.dbl("許容差");
        }
        return 0.5;
    }

    private static final class Bucket {
        private final String irai;
        private final List<String> keys = new ArrayList<>();
        private final List<String> keiyaku = new ArrayList<>();
        private final List<String> judges = new ArrayList<>();
        private final List<String> fixes = new ArrayList<>();
        private double amount1;
        private double amount2;
        private boolean has1;
        private boolean has2;

        private Bucket(String irai) {
            this.irai = irai == null ? "" : irai;
            String key = JuchuTransferValueNormalizer.normalizeKey(this.irai);
            if (!key.isEmpty()) {
                keys.add(key);
            }
        }

        private void add(RecordA rec) {
            if (rec.keiyaku() != null && !rec.keiyaku().isBlank() && !keiyaku.contains(rec.keiyaku())) {
                keiyaku.add(rec.keiyaku());
            }
            if (rec.judge() != null && !rec.judge().isBlank() && !judges.contains(rec.judge())) {
                judges.add(rec.judge());
            }
            if (rec.amount1() != null) {
                amount1 += rec.amount1();
                has1 = true;
            }
            if (rec.amount2() != null) {
                amount2 += rec.amount2();
                has2 = true;
            }
            for (String iraiNo : RequestFormOriginalAttacher.splitIrai(rec.iraiNo())) {
                String key = JuchuTransferValueNormalizer.normalizeKey(iraiNo);
                if (!key.isEmpty() && !keys.contains(key)) {
                    keys.add(key);
                }
            }
        }

        private Line toLine(FactoryProfile profile, RequestFormOriginalFee.Result fee, double tol) {
            Double original = fee == null ? null : fee.amountYen();
            boolean nagaokaDiffers = original != null && has2 && Math.abs(amount2 - original) > tol;
            boolean nagaokaMissing = original != null && !has2 && has1;
            boolean torayDiffers = original != null && has1 && Math.abs(amount1 - original) > tol;
            boolean internal = !fixes.isEmpty();
            boolean anomaly = false;
            for (String judge : judges) {
                if (RequestFormOriginalAttacher.anomalyA(judge)) {
                    anomaly = true;
                    break;
                }
            }
            if (original == null && !internal && !anomaly) {
                return null;
            }
            if (original != null && !nagaokaDiffers && !nagaokaMissing && !torayDiffers && !internal && !anomaly) {
                return null;
            }
            String judge = judges.isEmpty() ? (internal ? "検証D" : "") : String.join("、", judges);
            return new Line(
                    irai,
                    String.join(", ", keiyaku),
                    has1 ? amount1 : null,
                    has2 ? amount2 : null,
                    original,
                    judge,
                    action(profile, fee, tol, nagaokaDiffers, nagaokaMissing, torayDiffers));
        }

        private String action(
                FactoryProfile profile,
                RequestFormOriginalFee.Result fee,
                double tol,
                boolean nagaokaDiffers,
                boolean nagaokaMissing,
                boolean torayDiffers) {
            StringBuilder sb = new StringBuilder();
            if (fee == null) {
                sb.append("この契約NOの原本加工賃が無いため、金額の合わせ先は出せない。");
                sb.append("同じ依頼NOの契約を合算した金額は使わない。");
            } else if (nagaokaMissing) {
                sb.append("長岡側の②にこの契約が無い。原本加工賃 ")
                        .append(Fmt.n0(fee.amountYen()))
                        .append(" 円を②へ追加する。");
            } else if (nagaokaDiffers) {
                sb.append("長岡側の②（")
                        .append(Fmt.n0(amount2))
                        .append(" 円）を原本加工賃 ")
                        .append(Fmt.n0(fee.amountYen()))
                        .append(" 円に合わせる。");
                sb.append(where(profile));
                if (torayDiffers) {
                    sb.append("合わせたあとの①との差 ")
                            .append(Fmt.n0(amount1 - fee.amountYen()))
                            .append(" 円は東レへ報告する。");
                } else if (has1 && Math.abs(amount1 - fee.amountYen()) <= tol) {
                    sb.append("①東レは原本と一致している。直すのは長岡側だけ。");
                }
            } else if (has2) {
                sb.append("長岡側の②は原本加工賃と一致している。Excelの金額は直さない。");
                if (torayDiffers) {
                    sb.append("①との差 ")
                            .append(Fmt.n0(amount1 - fee.amountYen()))
                            .append(" 円は東レへ報告する。");
                }
            }
            if (!fixes.isEmpty()) {
                if (sb.length() > 0) {
                    sb.append(' ');
                }
                sb.append("検証Dの修正箇所: ").append(String.join(" / ", fixes));
            }
            if (fee != null && fee.reason() != null && !fee.reason().isBlank()) {
                sb.append(" 演算: ").append(fee.reason());
            }
            return sb.toString().strip();
        }

        private static String where(FactoryProfile profile) {
            if (profile != null && profile.id() == FactoryId.KOKUBU) {
                return "東レまとめのAAは直接直さない。元シートのH〜X（種類別単価）を直す。";
            }
            return "加工賃の合計セルは直接直さない。種類別の単価を直す。";
        }
    }
}
