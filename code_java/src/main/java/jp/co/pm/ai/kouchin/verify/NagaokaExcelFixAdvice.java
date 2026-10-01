package jp.co.pm.ai.kouchin.verify;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.reconciliation.JuchuTransferValueNormalizer;
import jp.co.pm.ai.desktop.reconciliation.RequestFormOriginalFee;

/**
 * 依頼書原本の演算結果を見て、長岡産業側の②を直すかどうかを依頼NOごとに書く。
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
        Map<String, Bucket> buckets = new LinkedHashMap<>();
        if (result.recordsA() != null) {
            for (RecordA rec : result.recordsA()) {
                if (rec == null) {
                    continue;
                }
                List<String> irais = RequestFormOriginalAttacher.splitIrai(rec.iraiNo());
                String key = irais.size() == 1
                        ? JuchuTransferValueNormalizer.normalizeKey(irais.get(0))
                        : JuchuTransferValueNormalizer.normalizeKey(rec.iraiNo());
                if (key.isEmpty()) {
                    key = rec.keiyaku() == null ? "" : "契約:" + rec.keiyaku();
                }
                Bucket bucket = buckets.computeIfAbsent(key, k -> new Bucket(displayIrai(rec.iraiNo(), irais)));
                bucket.add(rec);
            }
        }
        attachMatome(result, buckets);
        List<Line> out = new ArrayList<>();
        for (Bucket bucket : buckets.values()) {
            RequestFormOriginalFee.Result fee = feeOf(originals, bucket);
            Line line = bucket.toLine(result.profile(), fee, tol);
            if (line != null) {
                out.add(line);
            }
        }
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
                + "スライスは加工後数量2倍、端部トリミングのスリットは数量据え置き、分割スリットは整数倍）。"
                + "依頼書の加工1・加工2の並びは工程順ではない。"
                + "②が原本と違うときだけ長岡側を直す。"
                + "②が原本と一致し①だけ違うときはExcelを直さず、差額を東レへ報告する。";
    }

    private static void attachMatome(VerifyResult result, Map<String, Bucket> buckets) {
        MatomeCheckResult d = result.checkD();
        if (d == null || d.isSkipped() || d.rows() == null) {
            return;
        }
        for (MatomeCheckResult.MatomeRow row : d.rows()) {
            if (row == null) {
                continue;
            }
            List<String> irais = RequestFormOriginalAttacher.splitIrai(row.irai());
            if (irais.isEmpty()) {
                continue;
            }
            for (String irai : irais) {
                String key = JuchuTransferValueNormalizer.normalizeKey(irai);
                if (key.isEmpty()) {
                    continue;
                }
                Bucket bucket = buckets.computeIfAbsent(key, k -> new Bucket(irai));
                bucket.fixes.add(FileDiscovery.joinNonEmpty(" ", row.place(), row.judge(), row.detail()));
            }
        }
    }

    private static RequestFormOriginalFee.Result feeOf(
            RequestFormOriginalAttacher.Plan originals, Bucket bucket) {
        if (originals == null || bucket.keys.isEmpty()) {
            return null;
        }
        for (String key : bucket.keys) {
            RequestFormOriginalAttacher.Item item = originals.byKey(key);
            if (item != null && item.fee() != null) {
                return item.fee();
            }
        }
        return null;
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
            if (original == null && !internal) {
                return null;
            }
            if (original != null && !nagaokaDiffers && !nagaokaMissing && !torayDiffers && !internal) {
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
                sb.append("依頼書の演算結果がないため、金額の合わせ先は出せない。");
            } else if (nagaokaMissing) {
                sb.append("長岡側の②にこの依頼が無い。原本加工賃 ")
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
