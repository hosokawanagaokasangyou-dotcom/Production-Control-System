package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.NavigableMap;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 検証A（① vs ② 契約NO突合）と検証B（② vs ③ 依頼NO突合）の本体。
 * Python 版 {@code verify_kouchin.py} の {@code main()} の突合ロジックを移植したもの。
 */
public final class VerifyEngine {

    /** Python 正本どおり手動①正を内訳に含める。 */
    static double reportResidual(
            double total1,
            double total2,
            double adjTotal,
            double nextTotal,
            double diffTotal,
            double manual1Total,
            double only1Total,
            double only2Total) {
        return total1 - (total2 + adjTotal + nextTotal + diffTotal + manual1Total + only1Total - only2Total);
    }

    /** 依頼NOの月番号（例 Y7-52 → 7） */
    private static final Pattern IRAI_MONTH_PATTERN = Pattern.compile("^[A-Z]+(\\d{1,2})-");
    /** 枝番（例 C8-3-1 の末尾 -1） */
    private static final Pattern BRANCH_SUFFIX_PATTERN = Pattern.compile("-\\d+$");
    /** ③のファイル名 {@code _yyyymmdd_hhmmss} */
    private static final Pattern ALADDIN_DATE_PATTERN = Pattern.compile("_(\\d{4})(\\d{2})(\\d{2})_");
    private static final DateTimeFormatter TIMESTAMP = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    /** 過去月・翌月の②ファイルから読んだ契約NO情報。 */
    private record MonthData(String label, YearMonthKey ym,
                             Map<String, Double> byKeiyaku, Map<String, Set<String>> keiyakuToIrai) {
    }

    /** 過去月・翌月のヒット。 */
    private record Hit(String label, double amount, Set<String> irais) {
    }

    private record MismatchA(String keiyaku, double amount1, double amount2, double diff) {
    }

    private record ResolvedRow(String origin, String irai, double amount2, double amount3) {
    }

    private record BranchRow(String irai, double amount2, double amount3) {
    }

    private final FactoryProfile profile;
    private final KouchinPaths paths;
    private final double tol;

    private final List<String> warnings = new ArrayList<>();
    private final List<RecordA> recordsA = new ArrayList<>();
    private final List<RecordB> recordsB = new ArrayList<>();
    private final Map<String, Integer> zeroExcluded = new LinkedHashMap<>();
    private final Map<String, String> ctHint = new LinkedHashMap<>();
    private final Map<String, String> pairNotes = new LinkedHashMap<>();
    private final Map<String, String> resolvedB = new LinkedHashMap<>();
    private final Map<String, String> branchOk = new LinkedHashMap<>();
    private final Map<String, String> branchChild = new LinkedHashMap<>();

    private final List<MonthData> prevData = new ArrayList<>();
    private final List<MonthData> nextData = new ArrayList<>();
    private final NavigableMap<YearMonthKey, Map<String, Double>> iraiByYm = new TreeMap<>();
    private final NavigableMap<YearMonthKey, Map<String, Double>> aladdinByYm = new TreeMap<>();

    private YearMonthKey targetYm;
    private Path torayFile;
    private Path source2File;
    private Path aladdinFile;
    private Path dir2;
    private MonthlyFileSet fileSet2;
    private TorayCsvData toray;
    private MoneyMaps source2;
    private AladdinData aladdinData;
    private Map<String, Double> aladdin;
    private MatomeCheckResult checkD;
    private CheckCResult checkC;
    private Map<String, JudgmentCsv.ManualEntry> manualByKeiyaku = Map.of();
    private final List<JudgmentCsv.PriorRow> priorShortfall = new ArrayList<>();
    private final List<String> aladdinOtherFiles = new ArrayList<>();

    // 検証Aの結果
    private final List<MismatchA> mismatchA = new ArrayList<>();
    private final Map<String, Hit> prevHit = new LinkedHashMap<>();
    private final Map<String, Hit> nextHit = new LinkedHashMap<>();
    private List<String> prevAdjust = List.of();
    private List<String> nextMissrec = List.of();
    private List<String> only1 = List.of();
    private List<String> only2 = List.of();
    private int commonA;

    // 検証Bの結果
    private final List<MismatchA> mismatchB = new ArrayList<>();
    private final List<ResolvedRow> resolvedRows = new ArrayList<>();
    private final List<BranchRow> branchRows = new ArrayList<>();
    private List<String> only2b = List.of();
    private List<String> only3b = List.of();
    private int commonB;
    private int branchCommonCount;
    private boolean skipVerifyB;

    private VerifyEngine(FactoryProfile profile, KouchinPaths paths, double tol) {
        this.profile = profile;
        this.paths = paths;
        this.tol = tol;
    }

    /** 1工場分の検証を実行する。 */
    public static VerifyResult run(FactoryProfile profile, KouchinPaths paths, double tol) {
        return new VerifyEngine(profile, paths, tol).execute();
    }

    private VerifyResult execute() {
        loadSources();
        runCheckD();
        verifyA();
        verifyB();
        runCheckC();
        return buildResult();
    }

    // ---------------------------------------------------------------- 読み込み

    private void loadSources() {
        Path dir1 = paths.resolveTorayCsvDir();
        if (!Files.isDirectory(dir1)) {
            throw new VerifyException("①東レ送付CSVのフォルダにアクセスできません: " + dir1);
        }
        torayFile = FileDiscovery.findTorayCsv(dir1);
        targetYm = FileDiscovery.torayTargetYm(torayFile).orElse(null);
        if (targetYm == null) {
            warnings.add("①ファイル名から対象年月(RVSHEETyyyymm)を特定できず、月整合チェックをスキップしました: "
                    + torayFile.getFileName());
        }

        dir2 = paths.source2Dir(profile.id());
        if (!Files.isDirectory(dir2)) {
            throw new VerifyException("②" + profile.name2() + "のフォルダにアクセスできません: " + dir2);
        }
        fileSet2 = FileDiscovery.findSource2(profile, dir2, targetYm);
        source2File = fileSet2.current();

        Path dir3 = paths.source3Dir(profile.id());
        loadSource3(dir3);

        toray = TorayCsvReader.read(torayFile, profile.basho());
        warnings.addAll(toray.warnings());

        source2 = Source2Reader.read(profile, source2File);
        warnings.addAll(source2.warnings());
        if (targetYm != null) {
            iraiByYm.put(targetYm, source2.byIrai());
        }

        for (DatedFile f : fileSet2.prevs()) {
            readOtherMonth(f, prevData);
        }
        for (DatedFile f : fileSet2.nexts()) {
            readOtherMonth(f, nextData);
        }

        buildSubtotalHints();
    }

    /** ③月次実績。対象月データが無い・読めないときは検証BをスキップしてAを続行する。 */
    private void loadSource3(Path dir3) {
        if (!Files.isDirectory(dir3)) {
            markSkipB("③月次実績のフォルダにアクセスできません: " + dir3);
            return;
        }
        try {
            aladdinFile = FileDiscovery.findAladdin(dir3, targetYm);
        } catch (VerifyException e) {
            markSkipB(e.getMessage());
            return;
        }
        try {
            aladdinData = AladdinReader.read(aladdinFile, profile.customer3(), profile.warehouse3Contains());
        } catch (VerifyException e) {
            markSkipB(e.getMessage());
            return;
        }
        if (!AladdinData.coversTargetMonth(aladdinData, targetYm)) {
            String ymText = aladdinData.taishoYm() == null
                    ? "(不明)"
                    : aladdinData.taishoYm().ymLabel();
            markSkipB(aladdinData.count() == 0
                    ? "加工金額行が0件"
                    : "シート対象年月=" + ymText);
            return;
        }
        aladdin = aladdinData.byIrai();
        warnMidMonth(aladdinFile, aladdinData.taishoYm(), true);

        if (targetYm != null) {
            aladdinByYm.put(targetYm, aladdin);
            Map<YearMonthKey, Path> all;
            try {
                all = FileDiscovery.findAladdinAll(dir3);
            } catch (VerifyException e) {
                all = Map.of();
            }
            for (Map.Entry<YearMonthKey, Path> e : new TreeMap<>(all).entrySet()) {
                if (e.getKey().equals(targetYm)) {
                    continue;
                }
                try {
                    aladdinByYm.put(e.getKey(), AladdinReader.read(
                            e.getValue(), profile.customer3(), profile.warehouse3Contains()).byIrai());
                    aladdinOtherFiles.add(e.getValue().getFileName().toString());
                } catch (VerifyException ex) {
                    warnings.add("③他月照会 " + e.getValue().getFileName() + " を読めないためスキップしました: "
                            + ex.getMessage());
                    continue;
                }
                warnMidMonth(e.getValue(), e.getKey(), false);
            }
        }
    }

    private void markSkipB(String detail) {
        skipVerifyB = true;
        if (aladdin == null) {
            aladdin = Map.of();
        }
        if (aladdinData == null) {
            aladdinData = new AladdinData(Map.of(), "", null);
        }
        warnings.add(0, skipBWarning(targetYm, detail));
    }

    static String skipBWarning(YearMonthKey targetYm, String detail) {
        String ym = targetYm == null ? "(不明)" : targetYm.gatsudoLabel();
        String extra = detail == null || detail.isBlank() ? "" : " 詳細: " + detail;
        return "【検証Bスキップ】③月次実績に対象月（" + ym + "）のデータがありません。"
                + "検証B（② vs ③ 依頼NO突合）は実施していません。検証AとExcelは出力済みです。"
                + "③の依頼NO別問合せを対象月で再取得して再実行してください。"
                + extra;
    }

    private void readOtherMonth(DatedFile f, List<MonthData> into) {
        try {
            MoneyMaps m = Source2Reader.read(profile, f.path());
            into.add(new MonthData(f.label(), f.ym(), m.byKeiyaku(), m.keiyakuToIrai()));
            iraiByYm.put(f.ym(), m.byIrai());
        } catch (VerifyException e) {
            warnings.add("②他月明細 " + f.path().getFileName() + " を読めないためスキップしました: " + e.getMessage());
        }
    }

    /** 照会取得日が対象月の月中（28日より前）なら途中実績である旨を警告する。 */
    private void warnMidMonth(Path file, YearMonthKey ym, boolean current) {
        if (ym == null) {
            return;
        }
        Matcher m = ALADDIN_DATE_PATTERN.matcher(file.getFileName().toString());
        if (!m.find()) {
            return;
        }
        int y = Integer.parseInt(m.group(1));
        int mo = Integer.parseInt(m.group(2));
        int day = Integer.parseInt(m.group(3));
        if (y != ym.year() || mo != ym.month() || day >= 28) {
            return;
        }
        if (current) {
            warnings.add(0, "③当月照会(" + file.getFileName() + ")は" + mo + "月" + day + "日時点の途中実績です。"
                    + "月末確定後に取得したファイルを置いて再実行してください"
                    + "(同一対象年月が複数あればファイル名の日時が最新のものを採用します)。"
                    + "途中実績のままでは②のみが実態より多く出ます");
        } else {
            warnings.add("③" + ym.monthLabel() + "照会(" + file.getFileName() + ")は" + day + "日時点の途中実績です。"
                    + "月末確定後に再取得して再実行すると、検証Bの月ずれ解消(累計一致)の精度が上がります");
        }
    }

    /** ①CSVブロック小計(T行)の検算不一致を警告と備考ヒントに落とす。 */
    private void buildSubtotalHints() {
        for (TorayCsvData.SubtotalError e : toray.subtotalErrors()) {
            StringBuilder msg = new StringBuilder("①CSV " + e.blockStart() + "〜" + e.subtotalRow()
                    + "行目ブロック(入庫場所 " + e.bashos() + "): 小計(T=" + Fmt.n0(e.subtotal())
                    + "円)とデータ行合計(" + Fmt.n0(e.dataSum()) + "円)が " + Fmt.s0(e.diff())
                    + "円 不一致。東レ側で集計除外された行(取消等)の疑い");
            if (!e.candidates().isEmpty()) {
                List<String> parts = new ArrayList<>();
                for (TorayCsvData.Candidate c : e.candidates()) {
                    parts.add(c.keiyaku() + "(" + c.basho() + ") " + Fmt.n0(c.amount()) + "円");
                    if (Norm.norm(profile.basho()).equals(c.basho())) {
                        ctHint.put(c.keiyaku(), "①小計(T行)から除外の疑い: " + e.blockStart() + "〜" + e.subtotalRow()
                                + "行目ブロックの小計差 " + Fmt.s0(e.diff()) + "円 と同額");
                    }
                }
                msg.append(" → 差額と同額の契約NO: ").append(String.join(", ", parts));
            } else {
                msg.append("。①原本シートの該当行を確認");
            }
            warnings.add(msg.toString());
        }
    }

    private void runCheckD() {
        if (!profile.hasCheckD()) {
            return;
        }
        checkD = CheckMatome.check(source2File, tol);
        if (checkD.isSkipped()) {
            warnings.add("検証D(②まとめ整合性)をスキップしました: " + checkD.skipped());
            return;
        }
        int errors = checkD.errorCount();
        if (errors > 0) {
            warnings.add("②「" + FactoryProfile.MATOME_SHEET + "」が元4シートとずれています(検証D 要修正 "
                    + errors + "件 / 加工賃合計の差 " + Fmt.s0(checkD.totalGap()) + "円)。"
                    + "②総額の信頼性に関わるため②を修正して再実行すること。詳細は検証Dシート");
        }
    }

    // ---------------------------------------------------------------- 検証A

    private void verifyA() {
        Map<String, Double> t1 = toray.byKeiyaku();
        Map<String, Double> t2 = source2.byKeiyaku();
        Set<String> k1 = new TreeSet<>(t1.keySet());
        Set<String> k2 = new TreeSet<>(t2.keySet());

        List<String> common = new ArrayList<>();
        for (String k : k1) {
            if (k2.contains(k)) {
                common.add(k);
            }
        }
        commonA = common.size();

        Set<String> mismatchKeys = new LinkedHashSet<>();
        JudgmentCsv.ManualLoad manualLoad =
                JudgmentCsv.loadManual(JudgmentCsv.manualFile(paths), profile.label(), targetYm);
        warnings.addAll(manualLoad.warnings());
        manualByKeiyaku = manualLoad.byKeiyaku();
        List<MismatchA> manualApplied = new ArrayList<>();
        for (String k : common) {
            double d = t1.get(k) - t2.get(k);
            if (Math.abs(d) > tol) {
                JudgmentCsv.ManualEntry man = manualByKeiyaku.get(k);
                if (man != null) {
                    manualApplied.add(new MismatchA(k, t1.get(k), t2.get(k), d));
                } else {
                    mismatchA.add(new MismatchA(k, t1.get(k), t2.get(k), d));
                    mismatchKeys.add(k);
                }
            }
        }
        for (String k : new TreeSet<>(manualByKeiyaku.keySet())) {
            boolean applied = manualApplied.stream().anyMatch(m -> m.keiyaku().equals(k));
            if (applied) {
                continue;
            }
            if (k1.contains(k) && k2.contains(k)) {
                continue;
            }
            String state = k1.contains(k) ? "①のみ" : k2.contains(k) ? "②のみ" : "①②とも無し";
            warnings.add(JudgmentCsv.MANUAL_FILE + ": 契約NO " + k + " は当月の検証Aで不一致ではありません("
                    + state + ")。手動判定は適用されませんでした(契約NOの誤記?)");
        }
        buildPairNotes();

        for (String k : common) {
            if (isZero(t1.get(k)) && isZero(t2.get(k))) {
                bumpZero("A一致");
                continue;
            }
            double d = t1.get(k) - t2.get(k);
            boolean mismatch = mismatchKeys.contains(k);
            JudgmentCsv.ManualEntry man = mismatch ? null : manualByKeiyaku.get(k);
            String judge;
            if (mismatch) {
                judge = Judge.MISMATCH;
            } else if (man != null && Math.abs(d) > tol) {
                judge = man.side() == 1 ? Judge.MANUAL_1 : Judge.MANUAL_2;
            } else {
                judge = Judge.MATCH;
            }
            String note;
            if (Judge.MANUAL_1.equals(judge) || Judge.MANUAL_2.equals(judge)) {
                note = manualNote(k, t1.get(k), t2.get(k), d, man);
            } else if (mismatch) {
                note = FileDiscovery.joinNonEmpty(" / ", pairNotes.get(k), minusNote(k), ctHint.get(k));
            } else {
                note = FileDiscovery.joinNonEmpty(" / ", pairNotes.get(k));
            }
            addRecordA(k, t1.get(k), t2.get(k), d, judge, note, null);
        }

        // ①のみ → 過去月まとめ（前月調整）→ 翌月まとめ（翌月記載）→ 真の①のみ
        List<String> only1Keys = new ArrayList<>();
        for (String k : k1) {
            if (!k2.contains(k)) {
                only1Keys.add(k);
            }
        }

        List<String> prevAdjustKeys = new ArrayList<>();
        List<String> rest = new ArrayList<>();
        for (String k : only1Keys) {
            Hit hit = findHit(prevData, k);
            if (hit != null) {
                prevHit.put(k, hit);
                prevAdjustKeys.add(k);
            } else {
                rest.add(k);
            }
        }
        prevAdjust = List.copyOf(prevAdjustKeys);
        for (String k : prevAdjust) {
            Hit hit = prevHit.get(k);
            double d = t1.get(k) - hit.amount();
            String note = hit.label() + "の依頼NO " + String.join(", ", sorted(hit.irais()))
                    + " / 当時" + profile.amt2() + " " + Fmt.n0(hit.amount()) + "円 / ①との差 " + Fmt.s0(d) + "円";
            note = FileDiscovery.joinNonEmpty(" / ", note, nyukoNote(k));
            addRecordA(k, t1.get(k), null, null, Judge.PREV_ADJUST, note, hit);
        }

        applyPriorShortfalls(k1, k2);

        List<String> nextKeys = new ArrayList<>();
        List<String> rest2 = new ArrayList<>();
        for (String k : rest) {
            Hit hit = Math.abs(t1.get(k)) > tol ? findHit(nextData, k) : null;
            if (hit != null) {
                nextHit.put(k, hit);
                nextKeys.add(k);
            } else {
                rest2.add(k);
            }
        }
        nextMissrec = List.copyOf(nextKeys);
        String curLabel = targetYm != null ? targetYm.monthLabel() + "度" : "当月";
        for (String k : nextMissrec) {
            Hit hit = nextHit.get(k);
            double d = t1.get(k) - hit.amount();
            String note = hit.label() + "の依頼NO " + String.join(", ", sorted(hit.irais()))
                    + " に記載あり(" + profile.amt2() + " " + Fmt.n0(hit.amount()) + "円 / ①との差 "
                    + Fmt.s0(d) + "円)。本来は" + curLabel + "の②に記載すべきもの → ②" + curLabel
                    + "へ追記し" + hit.label() + "から削除する修正が必要";
            note = FileDiscovery.joinNonEmpty(" / ", note, nyukoNote(k));
            addRecordA(k, t1.get(k), null, null, Judge.NEXT_MONTH, note, hit);
        }

        only1 = reportOnlyA(rest2, t1, Judge.ONLY_1,
                "②" + profile.name2() + "に記載なし(過去月・翌月にも無し)", true);

        List<String> only2Keys = new ArrayList<>();
        for (String k : k2) {
            if (!k1.contains(k)) {
                only2Keys.add(k);
            }
        }
        only2 = reportOnlyA(only2Keys, t2, Judge.ONLY_2, "①東レCSVに検収なし", false);

        for (MoneyMaps.BadKeiyaku bad : source2.badKeiyaku()) {
            if (Math.abs(bad.amount()) > tol) {
                addRecordA(bad.value(), null, bad.amount(), null, Judge.BAD_FORMAT,
                        "②" + profile.name2() + " " + bad.place() + "行目 C列が契約NO形式でないが金額あり", null);
            }
        }
    }

    private void buildPairNotes() {
        Map<String, List<MismatchA>> byIrai = new LinkedHashMap<>();
        for (MismatchA m : mismatchA) {
            for (String irai : source2.iraiOf(m.keiyaku())) {
                byIrai.computeIfAbsent(irai, k -> new ArrayList<>()).add(m);
            }
        }
        for (Map.Entry<String, List<MismatchA>> e : byIrai.entrySet()) {
            List<MismatchA> items = e.getValue();
            if (items.size() < 2) {
                continue;
            }
            double net = 0.0;
            double gross = 0.0;
            for (MismatchA m : items) {
                net += m.diff();
                gross += Math.abs(m.diff());
            }
            if (Math.abs(net) < gross * 0.5) {
                for (MismatchA m : items) {
                    pairNotes.put(m.keiyaku(), "契約間振替の疑い: 同一依頼NO " + e.getKey() + " 内の不一致"
                            + items.size() + "件と相殺関係 (ネット差額 " + Fmt.s0(net) + "円)");
                }
            }
        }
    }

    /** 片側のみのキーを記録する。金額0円のキーは無効行として除外し、金額ありのキーを返す。 */
    private List<String> reportOnlyA(List<String> keys, Map<String, Double> amounts,
                                     String judge, String note, boolean withNyukoNote) {
        List<String> nonzero = new ArrayList<>();
        int zeros = 0;
        for (String k : keys) {
            if (Math.abs(amounts.get(k)) > tol) {
                nonzero.add(k);
            } else {
                zeros++;
            }
        }
        if (zeros > 0) {
            bumpZero("A" + judge, zeros);
        }
        for (String k : nonzero) {
            String full = withNyukoNote ? FileDiscovery.joinNonEmpty(" / ", note, nyukoNote(k)) : note;
            if (Judge.ONLY_1.equals(judge)) {
                addRecordA(k, amounts.get(k), null, null, judge, full, null);
            } else {
                addRecordA(k, null, amounts.get(k), null, judge, full, null);
            }
        }
        return List.copyOf(nonzero);
    }

    private void addRecordA(String keiyaku, Double a1, Double a2, Double diff,
                            String judge, String note, Hit hit) {
        recordsA.add(new RecordA(keiyaku, iraiLabel(keiyaku, judge, hit), a1, a2, diff,
                reportAmount(judge, a1, a2, diff), judge, note));
    }

    /** 検証Aの依頼NO列。前月調整・翌月記載は「(7月)依頼NO」のような月ラベル付き。 */
    private String iraiLabel(String keiyaku, String judge, Hit hit) {
        if (hit != null && (Judge.PREV_ADJUST.equals(judge) || Judge.NEXT_MONTH.equals(judge))) {
            String tag = YearMonthKey.parseGatsudo(hit.label())
                    .map(ym -> "(" + ym.monthLabel() + ")")
                    .orElse("(他月)");
            List<String> labelled = new ArrayList<>();
            for (String x : sorted(hit.irais())) {
                labelled.add(tag + x);
            }
            return String.join(", ", labelled);
        }
        return String.join(", ", sorted(source2.iraiOf(keiyaku)));
    }

    /**
     * 「報告する過不足(当月)」への計上額。
     * 不一致=差額(①-②) / 翌月記載・①のみ=①金額 / ②のみ=-②金額 / それ以外は空欄。
     */
    private static Double reportAmount(String judge, Double a1, Double a2, Double diff) {
        if (Judge.MISMATCH.equals(judge) || Judge.MANUAL_2.equals(judge) || Judge.PREV_GAP.equals(judge)) {
            return diff;
        }
        if (Judge.NEXT_MONTH.equals(judge) || Judge.ONLY_1.equals(judge)) {
            return a1;
        }
        if (Judge.ONLY_2.equals(judge)) {
            return a2 == null ? null : -a2;
        }
        return null;
    }

    private String manualNote(String k, double v1, double v2, double d, JudgmentCsv.ManualEntry man) {
        String ymLabel = targetYm != null ? targetYm.gatsudoLabel() : "(不明)";
        String act = man.side() == 1
                ? "②を①金額 " + Fmt.n0(v1) + "円 に修正する(差 " + Fmt.s0(d) + "円)。東レへは報告しない"
                : "①の誤り " + Fmt.s0(d) + "円 として東レへ報告する(報告計上額に含む)";
        String sideLabel = man.side() == 1 ? "①東レ" : "②" + profile.name2();
        String note = "手動判定: 今回(" + ymLabel + ")に限り" + sideLabel + "を正とする → " + act;
        if (man.reason() != null && !man.reason().isBlank()) {
            note += " / 理由: " + man.reason();
        }
        return note;
    }

    private void applyPriorShortfalls(Set<String> k1, Set<String> k2) {
        JudgmentCsv.PriorLoad load =
                JudgmentCsv.loadPrior(JudgmentCsv.priorFile(paths), profile.label(), targetYm);
        warnings.addAll(load.warnings());
        for (JudgmentCsv.PriorRow row : load.rows()) {
            String k = row.keiyaku();
            if (prevAdjust.contains(k)) {
                warnings.add(JudgmentCsv.PRIOR_FILE + ": 契約NO " + k
                        + " は既に前月調整(①に調整行あり)のため重複を避けスキップしました");
                continue;
            }
            if (k1.contains(k) && k2.contains(k)) {
                warnings.add(JudgmentCsv.PRIOR_FILE + ": 契約NO " + k
                        + " は当月①②の両方に存在します。当月の不一致として扱うため前月過不足はスキップしました");
                continue;
            }
            if (k1.contains(k)) {
                warnings.add(JudgmentCsv.PRIOR_FILE + ": 契約NO " + k
                        + " は当月①に存在します(前月調整判定を確認)。前月過不足はスキップしました");
                continue;
            }
            priorShortfall.add(row);
            double gap = row.amount2() - row.amount1();
            String note = "当月①CSVに未反映"
                    + (gap > 0.5
                            ? "（プラス差 " + Fmt.s0(gap) + "円は①データに含まれない。マイナス側は取消行で前月調整に出る）"
                            : "（差額 " + Fmt.s0(row.diff()) + "円）")
                    + " / 前月当時① " + Fmt.n0(row.amount1()) + "円・② " + Fmt.n0(row.amount2())
                    + "円 / 報告計上(①-②) " + Fmt.s0(row.diff())
                    + (row.reason() == null || row.reason().isBlank() ? "" : " / " + row.reason());
            recordsA.add(new RecordA(
                    k,
                    row.irai() == null || row.irai().isBlank() ? "" : row.irai(),
                    null,
                    gap,
                    row.diff(),
                    row.diff(),
                    Judge.PREV_GAP,
                    note));
        }
    }

    private void runCheckC() {
        if (profile.id() != FactoryId.KONAN) {
            return;
        }
        Path monthlyDir = paths.monthlyDir(profile.id());
        Path monthlyFile = CheckVerifyC.findMonthlyFile(monthlyDir, targetYm);
        if (monthlyFile == null) {
            warnings.add("月次処理ファイルが見つからないため検証Cをスキップしました");
            checkC = CheckCResult.skipped("月次処理ファイルがありません");
            return;
        }
        Map<String, Double> monthly = CheckVerifyC.readMonthlyFile(monthlyFile);
        double total2Irai = source2.totalIrai();
        double total3 = aladdinData.total();
        double total1 = toray.total();
        double adjTotal = sumOf(prevAdjust, toray.byKeiyaku());
        double nextTotal = sumOf(nextMissrec, toray.byKeiyaku());
        double manual1Total = recordsA.stream()
                .filter(r -> Judge.MANUAL_1.equals(r.judge()) && r.diff() != null)
                .mapToDouble(RecordA::diff)
                .sum();
        double diffTotal = mismatchA.stream().mapToDouble(MismatchA::diff).sum()
                + recordsA.stream()
                        .filter(r -> Judge.MANUAL_2.equals(r.judge()) && r.diff() != null)
                        .mapToDouble(RecordA::diff)
                        .sum();
        double only1Total = sumOf(only1, toray.byKeiyaku());
        double only2Total = sumOf(only2, source2.byKeiyaku());
        double explained = adjTotal + nextTotal + diffTotal + manual1Total + only1Total - only2Total;
        checkC = CheckVerifyC.evaluate(monthlyFile, monthly, total2Irai, total3, total1, explained, tol);
        for (CheckCResult.Row row : checkC.rows()) {
            if (CheckCResult.MISMATCH.equals(row.judge())
                    || CheckCResult.NEED_CHECK.equals(row.judge())
                    || CheckCResult.UNREADABLE.equals(row.judge())
                    || CheckCResult.SHEET_INTERNAL.equals(row.judge())) {
                warnings.add("検証C " + row.item() + " " + row.judge()
                        + (row.note() == null ? "" : " (" + row.note() + ")"));
            }
        }
    }

    // ---------------------------------------------------------------- 検証B

    private void verifyB() {
        if (skipVerifyB) {
            return;
        }
        Map<String, Double> m2 = source2.byIrai();
        Map<String, Double> m3 = aladdin;
        Set<String> i2 = new TreeSet<>(m2.keySet());
        Set<String> i3 = new TreeSet<>(m3.keySet());

        List<String> common = new ArrayList<>();
        for (String k : i2) {
            if (i3.contains(k)) {
                common.add(k);
            }
        }
        commonB = common.size();

        buildBranchMerge(i2, i3, m2, m3);

        for (String k : common) {
            double v2 = m2.get(k);
            double v3 = m3.get(k);
            double d = v2 - v3;
            if (isZero(v2) && isZero(v3)) {
                bumpZero("B一致");
                continue;
            }
            if (Math.abs(d) <= tol) {
                recordsB.add(new RecordB(k, v2, v3, d, Judge.MATCH, ""));
            } else if (branchOk.containsKey(k)) {
                branchRows.add(new BranchRow(k, v2, v3));
                recordsB.add(new RecordB(k, v2, v3, d, Judge.BRANCH_MERGED, branchOk.get(k)));
            } else if (tryResolve(k)) {
                resolvedRows.add(new ResolvedRow(Judge.MISMATCH, k, v2, v3));
                recordsB.add(new RecordB(k, v2, v3, d, Judge.RESOLVED, resolvedB.get(k)));
            } else {
                mismatchB.add(new MismatchA(k, v2, v3, d));
                recordsB.add(new RecordB(k, v2, v3, d, Judge.MISMATCH, noteB(k)));
            }
        }

        List<String> only2Keys = new ArrayList<>();
        for (String k : i2) {
            if (i3.contains(k)) {
                continue;
            }
            if (branchOk.containsKey(k)) {
                branchRows.add(new BranchRow(k, m2.get(k), 0.0));
                recordsB.add(new RecordB(k, m2.get(k), null, null, Judge.BRANCH_MERGED, branchOk.get(k)));
            } else if (Math.abs(m2.get(k)) > tol && tryResolve(k)) {
                resolvedRows.add(new ResolvedRow(Judge.ONLY_2, k, m2.get(k), 0.0));
                recordsB.add(new RecordB(k, m2.get(k), null, null, Judge.RESOLVED, resolvedB.get(k)));
            } else {
                only2Keys.add(k);
            }
        }
        only2b = reportOnlyB(only2Keys, m2);

        List<String> only3Keys = new ArrayList<>();
        for (String k : i3) {
            if (i2.contains(k)) {
                continue;
            }
            if (Math.abs(m3.get(k)) <= tol) {
                bumpZero("B③のみ");
            } else if (branchChild.containsKey(k)) {
                branchRows.add(new BranchRow(k, 0.0, m3.get(k)));
                recordsB.add(new RecordB(k, null, m3.get(k), null, Judge.BRANCH_MERGED,
                        "親依頼NO " + branchChild.get(k) + " に合算して②と一致 / " + branchOk.get(branchChild.get(k))));
            } else if (tryResolve(k)) {
                resolvedRows.add(new ResolvedRow(Judge.ONLY_3, k, 0.0, m3.get(k)));
                recordsB.add(new RecordB(k, null, m3.get(k), null, Judge.RESOLVED, resolvedB.get(k)));
            } else {
                only3Keys.add(k);
                recordsB.add(new RecordB(k, null, m3.get(k), null, Judge.ONLY_3, note3(k, i2, m2, m3)));
            }
        }
        only3b = List.copyOf(only3Keys);

        Set<String> commonSet = new LinkedHashSet<>(common);
        branchCommonCount = (int) branchRows.stream().filter(r -> commonSet.contains(r.irai())).count();
    }

    /**
     * 枝番統合。③が枝番（例 C8-3-1）に分割計上し②は親依頼NO（C8-3）に一本で計上しているケースで、
     * ③の親+枝番合計が②の親と一致すれば「枝番統合一致」とする。
     */
    private void buildBranchMerge(Set<String> i2, Set<String> i3,
                                  Map<String, Double> m2, Map<String, Double> m3) {
        Map<String, List<String>> children = new LinkedHashMap<>();
        for (String k : i3) {
            if (i2.contains(k)) {
                continue;
            }
            String parent = BRANCH_SUFFIX_PATTERN.matcher(k).replaceAll("");
            if (!parent.equals(k) && i2.contains(parent) && Math.abs(m3.get(k)) > tol) {
                children.computeIfAbsent(parent, x -> new ArrayList<>()).add(k);
            }
        }
        for (Map.Entry<String, List<String>> e : children.entrySet()) {
            String parent = e.getKey();
            List<String> kids = e.getValue();
            if (i3.contains(parent) && Math.abs(m2.get(parent) - m3.get(parent)) <= tol) {
                continue;   // 親だけで一致済み → 枝番は本当の③のみ
            }
            double merged = m3.getOrDefault(parent, 0.0);
            for (String c : kids) {
                merged += m3.get(c);
            }
            if (Math.abs(m2.get(parent) - merged) <= tol) {
                List<String> parts = new ArrayList<>();
                parts.add(parent + " " + Fmt.n0(m3.getOrDefault(parent, 0.0)));
                for (String c : kids) {
                    parts.add(c + " " + Fmt.n0(m3.get(c)));
                }
                branchOk.put(parent, "枝番統合一致: ③ " + String.join(" + ", parts) + " = " + Fmt.n0(merged)
                        + "円 = ② " + Fmt.n0(m2.get(parent)) + "円 → ③が枝番に分割計上");
                for (String c : kids) {
                    branchChild.put(c, parent);
                }
            }
        }
    }

    private List<String> reportOnlyB(List<String> keys, Map<String, Double> amounts) {
        List<String> nonzero = new ArrayList<>();
        int zeros = 0;
        for (String k : keys) {
            if (Math.abs(amounts.get(k)) > tol) {
                nonzero.add(k);
            } else {
                zeros++;
            }
        }
        if (zeros > 0) {
            bumpZero("B" + Judge.ONLY_2, zeros);
        }
        for (String k : nonzero) {
            recordsB.add(new RecordB(k, amounts.get(k), null, null, Judge.ONLY_2,
                    FileDiscovery.joinNonEmpty(" / ", "③月次実績に計上なし(他月にも累計一致なし)", noteB(k))));
        }
        return List.copyOf(nonzero);
    }

    /** ②の全月度累計と③の全照会月累計が一致すれば「月ずれ解消」とする。 */
    private boolean tryResolve(String key) {
        if (resolvedB.containsKey(key)) {
            return true;
        }
        if (aladdinByYm.size() < 2) {
            return false;
        }
        List<Map.Entry<YearMonthKey, Double>> parts2 = cumParts(iraiByYm, key);
        List<Map.Entry<YearMonthKey, Double>> parts3 = cumParts(aladdinByYm, key);
        if (parts2.isEmpty() || parts3.isEmpty()) {
            return false;
        }
        double cum2 = sum(parts2);
        double cum3 = sum(parts3);
        if (Math.abs(cum2 - cum3) > tol) {
            return false;
        }
        resolvedB.put(key, "累計一致: ②累計 " + Fmt.n0(cum2) + "円(" + labels(parts2, "度")
                + ") = ③累計 " + Fmt.n0(cum3) + "円(" + labels(parts3, "") + ") → 計上月ずれと判明");
        return true;
    }

    /** 非解消項目の備考用の累計内訳。 */
    private String cumSummary(String key) {
        if (aladdinByYm.size() < 2) {
            return "";
        }
        List<Map.Entry<YearMonthKey, Double>> parts2 = cumParts(iraiByYm, key);
        List<Map.Entry<YearMonthKey, Double>> parts3 = cumParts(aladdinByYm, key);
        if (parts3.isEmpty()) {
            return "";
        }
        double cum2 = sum(parts2);
        double cum3 = sum(parts3);
        return "②累計 " + Fmt.n0(cum2) + "円(" + labels(parts2, "度") + ") / ③累計 " + Fmt.n0(cum3)
                + "円(" + labels(parts3, "") + ") / 累計差 " + Fmt.s0(cum2 - cum3) + "円";
    }

    private String noteB(String key) {
        return FileDiscovery.joinNonEmpty(" / ", monthNote(key), cumSummary(key));
    }

    /** ③のみ依頼NOの注記（枝番の親依頼NO確認・月ずれ）。 */
    private String note3(String key, Set<String> i2, Map<String, Double> m2, Map<String, Double> m3) {
        String parent = BRANCH_SUFFIX_PATTERN.matcher(key).replaceAll("");
        String note;
        if (!parent.equals(key) && i2.contains(parent)) {
            double merged = m3.getOrDefault(parent, 0.0) + m3.get(key);
            note = "枝番? 親依頼NO " + parent + " は②に存在。③親+枝番計 " + Fmt.n0(merged)
                    + "円 / ②親 " + Fmt.n0(m2.get(parent)) + "円 で突合してみること";
        } else {
            note = "②" + profile.name2() + "に依頼NOなし";
        }
        return FileDiscovery.joinNonEmpty(" / ", note, monthNote(key));
    }

    /** 依頼NOの月番号が対象月と異なる場合の月ずれ注記。 */
    private String monthNote(String key) {
        Matcher m = IRAI_MONTH_PATTERN.matcher(key);
        if (targetYm != null && m.find() && Integer.parseInt(m.group(1)) != targetYm.month()) {
            return m.group(1) + "月系依頼(計上月ずれの可能性)";
        }
        return "";
    }

    private static List<Map.Entry<YearMonthKey, Double>> cumParts(
            NavigableMap<YearMonthKey, Map<String, Double>> byYm, String key) {
        List<Map.Entry<YearMonthKey, Double>> parts = new ArrayList<>();
        for (Map.Entry<YearMonthKey, Map<String, Double>> e : byYm.entrySet()) {
            Double v = e.getValue().get(key);
            if (v != null) {
                parts.add(Map.entry(e.getKey(), v));
            }
        }
        return parts;
    }

    private static double sum(List<Map.Entry<YearMonthKey, Double>> parts) {
        double s = 0.0;
        for (Map.Entry<YearMonthKey, Double> e : parts) {
            s += e.getValue();
        }
        return s;
    }

    private static String labels(List<Map.Entry<YearMonthKey, Double>> parts, String suffix) {
        List<String> out = new ArrayList<>(parts.size());
        for (Map.Entry<YearMonthKey, Double> e : parts) {
            out.add(e.getKey().monthLabel() + suffix);
        }
        return String.join("+", out);
    }

    // ---------------------------------------------------------------- 集計

    private VerifyResult buildResult() {
        Map<String, Double> t1 = toray.byKeiyaku();
        Map<String, Double> t2 = source2.byKeiyaku();

        double total1 = toray.total();
        double total2 = source2.totalKeiyaku();
        double adjTotal = sumOf(prevAdjust, t1);
        double nextTotal = sumOf(nextMissrec, t1);
        double mismatchTotal = mismatchA.stream().mapToDouble(MismatchA::diff).sum();
        double manual2Total = recordsA.stream()
                .filter(r -> Judge.MANUAL_2.equals(r.judge()) && r.diff() != null)
                .mapToDouble(RecordA::diff)
                .sum();
        double manual1Total = recordsA.stream()
                .filter(r -> Judge.MANUAL_1.equals(r.judge()) && r.diff() != null)
                .mapToDouble(RecordA::diff)
                .sum();
        int manual1Count = (int) recordsA.stream().filter(r -> Judge.MANUAL_1.equals(r.judge())).count();
        int manual2Count = (int) recordsA.stream().filter(r -> Judge.MANUAL_2.equals(r.judge())).count();
        double priorTotal = priorShortfall.stream().mapToDouble(JudgmentCsv.PriorRow::diff).sum();
        double diffTotal = mismatchTotal + manual2Total;
        double only1Total = sumOf(only1, t1);
        double only2Total = sumOf(only2, t2);
        double residual = reportResidual(
                total1, total2, adjTotal, nextTotal, diffTotal, manual1Total, only1Total, only2Total);
        double mailTougetsu = diffTotal + nextTotal + only1Total - only2Total;
        double tougetsu = mailTougetsu + priorTotal;

        int issues = mismatchA.size() + mismatchB.size() + nextMissrec.size()
                + only1.size() + only2.size() + only2b.size() + only3b.size() + priorShortfall.size();

        recordsA.sort(Comparator
                .comparingInt((RecordA r) -> Judge.rank(r.judge()))
                .thenComparingDouble(r -> {
                    if (Judge.MISMATCH.equals(r.judge())) {
                        return r.diff() == null ? 0.0 : -Math.abs(r.diff());
                    }
                    return -r.scale();
                })
                .thenComparing(RecordA::keiyaku));
        recordsB.sort(Comparator
                .comparingInt((RecordB r) -> Judge.rank(r.judge()))
                .thenComparingDouble(r -> {
                    if (Judge.MISMATCH.equals(r.judge())) {
                        return r.diff() == null ? 0.0 : -Math.abs(r.diff());
                    }
                    return -r.scale();
                })
                .thenComparing(RecordB::irai));

        String ymLabel = targetYm != null ? targetYm.gatsudoLabel() : "(不明)";
        String nextYmLabel = targetYm != null ? targetYm.plusMonths(1).gatsudoLabel() : "翌月";

        Map<String, Object> info = new LinkedHashMap<>();
        info.put("工場", profile.label());
        info.put("他工場", profile.otherLabel());
        info.put("②名称", profile.name2());
        info.put("②説明", profile.desc2());
        info.put("②金額列", profile.amt2());
        info.put("③得意先", profile.customer3());
        info.put("対象月ラベル", ymLabel);
        info.put("翌月ラベル", nextYmLabel);
        info.put("許容差", tol);
        info.put("入庫場所", Norm.norm(profile.basho()));
        info.put("対象外入庫場所", excludedBasho());
        info.put("実行日時", LocalDateTime.now().format(TIMESTAMP));
        info.put("実行者", executor());
        info.put("①ファイル", torayFile.getFileName().toString());
        info.put("②ファイル", FileDiscovery.displayName(dir2, source2File));
        info.put("③ファイル", aladdinFile == null ? "(なし)" : aladdinFile.getFileName().toString());
        info.put("前月ファイル", FileDiscovery.displayNames(dir2, fileSet2.prevs()));
        info.put("翌月ファイル", FileDiscovery.displayNames(dir2, fileSet2.nexts()));
        info.put("③他月照会", aladdinOtherFiles.isEmpty() ? "なし" : String.join(", ", aladdinOtherFiles));
        info.put("対象年月", aladdinData == null ? "" : aladdinData.taishoText());
        info.put("①取消行", minusLabel());
        info.put("0円除外", zeroLabel());
        info.put("①総額", Math.round(total1));
        info.put("②総額", Math.round(source2.totalIrai()));
        info.put("③総額", Math.round(aladdinData == null ? 0 : aladdinData.total()));
        info.put("A共通", commonA);
        info.put("A不一致", mismatchA.size());
        info.put("A①のみ", only1.size());
        info.put("A②のみ", only2.size());
        info.put("A前月調整件数", prevAdjust.size());
        info.put("A前月調整額", adjTotal);
        info.put("A翌月記載件数", nextMissrec.size());
        info.put("A翌月記載額", nextTotal);
        info.put("A前月過不足件数", priorShortfall.size());
        info.put("A前月過不足額", priorTotal);
        info.put("A手動判定件数", manual1Count + manual2Count);
        info.put("A手動①正件数", manual1Count);
        info.put("A手動①正額", manual1Total);
        info.put("A手動②正件数", manual2Count);
        info.put("A手動②正額", manual2Total);
        info.put("①パス", torayFile.toAbsolutePath().toString());
        info.put("②パス", source2File.toAbsolutePath().toString());
        info.put("③パス", aladdinFile == null ? "" : aladdinFile.toAbsolutePath().toString());
        info.put("報告②実売上", Math.round(total2));
        info.put("報告当月差異", Math.round(diffTotal));
        info.put("報告①のみ額", only1Total);
        info.put("報告②のみ額", only2Total);
        info.put("報告する過不足", Math.round(tougetsu));
        info.put("報告残差", Math.round(residual));
        info.put("検証C要確認", checkC == null ? 0 : checkC.needCheckCount());
        info.put("B共通", commonB);
        info.put("B不一致", mismatchB.size());
        info.put("B月ずれ解消", resolvedRows.size());
        info.put("B月ずれ解消共通",
                (int) resolvedRows.stream().filter(r -> Judge.MISMATCH.equals(r.origin())).count());
        info.put("B枝番統合", branchRows.size());
        info.put("B枝番統合共通", branchCommonCount);
        info.put("B②のみ", only2b.size());
        info.put("B③のみ", only3b.size());
        info.put("検証Bスキップ", skipVerifyB);
        if (skipVerifyB) {
            String reason = warnings.stream()
                    .filter(w -> w.contains("【検証Bスキップ】"))
                    .findFirst()
                    .orElse("");
            info.put("検証Bスキップ理由", reason);
        }
        info.put("要確認", issues);
        info.put("警告", List.copyOf(warnings));

        MailSnapshot mail = new MailSnapshot(
                profile.id().code(),
                profile.label(),
                ymLabel,
                nextYmLabel,
                Math.round(total1),
                Math.round(total2),
                Fmt.round0(adjTotal - priorTotal),
                Fmt.round0(mailTougetsu),
                mismatchA.size() + manual2Count + nextMissrec.size() + only1.size() + only2.size(),
                mismatchA.size(),
                Fmt.round0(diffTotal),
                nextMissrec.size(),
                Fmt.round0(nextTotal),
                only1.size(),
                Fmt.round0(only1Total),
                only2.size(),
                Fmt.round0(only2Total),
                excludedBasho(),
                LocalDateTime.now().format(TIMESTAMP));

        return new VerifyResult(profile, targetYm, List.copyOf(recordsA), List.copyOf(recordsB),
                Collections.unmodifiableMap(info), List.copyOf(warnings), mail, checkD, checkC);
    }

    // ---------------------------------------------------------------- 補助

    private boolean isZero(double v) {
        return Math.abs(v) <= tol;
    }

    private void bumpZero(String key) {
        bumpZero(key, 1);
    }

    private void bumpZero(String key, int count) {
        zeroExcluded.merge(key, count, Integer::sum);
    }

    private String zeroLabel() {
        List<String> a = new ArrayList<>();
        List<String> b = new ArrayList<>();
        for (Map.Entry<String, Integer> e : zeroExcluded.entrySet()) {
            if (e.getValue() == 0) {
                continue;
            }
            String text = e.getKey().substring(1) + " " + e.getValue() + "件";
            if (e.getKey().startsWith("A")) {
                a.add(text);
            } else {
                b.add(text);
            }
        }
        String label = FileDiscovery.joinNonEmpty(" / ",
                a.isEmpty() ? "" : "検証A: " + String.join("・", a),
                b.isEmpty() ? "" : "検証B: " + String.join("・", b));
        return label.isEmpty() ? "なし" : label;
    }

    private String excludedBasho() {
        String target = Norm.norm(profile.basho());
        List<String> parts = new ArrayList<>();
        for (Map.Entry<String, Double> e : new TreeMap<>(toray.byBasho()).entrySet()) {
            if (!e.getKey().equals(target)) {
                parts.add(e.getKey() + " " + Fmt.n0(e.getValue()) + "円");
            }
        }
        return String.join(" / ", parts);
    }

    private String minusLabel() {
        int n = toray.minusRowCount();
        if (n == 0) {
            return "なし";
        }
        return n + "件 " + Fmt.n0(toray.minusTotal()) + "円 を符号反転して集計 (契約NO: "
                + String.join(", ", sorted(toray.minusRows().keySet())) + ")";
    }

    private String minusNote(String key) {
        List<TorayCsvData.MinusRow> rows = toray.minusRows().get(key);
        if (rows == null || rows.isEmpty()) {
            return "";
        }
        List<String> detail = new ArrayList<>(rows.size());
        for (TorayCsvData.MinusRow r : rows) {
            detail.add(r.rowNo() + "行目 " + Fmt.n0(r.amount()) + "円");
        }
        return "①に取消行(H列「-」)" + rows.size() + "件あり: " + String.join(", ", detail)
                + " → 純額 " + Fmt.n0(toray.byKeiyaku().getOrDefault(key, 0.0)) + "円";
    }

    private String nyukoNote(String key) {
        List<String> parts = new ArrayList<>();
        Set<String> dates = toray.nyukoDates().get(key);
        if (dates != null && !dates.isEmpty()) {
            parts.add("①入庫日 " + String.join(", ", sorted(dates)));
        }
        String minus = minusNote(key);
        if (!minus.isEmpty()) {
            parts.add(minus);
        }
        String hint = ctHint.get(key);
        if (hint != null) {
            parts.add(hint);
        }
        return String.join(" / ", parts);
    }

    private static Hit findHit(List<MonthData> months, String keiyaku) {
        for (MonthData m : months) {
            Double v = m.byKeiyaku().get(keiyaku);
            if (v != null) {
                return new Hit(m.label(), v, m.keiyakuToIrai().getOrDefault(keiyaku, Set.of()));
            }
        }
        return null;
    }

    private static double sumOf(List<String> keys, Map<String, Double> amounts) {
        double s = 0.0;
        for (String k : keys) {
            s += amounts.getOrDefault(k, 0.0);
        }
        return s;
    }

    private static List<String> sorted(Set<String> values) {
        List<String> list = new ArrayList<>(values);
        Collections.sort(list);
        return list;
    }

    private static String executor() {
        String user = System.getProperty("user.name");
        if (user == null || user.isBlank()) {
            user = System.getenv("USERNAME");
        }
        return user == null || user.isBlank() ? "(不明)" : user;
    }
}
