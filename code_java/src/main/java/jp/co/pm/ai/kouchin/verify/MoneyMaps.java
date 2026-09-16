package jp.co.pm.ai.kouchin.verify;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * ②（長岡明細 / 加工賃試算）の読み取り結果。
 * Python 版の {@code read_nagaoka} / {@code read_shisan} の戻り値
 * {@code (依頼NO->金額, 契約NO->金額, 契約NO->依頼NO集合, 形式不正リスト, 警告)} に対応する。
 */
public final class MoneyMaps {

    /** 契約NO形式でないC列値（記入漏れ疑いの検出用）。 */
    public record BadKeiyaku(String place, String value, double amount) {
    }

    private final Map<String, Double> byIrai = new LinkedHashMap<>();
    private final Map<String, Double> byKeiyaku = new LinkedHashMap<>();
    private final Map<String, Set<String>> keiyakuToIrai = new LinkedHashMap<>();
    private final List<BadKeiyaku> badKeiyaku = new ArrayList<>();
    private final List<String> warnings = new ArrayList<>();

    /** 依頼NO別金額を加算する。 */
    public void addIrai(String irai, double amount) {
        byIrai.merge(irai, amount, Double::sum);
    }

    /** 契約NO別金額を加算する。 */
    public void addKeiyaku(String keiyaku, double amount) {
        byKeiyaku.merge(keiyaku, amount, Double::sum);
    }

    /** 契約NO → 依頼NO の対応を記録する。 */
    public void linkIrai(String keiyaku, String irai) {
        keiyakuToIrai.computeIfAbsent(keiyaku, k -> new LinkedHashSet<>()).add(irai);
    }

    public void addBadKeiyaku(String place, String value, double amount) {
        badKeiyaku.add(new BadKeiyaku(place, value, amount));
    }

    public void addWarning(String message) {
        warnings.add(message);
    }

    public Map<String, Double> byIrai() {
        return Collections.unmodifiableMap(byIrai);
    }

    public Map<String, Double> byKeiyaku() {
        return Collections.unmodifiableMap(byKeiyaku);
    }

    public Map<String, Set<String>> keiyakuToIrai() {
        return Collections.unmodifiableMap(keiyakuToIrai);
    }

    public List<BadKeiyaku> badKeiyaku() {
        return Collections.unmodifiableList(badKeiyaku);
    }

    public List<String> warnings() {
        return Collections.unmodifiableList(warnings);
    }

    public boolean isEmpty() {
        return byIrai.isEmpty() && byKeiyaku.isEmpty();
    }

    public double totalIrai() {
        return byIrai.values().stream().mapToDouble(Double::doubleValue).sum();
    }

    public double totalKeiyaku() {
        return byKeiyaku.values().stream().mapToDouble(Double::doubleValue).sum();
    }

    /** 依頼NOの金額（無ければ 0）。 */
    public double irai(String key) {
        return byIrai.getOrDefault(key, 0.0);
    }

    /** 契約NOの金額（無ければ 0）。 */
    public double keiyaku(String key) {
        return byKeiyaku.getOrDefault(key, 0.0);
    }

    /** 契約NOに紐づく依頼NO集合（無ければ空）。 */
    public Set<String> iraiOf(String keiyaku) {
        return keiyakuToIrai.getOrDefault(keiyaku, Set.of());
    }
}
