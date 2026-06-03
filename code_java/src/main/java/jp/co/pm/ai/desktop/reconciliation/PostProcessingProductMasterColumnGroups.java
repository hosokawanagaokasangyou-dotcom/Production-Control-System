package jp.co.pm.ai.desktop.reconciliation;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * {@code 後加工商品マスタ.xlsx} の列を Aladdin 画面タブ相当のグループに分類する。
 * 見出し順の正本は参照マスタファイルの1行目。
 */
public final class PostProcessingProductMasterColumnGroups {

    public static final String TAB_BASIC = "基本・発泡体仕様";
    public static final String TAB_KUBUN = "区分情報";
    public static final String TAB_OTHER = "その他情報";
    public static final String TAB_FOAM = "発泡体";
    public static final String TAB_PROCESS = "工程情報";
    public static final String TAB_GENPAN = "原反情報";
    public static final String TAB_EXTENSION = "その他（拡張）";

    private static final List<String> BASIC =
            List.of(
                    "商品コード",
                    "製品コード",
                    "商品名1",
                    "商品名2",
                    "単位名",
                    "入数",
                    "自社後加工区分",
                    "発泡体品名",
                    "発泡体品番",
                    "発泡体タイプ",
                    "発泡体幅",
                    "発泡体長さ",
                    "発泡体梱等",
                    "発泡体色",
                    "発泡体区分",
                    "発泡体厚み",
                    "単重",
                    "商品特記事項");

    private static final List<String> KUBUN =
            List.of(
                    "商品分類1コード",
                    "商品分類2コード",
                    "商品分類3コード",
                    "商品分類4コード",
                    "商品分類5コード",
                    "商品分類6コード",
                    "商品分類7コード",
                    "単価分類コード",
                    "適正在庫数量",
                    "名称入力区分",
                    "在庫管理区分",
                    "税率区分コード",
                    "品区分",
                    "ロット管理区分",
                    "削除区分",
                    "仕入先コード",
                    "JANコード",
                    "画像ファイル名",
                    "展開区分",
                    "原価洗替区分",
                    "原価単価取得区分",
                    "売上時原価引当区分",
                    "直送原価取得区分",
                    "手配区分",
                    "AEC連携対象フラグ",
                    "原価掛率",
                    "登録画面区分");

    private static final List<String> OTHER =
            List.of("発注ロット", "リードタイム", "名カナ", "備考1", "備考2", "備考3", "メモ");

    private static final List<String> FOAM =
            List.of(
                    "トリミング",
                    "加工回数",
                    "EC面指定コード",
                    "融着",
                    "ユーザ",
                    "在庫場所",
                    "UL規格",
                    "長さ換算区分",
                    "梱包仕様コード",
                    "梱包仕様名");

    private static final List<String> PROCESS =
            List.of(
                    "工程コード1",
                    "加工単価1",
                    "工程コード2",
                    "加工単価2",
                    "工程コード3",
                    "加工単価3",
                    "工程コード4",
                    "加工単価4",
                    "工程コード5",
                    "加工単価5",
                    "工程コード6",
                    "加工単価6",
                    "工程コード7",
                    "加工単価7",
                    "工程コード8",
                    "加工単価8",
                    "材料単価",
                    "加工単価",
                    "機械コード1",
                    "機械コード2",
                    "機械コード3",
                    "機械コード4",
                    "機械コード5",
                    "機械コード6",
                    "機械コード7",
                    "機械コード8",
                    "加工単価区分",
                    "合算加工単価",
                    "加工内容コード1",
                    "加工内容コード2",
                    "加工内容コード3",
                    "加工内容コード4",
                    "加工内容コード5",
                    "加工内容コード6",
                    "加工内容コード7",
                    "加工内容コード8");

    private static final List<String> GENPAN =
            List.of(
                    "原反商品コード1",
                    "原反商品コード2",
                    "原反商品コード3",
                    "原反商品コード4");

    private static final List<String> EXTENSION =
            List.of(
                    "プラマキシン品番",
                    "内径1",
                    "内径2",
                    "肉厚",
                    "プラマキシン長さ",
                    "胴長さ",
                    "端部カット仕様区分",
                    "端部加工寸法",
                    "溝幅",
                    "端部カット長",
                    "外径",
                    "表面積",
                    "指数",
                    "表面積指数",
                    "コア規格材質コード",
                    "コア規格材質名1",
                    "コア規格材質名2",
                    "コア規格材質商品名称入力区分",
                    "コア材質",
                    "粘着シートコード",
                    "粘着シート名1",
                    "粘着シート名2",
                    "粘着シート商品名称入力区分",
                    "接着幅",
                    "接着使用量",
                    "接着仕様",
                    "クッションコード",
                    "クッション名1",
                    "クッション名2",
                    "クッション商品名称入力区分",
                    "クッション幅",
                    "クッション使用量",
                    "梱包資材コード",
                    "梱包資材名1",
                    "梱包資材名2",
                    "梱包資材商品名称入力区分",
                    "梱包資材使用量",
                    "使用ケースコード1",
                    "使用ケース名11",
                    "使用ケース名12",
                    "使用ケース1商品名称入力区分",
                    "構成数1",
                    "使用ケースコード2",
                    "使用ケース名21",
                    "使用ケース名22",
                    "使用ケース2商品名称入力区分",
                    "構成数2",
                    "使用パレットコード",
                    "使用パレット名1",
                    "使用パレット名2",
                    "使用パレット商品名称入力区分",
                    "梱包入数");

    public record TabGroup(String tabTitle, List<String> columnNames) {}

    private PostProcessingProductMasterColumnGroups() {}

    public static List<TabGroup> tabGroups() {
        return List.of(
                new TabGroup(TAB_BASIC, BASIC),
                new TabGroup(TAB_KUBUN, KUBUN),
                new TabGroup(TAB_OTHER, OTHER),
                new TabGroup(TAB_FOAM, FOAM),
                new TabGroup(TAB_PROCESS, PROCESS),
                new TabGroup(TAB_GENPAN, GENPAN),
                new TabGroup(TAB_EXTENSION, EXTENSION));
    }

    /** 既知グループに含まれる列名（154列想定）。 */
    public static Set<String> knownColumnNames() {
        LinkedHashSet<String> names = new LinkedHashSet<>();
        for (TabGroup g : tabGroups()) {
            names.addAll(g.columnNames());
        }
        return names;
    }

    /**
     * 参照マスタの見出し順を維持しつつ、未知列があれば末尾に追加する。
     */
    public static List<String> alignHeadersToReference(List<String> referenceHeaders) {
        List<String> ref = referenceHeaders != null ? referenceHeaders : List.of();
        List<String> out = new ArrayList<>(ref.size() + 16);
        Set<String> known = knownColumnNames();
        for (String h : ref) {
            if (h != null && !h.isBlank()) {
                out.add(h.trim());
            }
        }
        for (String k : known) {
            if (!out.contains(k)) {
                out.add(k);
            }
        }
        return List.copyOf(out);
    }

    public static void validateHeadersMatch(List<String> reference, List<String> candidate)
            throws IllegalArgumentException {
        List<String> ref = reference != null ? reference : List.of();
        List<String> cand = candidate != null ? candidate : List.of();
        if (ref.isEmpty()) {
            throw new IllegalArgumentException("参照マスタの見出し行が空です。");
        }
        if (ref.size() != cand.size()) {
            throw new IllegalArgumentException(
                    "列数が一致しません（参照="
                            + ref.size()
                            + "、アップロード="
                            + cand.size()
                            + "）。");
        }
        for (int i = 0; i < ref.size(); i++) {
            String r = ref.get(i) != null ? ref.get(i).trim() : "";
            String c = i < cand.size() && cand.get(i) != null ? cand.get(i).trim() : "";
            if (!r.equals(c)) {
                throw new IllegalArgumentException(
                        "見出しが一致しません（列"
                                + (i + 1)
                                + " 参照=\""
                                + r
                                + "\" アップロード=\""
                                + c
                                + "\"）。");
            }
        }
    }
}
