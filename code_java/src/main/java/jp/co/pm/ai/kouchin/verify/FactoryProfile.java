package jp.co.pm.ai.kouchin.verify;

import java.util.List;

/**
 * 工場プロファイル。Python 版 {@code _factories()} のプロファイル辞書に対応する。
 *
 * @param id          工場
 * @param label       自工場名（国分工場 / 湖南工場）
 * @param otherLabel  他工場名（メール文面の並びに使う）
 * @param basho       ①東レCSVの対象入庫場所（A010 / A010P）
 * @param customer3   ③月次実績の得意先絞り込みコード（湖南のみ 049006・国分は null）
 * @param hasCheckD   検証D（②まとめの内部整合性）の対象か
 * @param name2       ②の名称（長岡明細 / 加工賃試算）
 * @param desc2       ②の説明（シート構成）
 * @param amt2        ②の金額列名（AA / 加工賃）
 * @param sheets2     ②の対象シート名（原本コピー・読み取り対象）
 * @param keyA        検証Aの突合キー説明
 * @param keyB        検証Bの突合キー説明
 * @param take2       ②の金額の取り方の説明
 */
public record FactoryProfile(
        FactoryId id,
        String label,
        String otherLabel,
        String basho,
        String customer3,
        boolean hasCheckD,
        String name2,
        String desc2,
        String amt2,
        List<String> sheets2,
        String keyA,
        String keyB,
        String take2) {

    /** 湖南② 加工賃試算の東レ3シート。 */
    public static final List<String> SHISAN_SHEETS = List.of("東レT.V.C", "東レY.S", "東レW.E..");
    /** 国分② 長岡明細のまとめシート。 */
    public static final String MATOME_SHEET = "東レまとめ";
    /** 国分② まとめの元4シート。 */
    public static final List<String> MATOME_SRC_SHEETS = List.of("東レT", "東レV.C", "東レY", "東レW.E");

    private static final FactoryProfile KOKUBU = new FactoryProfile(
            FactoryId.KOKUBU,
            "国分工場",
            "湖南工場",
            "A010",
            null,
            true,
            "長岡明細",
            "東レまとめシート",
            "AA",
            List.of(MATOME_SHEET),
            "契約NO = ①発注No.のハイフン除去(例 191-352R→191352R) = ②「東レまとめ」C列。ハイフン無し7桁英数字",
            "依頼NO = ②A列&\"-\"&B列(例 C7-52) = ③依頼NO列。全角英数字は半角に直して比較",
            "② =「東レまとめ」AA列(同一キー複数行は合算)");

    private static final FactoryProfile KONAN = new FactoryProfile(
            FactoryId.KONAN,
            "湖南工場",
            "国分工場",
            "A010P",
            "049006",
            false,
            "加工賃試算",
            "東レT.V.C/東レY.S/東レW.E..シート",
            "加工賃",
            SHISAN_SHEETS,
            "契約NO = ①発注No.のハイフン除去 = ②加工賃試算の東レ3シート C列「契約No.」。ハイフン無し7桁英数字",
            "依頼NO = ②A列「加工依頼No.」(例 C8-1) = ③依頼NO列(得意先 049006 東ﾚ自材部の行のみ)",
            "② = 東レT.V.C/東レY.S/東レW.E..シートの「加工種類別 加工賃」合計列(V列)を3シート合算");

    public static FactoryProfile of(FactoryId id) {
        return id == FactoryId.KOKUBU ? KOKUBU : KONAN;
    }

    /** ③の得意先絞り込みがあるか。 */
    public boolean hasCustomerFilter() {
        return customer3 != null && !customer3.isEmpty();
    }
}
