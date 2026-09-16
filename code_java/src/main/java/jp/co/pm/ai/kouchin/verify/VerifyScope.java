package jp.co.pm.ai.kouchin.verify;

/**
 * 後加工工賃の検証対象。TPI（東レペフ加工品）と自社加工は東レ突合の対象外。
 */
public final class VerifyScope {

    /** ③アラジンの TPI 得意先コード（東ﾚﾍﾟﾌ加工品）。 */
    public static final String TPI_CUSTOMER = "049052";

    private VerifyScope() {}

    /**
     * @param irai     依頼NO（未正規化可）
     * @param customer 得意先コード（未正規化可。無ければ空）
     */
    public static boolean outOfScope(String irai, String customer) {
        String cust = Norm.norm(customer);
        if (TPI_CUSTOMER.equals(cust)) {
            return true;
        }
        String n = Norm.norm(irai);
        if (n.isEmpty()) {
            return false;
        }
        return n.startsWith("TPI") || n.charAt(0) == '2';
    }

    /** 依頼NOだけ見る（②明細用）。 */
    public static boolean outOfScopeIrai(String irai) {
        return outOfScope(irai, "");
    }
}
