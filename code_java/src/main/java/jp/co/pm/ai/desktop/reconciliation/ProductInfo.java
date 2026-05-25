package jp.co.pm.ai.desktop.reconciliation;

public class ProductInfo {
    private final String shohinCode;
    private final String seihinCode;
    private final String shohinName1;
    private final String shohinName2;
    private final String unitName;
    private final String quantityPerCase;
    private final String selfKakoKbn;
    private final String foamName;
    private final String foamPartNo;
    private final String foamWidth;
    private final String foamLength;
    private final String foamColor;
    private final String foamThickness;
    private final String kakoNaiyo;

    public ProductInfo(String shohinCode, String seihinCode, String shohinName1, String shohinName2,
                       String unitName, String quantityPerCase, String selfKakoKbn, String foamName,
                       String foamPartNo, String foamWidth, String foamLength, String foamColor, String foamThickness,
                       String kakoNaiyo) {
        this.shohinCode = shohinCode != null ? shohinCode.trim() : "";
        this.seihinCode = seihinCode != null ? seihinCode.trim() : "";
        this.shohinName1 = shohinName1 != null ? shohinName1.trim() : "";
        this.shohinName2 = shohinName2 != null ? shohinName2.trim() : "";
        this.unitName = unitName != null ? unitName.trim() : "";
        this.quantityPerCase = quantityPerCase != null ? quantityPerCase.trim() : "";
        this.selfKakoKbn = selfKakoKbn != null ? selfKakoKbn.trim() : "";
        this.foamName = foamName != null ? foamName.trim() : "";
        this.foamPartNo = foamPartNo != null ? foamPartNo.trim() : "";
        this.foamWidth = foamWidth != null ? foamWidth.trim() : "";
        this.foamLength = foamLength != null ? foamLength.trim() : "";
        this.foamColor = foamColor != null ? foamColor.trim() : "";
        this.foamThickness = foamThickness != null ? foamThickness.trim() : "";
        this.kakoNaiyo = kakoNaiyo != null ? kakoNaiyo.trim() : "";
    }

    public String getShohinCode() { return shohinCode; }
    public String getSeihinCode() { return seihinCode; }
    public String getShohinName1() { return shohinName1; }
    public String getShohinName2() { return shohinName2; }
    public String getUnitName() { return unitName; }
    public String getQuantityPerCase() { return quantityPerCase; }
    public String getSelfKakoKbn() { return selfKakoKbn; }
    public String getFoamName() { return foamName; }
    public String getFoamPartNo() { return foamPartNo; }
    public String getFoamWidth() { return foamWidth; }
    public String getFoamLength() { return foamLength; }
    public String getFoamColor() { return foamColor; }
    public String getFoamThickness() { return foamThickness; }
    public String getKakoNaiyo() { return kakoNaiyo; }

    @Override
    public String toString() {
        return shohinCode + " - " + shohinName1 + " (" + foamPartNo + ")";
    }
}
