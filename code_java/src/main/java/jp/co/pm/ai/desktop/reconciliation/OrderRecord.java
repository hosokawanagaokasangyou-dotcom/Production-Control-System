package jp.co.pm.ai.desktop.reconciliation;

import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import java.util.Map;

public class OrderRecord {
    private final StringProperty reqNo;
    private final StringProperty status;
    private final StringProperty user;
    private final StringProperty product;
    private final StringProperty discrepancy;
    
    private final Map<String, String> rawValues;
    private final Map<String, String> dbValues;

    public OrderRecord(String reqNo, String status, String user, String product, String discrepancy, 
                       Map<String, String> rawValues, Map<String, String> dbValues) {
        this.reqNo = new SimpleStringProperty(reqNo != null ? reqNo : "");
        this.status = new SimpleStringProperty(status != null ? status : "");
        this.user = new SimpleStringProperty(user != null ? user : "");
        this.product = new SimpleStringProperty(product != null ? product : "");
        this.discrepancy = new SimpleStringProperty(discrepancy != null ? discrepancy : "");
        this.rawValues = rawValues;
        this.dbValues = dbValues;
    }

    public String getReqNo() { return reqNo.get(); }
    public void setReqNo(String value) { reqNo.set(value); }
    public StringProperty reqNoProperty() { return reqNo; }

    public String getStatus() { return status.get(); }
    public void setStatus(String value) { status.set(value); }
    public StringProperty statusProperty() { return status; }

    public String getUser() { return user.get(); }
    public void setUser(String value) { user.set(value); }
    public StringProperty userProperty() { return user; }

    public String getProduct() { return product.get(); }
    public void setProduct(String value) { product.set(value); }
    public StringProperty productProperty() { return product; }

    public String getDiscrepancy() { return discrepancy.get(); }
    public void setDiscrepancy(String value) { discrepancy.set(value); }
    public StringProperty discrepancyProperty() { return discrepancy; }

    public Map<String, String> getRawValues() { return rawValues; }
    public Map<String, String> getDbValues() { return dbValues; }
}
