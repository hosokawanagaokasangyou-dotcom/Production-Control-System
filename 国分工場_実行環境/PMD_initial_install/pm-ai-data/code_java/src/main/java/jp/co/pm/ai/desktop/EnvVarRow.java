package jp.co.pm.ai.desktop;

import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.scene.control.TableView;

/** Editable row for {@link TableView} (setting_ env sheet parity). */
public final class EnvVarRow {
    private final StringProperty name = new SimpleStringProperty("");
    private final StringProperty value = new SimpleStringProperty("");
    private final StringProperty description = new SimpleStringProperty("");

    public String getName() {
        return name.get();
    }

    public void setName(String v) {
        name.set(v != null ? v : "");
    }

    public StringProperty nameProperty() {
        return name;
    }

    public String getValue() {
        return value.get();
    }

    public void setValue(String v) {
        value.set(v != null ? v : "");
    }

    public StringProperty valueProperty() {
        return value;
    }

    public String getDescription() {
        return description.get();
    }

    public void setDescription(String v) {
        description.set(v != null ? v : "");
    }

    public StringProperty descriptionProperty() {
        return description;
    }
}
