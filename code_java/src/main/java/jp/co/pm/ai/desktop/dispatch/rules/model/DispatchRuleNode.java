package jp.co.pm.ai.desktop.dispatch.rules.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonIgnoreProperties(ignoreUnknown = true)
public final class DispatchRuleNode {

    @JsonProperty("id")
    public String id = "";

    @JsonProperty("type")
    public String type = "";

    @JsonProperty("label")
    public String label = "";

    @JsonProperty("x")
    public double x;

    @JsonProperty("y")
    public double y;

    @JsonProperty("params")
    public java.util.LinkedHashMap<String, Object> params = new java.util.LinkedHashMap<>();
}
