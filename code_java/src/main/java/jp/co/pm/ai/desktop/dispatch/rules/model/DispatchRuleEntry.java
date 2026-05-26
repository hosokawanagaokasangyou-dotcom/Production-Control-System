package jp.co.pm.ai.desktop.dispatch.rules.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonIgnoreProperties(ignoreUnknown = true)
public final class DispatchRuleEntry {

    @JsonProperty("id")
    public String id = "";

    @JsonProperty("name")
    public String name = "";

    @JsonProperty("enabled")
    public boolean enabled = true;

    @JsonProperty("applyOrder")
    public int applyOrder = 100;

    @JsonProperty("executionMode")
    public String executionMode = "auto";

    @JsonProperty("legacyFallback")
    public boolean legacyFallback = true;

    @JsonProperty("graph")
    public DispatchRuleGraph graph = new DispatchRuleGraph();
}
