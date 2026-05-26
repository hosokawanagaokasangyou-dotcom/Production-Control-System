package jp.co.pm.ai.desktop.dispatch.rules.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonIgnoreProperties(ignoreUnknown = true)
public final class DispatchRuleEdge {

    @JsonProperty("id")
    public String id = "";

    @JsonProperty("from")
    public String from = "";

    @JsonProperty("to")
    public String to = "";

    @JsonProperty("fromPort")
    public String fromPort = "out";

    @JsonProperty("toPort")
    public String toPort = "in";
}
