package jp.co.pm.ai.desktop.dispatch.rules.model;

import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonIgnoreProperties(ignoreUnknown = true)
public final class DispatchRuleGraph {

    @JsonProperty("nodes")
    public List<DispatchRuleNode> nodes = new ArrayList<>();

    @JsonProperty("edges")
    public List<DispatchRuleEdge> edges = new ArrayList<>();
}
