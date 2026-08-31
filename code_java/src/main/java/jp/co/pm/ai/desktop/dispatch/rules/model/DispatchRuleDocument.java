package jp.co.pm.ai.desktop.dispatch.rules.model;

import java.util.ArrayList;
import java.util.List;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

@JsonIgnoreProperties(ignoreUnknown = true)
public final class DispatchRuleDocument {

    @JsonProperty("schemaVersion")
    public int schemaVersion = 1;

    @JsonProperty("engineMinVersion")
    public String engineMinVersion = "1.0.0";

    @JsonProperty("savedAt")
    public String savedAt = "";

    @JsonInclude(JsonInclude.Include.NON_NULL)
    @JsonProperty("processMachinePriorities")
    public List<ProcessMachinePriorityEntry> processMachinePriorities;

    @JsonProperty("rules")
    public List<DispatchRuleEntry> rules = new ArrayList<>();
}
