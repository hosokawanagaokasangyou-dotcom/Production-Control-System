package jp.co.pm.ai.desktop.dispatch.rules.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

/** 工程名+機械名の配台優先（同一実機械上の連続選好）。 */
@JsonIgnoreProperties(ignoreUnknown = true)
public final class ProcessMachinePriorityEntry {

    public static final String PRIORITY_HIGHEST = "最優先";
    public static final String PRIORITY_HIGH = "優先";
    public static final String PRIORITY_NORMAL = "通常";
    public static final String PRIORITY_LOW = "優先度低";

    @JsonProperty("processName")
    public String processName = "";

    @JsonProperty("machineName")
    public String machineName = "";

    @JsonProperty("priority")
    public String priority = PRIORITY_NORMAL;

    @JsonProperty("consecutive")
    public boolean consecutive = true;

    @JsonProperty("enabled")
    public boolean enabled = true;

    public static ProcessMachinePriorityEntry defaultEmboss() {
        ProcessMachinePriorityEntry e = new ProcessMachinePriorityEntry();
        e.processName = "エンボス";
        e.machineName = "";
        e.priority = PRIORITY_NORMAL;
        e.consecutive = true;
        e.enabled = true;
        return e;
    }
}
