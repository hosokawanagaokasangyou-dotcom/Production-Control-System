package jp.co.pm.ai.desktop.dispatch.rules.ui.trace;

import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.dispatch.rules.trace.DispatchRuleTraceLoader;

/** Lookup badges for dispatch interactive rows from sidecar events. */
public final class DispatchRuleApplicationBadgeSupport {

    private final Map<String, List<DispatchRuleTraceLoader.ApplicationEvent>> byTaskKey =
            new LinkedHashMap<>();

    public void reload(Map<String, String> ui) {
        byTaskKey.clear();
        try {
            for (DispatchRuleTraceLoader.ApplicationEvent e :
                    DispatchRuleTraceLoader.loadFromWorkDir(ui)) {
                String key = taskKey(e.taskId(), "", "");
                byTaskKey.computeIfAbsent(key, k -> new ArrayList<>()).add(e);
            }
        } catch (IOException ex) {
            // keep empty
        }
    }

    public String badgeForRow(String orderNo, String process, String machine) {
        String key = taskKey(orderNo, process, machine);
        List<DispatchRuleTraceLoader.ApplicationEvent> events = byTaskKey.get(key);
        if (events == null || events.isEmpty()) {
            events = byTaskKey.get(taskKey(orderNo, "", ""));
        }
        if (events == null || events.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        for (DispatchRuleTraceLoader.ApplicationEvent e : events) {
            if (!sb.isEmpty()) {
                sb.append(' ');
            }
            sb.append(iconFor(e.effect())).append(e.ruleId());
        }
        return sb.toString();
    }

    public String tooltipForRow(String orderNo, String process, String machine) {
        String key = taskKey(orderNo, process, machine);
        List<DispatchRuleTraceLoader.ApplicationEvent> events = byTaskKey.get(key);
        if (events == null || events.isEmpty()) {
            return "";
        }
        return events.stream()
                .map(e -> e.ruleId() + ": " + e.summaryJa())
                .reduce((a, b) -> a + "\n" + b)
                .orElse("");
    }

    private static String taskKey(String orderNo, String process, String machine) {
        return (orderNo != null ? orderNo.strip() : "")
                + "|"
                + (process != null ? process.strip() : "")
                + "|"
                + (machine != null ? machine.strip() : "");
    }

    private static String iconFor(String effect) {
        if (effect == null) {
            return "";
        }
        return switch (effect) {
            case "block_candidate" -> "🚫";
            case "set_speed_mpm" -> "⚡";
            default -> "●";
        };
    }
}
