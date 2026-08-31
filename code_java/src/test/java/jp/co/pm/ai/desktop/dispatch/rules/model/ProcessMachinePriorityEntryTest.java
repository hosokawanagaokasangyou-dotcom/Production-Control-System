package jp.co.pm.ai.desktop.dispatch.rules.model;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;

import com.fasterxml.jackson.databind.ObjectMapper;

import org.junit.jupiter.api.Test;

class ProcessMachinePriorityEntryTest {

    @Test
    void defaultEmbossIsNormalConsecutive() {
        ProcessMachinePriorityEntry e = ProcessMachinePriorityEntry.defaultEmboss();
        assertEquals("エンボス", e.processName);
        assertEquals("", e.machineName);
        assertEquals(ProcessMachinePriorityEntry.PRIORITY_NORMAL, e.priority);
        assertTrue(e.consecutive);
        assertTrue(e.enabled);
    }

    @Test
    void documentRoundtripKeepsPriorities() throws Exception {
        ObjectMapper json = new ObjectMapper();
        DispatchRuleDocument doc = new DispatchRuleDocument();
        doc.schemaVersion = 1;
        doc.processMachinePriorities = new ArrayList<>();
        doc.processMachinePriorities.add(ProcessMachinePriorityEntry.defaultEmboss());
        String raw = json.writeValueAsString(doc);
        DispatchRuleDocument back = json.readValue(raw, DispatchRuleDocument.class);
        assertEquals(1, back.processMachinePriorities.size());
        assertEquals("エンボス", back.processMachinePriorities.get(0).processName);
        assertEquals("通常", back.processMachinePriorities.get(0).priority);
    }

    @Test
    void emptyListWritesKey() throws Exception {
        ObjectMapper json = new ObjectMapper();
        DispatchRuleDocument doc = new DispatchRuleDocument();
        doc.processMachinePriorities = new ArrayList<>();
        String raw = json.writeValueAsString(doc);
        assertTrue(raw.contains("processMachinePriorities"));
        DispatchRuleDocument back = json.readValue(raw, DispatchRuleDocument.class);
        assertEquals(0, back.processMachinePriorities.size());
    }

    @Test
    void nullListOmitsKey() throws Exception {
        ObjectMapper json = new ObjectMapper();
        DispatchRuleDocument doc = new DispatchRuleDocument();
        doc.processMachinePriorities = null;
        String raw = json.writeValueAsString(doc);
        assertFalse(raw.contains("processMachinePriorities"));
    }
}
