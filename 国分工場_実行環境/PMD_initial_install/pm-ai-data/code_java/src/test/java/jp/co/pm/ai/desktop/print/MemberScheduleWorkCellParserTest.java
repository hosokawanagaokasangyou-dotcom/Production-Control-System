package jp.co.pm.ai.desktop.print;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class MemberScheduleWorkCellParserTest {

    @Test
    void parse_sample_dispatches_req_process_machine() {
        MemberScheduleWorkCellParser.ParsedWorkCell p =
                MemberScheduleWorkCellParser.parse(
                        "[Y5-1] \u30b9\u30e9\u30a4\u30b9+\u30b9\u30e9\u30a4\u30b9\u6a5f1\u3000\u6e56\u5357(\u88dc)");
        assertEquals("Y5-1", p.requestNo());
        assertEquals("\u30b9\u30e9\u30a4\u30b9", p.processName());
        assertEquals("\u30b9\u30e9\u30a4\u30b9\u6a5f1\u3000\u6e56\u5357", p.machineName());
    }

    @Test
    void parse_plainText_falls_back_to_raw() {
        MemberScheduleWorkCellParser.ParsedWorkCell p =
                MemberScheduleWorkCellParser.parse("\u4f11");
        assertEquals("", p.requestNo());
        assertEquals("\u4f11", p.processName());
    }
}
