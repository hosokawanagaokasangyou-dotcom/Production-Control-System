package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.EnumSet;
import java.util.List;

import org.junit.jupiter.api.Test;

class MasterDispatchSetupCompletenessTest {

    @Test
    void evaluate_requiresAllFourStepsWithContent() {
        List<MasterDispatchSetupCompleteness.EquipmentRef> equipment =
                List.of(new MasterDispatchSetupCompleteness.EquipmentRef("穴あけ", "SEC機 湘南", "T1"));
        List<List<String>> skills =
                List.of(
                        List.of("工程名", "穴あけ"),
                        List.of("機械名", "SEC機 湘南"),
                        List.of("岡田", "OP"),
                        List.of("宮島", "OP"));
        List<List<String>> need =
                List.of(
                        List.of("工程名", "", "", "穴あけ"),
                        List.of("機械名", "", "", "SEC機 湘南"),
                        List.of("基本必要人数", "", "", "1"));
        List<List<String>> combo =
                List.of(
                        List.of(
                                "組み合わせ行ID",
                                "工程名",
                                "機械名",
                                "工程+機械",
                                "必須人数",
                                "メンバー1"),
                        List.of("1", "穴あけ", "SEC機 湘南", "穴あけ+SEC機 湘南", "1", "OP 岡田"));
        List<List<String>> speed =
                List.of(
                        List.of("工程名", "", "", "穴あけ"),
                        List.of("機械名", "", "", "SEC機 湘南"),
                        List.of("基本速度", "", "", "20.0"));

        var ok =
                MasterDispatchSetupCompleteness.evaluate(
                        equipment, skills, need, combo, speed);
        assertTrue(ok.allComplete());

        var noSkill =
                MasterDispatchSetupCompleteness.evaluate(
                        equipment,
                        List.of(
                                List.of("工程名", "穴あけ"),
                                List.of("機械名", "SEC機 湘南"),
                                List.of("岡田", "")),
                        need,
                        combo,
                        speed);
        assertFalse(noSkill.allComplete());
        assertEquals(
                EnumSet.of(MasterDispatchSetupCompleteness.Step.SKILLS),
                noSkill.incomplete().get(0).incompleteSteps());

        var noSpeed =
                MasterDispatchSetupCompleteness.evaluate(
                        equipment,
                        skills,
                        need,
                        combo,
                        List.of(
                                List.of("工程名", "", "", "穴あけ"),
                                List.of("機械名", "", "", "SEC機 湘南"),
                                List.of("基本速度", "", "", "")));
        assertFalse(noSpeed.allComplete());
        assertTrue(
                noSpeed.incomplete()
                        .get(0)
                        .incompleteSteps()
                        .contains(MasterDispatchSetupCompleteness.Step.SPEED));
    }

    @Test
    void combinationsComplete_requiresMemberCountMatchingNeed() {
        List<List<String>> combo =
                List.of(
                        List.of(
                                "組み合わせ行ID",
                                "工程名",
                                "機械名",
                                "工程+機械",
                                "必須人数",
                                "メンバー1",
                                "メンバー2"),
                        List.of(
                                "1",
                                "分割",
                                "スライス機1",
                                "分割+スライス機1",
                                "2",
                                "OP 佐藤",
                                ""));
        assertFalse(
                MasterDispatchSetupCompleteness.combinationsComplete(
                        combo, "分割", "スライス機1", 2));
        List<List<String>> filled =
                List.of(
                        combo.get(0),
                        List.of(
                                "1",
                                "分割",
                                "スライス機1",
                                "分割+スライス機1",
                                "2",
                                "OP 佐藤",
                                "AS 鈴木"));
        assertTrue(
                MasterDispatchSetupCompleteness.combinationsComplete(
                        filled, "分割", "スライス機1", 2));
    }

    @Test
    void ensureSkillCombinations_generatesRowsByNeed() {
        List<List<String>> skills =
                List.of(
                        List.of("工程名", "分割"),
                        List.of("機械名", "スライス機1"),
                        List.of("佐藤", "OP"),
                        List.of("鈴木", "AS"),
                        List.of("高橋", "OP"));
        List<List<String>> need =
                List.of(
                        List.of("工程名", "", "", "分割"),
                        List.of("機械名", "", "", "スライス機1"),
                        List.of("必須人数", "", "", "1"));
        List<List<String>> combo =
                List.of(
                        List.of(
                                "組み合わせ行ID",
                                "工程名",
                                "機械名",
                                "工程+機械",
                                "必須人数",
                                "メンバー1"));
        List<List<String>> out =
                MasterDispatchSheetEditRules.ensureSkillCombinations(
                        combo, "分割", "スライス機1", skills, need);
        assertEquals(4, out.size());
        assertEquals("OP 佐藤", out.get(1).get(5));
        assertEquals("AS 鈴木", out.get(2).get(5));
        assertEquals("OP 高橋", out.get(3).get(5));
    }
}
