package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class MasterDispatchSheetEditRulesTest {

    @Test
    void skipFilterRow_dropsLeadingFilterRowAndKeepsData() {
        List<List<String>> grid =
                List.of(
                        List.of("", ""),
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"));
        assertEquals(
                List.of(List.of("工程名", "巻返し"), List.of("機械名", "機1")),
                MasterDispatchSheetEditRules.skipFilterRow(grid));
    }

    @Test
    void preferredColumnWidths_longProcessNameIsWiderThanShort() {
        List<List<String>> rows =
                List.of(List.of("工程名", "巻", "巻返し工程（湖南）"));
        List<Double> widths = MasterDispatchSheetEditRules.preferredColumnWidths(rows, 4);
        assertTrue(widths.get(2) > widths.get(1), widths.toString());
        assertTrue(widths.get(2) >= 120.0, widths.toString());
        assertTrue(widths.get(3) < widths.get(2), widths.toString());
    }

    @Test
    void columnTitles_skillsUseTwoLineProcessAndMachine() {
        List<List<String>> rows =
                List.of(
                        List.of("工程名", "巻返し", "EC"),
                        List.of("機械名", "機1", "EC機 湖南"));
        List<String> titles =
                MasterDispatchSheetEditRules.columnTitles(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, rows, 4);
        assertEquals("メンバー", titles.get(0));
        assertEquals("巻返し\n機1", titles.get(1));
        assertEquals("EC\nEC機 湖南", titles.get(2));
        assertEquals("", titles.get(3));
    }

    @Test
    void columnTitles_needUsesDetectedHeadersAndNeedLabels() {
        List<List<String>> rows =
                List.of(
                        List.of("工程名", "", "", "巻返し"),
                        List.of("機械名", "", "", "機1"),
                        List.of("必須人数", "", "", "2"));
        List<String> titles =
                MasterDispatchSheetEditRules.columnTitles(
                        MasterDispatchSheetEditRules.SheetKind.NEED, rows, 4);
        assertEquals("項目", titles.get(0));
        assertEquals("依頼NO条件", titles.get(1));
        assertEquals("備考", titles.get(2));
        assertEquals("巻返し\n機1", titles.get(3));
    }

    @Test
    void columnTitles_combinationsUseFirstRow() {
        List<List<String>> rows =
                List.of(
                        List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械"),
                        List.of("1", "巻返し", "機1", "巻返し+機1"));
        List<String> titles =
                MasterDispatchSheetEditRules.columnTitles(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, rows, 4);
        assertEquals("組み合わせ行ID", titles.get(0));
        assertEquals("工程名", titles.get(1));
        assertEquals("機械名", titles.get(2));
        assertEquals("工程+機械", titles.get(3));
    }

    @Test
    void displayRows_keepProcessMachineRowsForSkills_andRestoreFromColumnTitles() {
        List<List<String>> skills =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP1"));
        assertEquals(
                skills,
                MasterDispatchSheetEditRules.displayRows(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, skills));
        List<String> titles =
                MasterDispatchSheetEditRules.columnTitles(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, skills, 2);
        assertEquals(
                skills,
                MasterDispatchSheetEditRules.restoreTitleRows(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS,
                        titles,
                        skills));

        List<List<String>> combo =
                List.of(
                        List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械"),
                        List.of("1", "巻返し", "機1", "巻返し+機1"));
        assertEquals(
                List.of(List.of("1", "巻返し", "機1", "巻返し+機1")),
                MasterDispatchSheetEditRules.displayRows(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, combo));
        assertEquals(
                combo,
                MasterDispatchSheetEditRules.restoreTitleRows(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS,
                        combo.get(0),
                        List.of(List.of("1", "巻返し", "機1", "巻返し+機1"))));
    }

    @Test
    void displayRows_omitProcessAliasAndNeedCaptionRows() {
        List<List<String>> need =
                List.of(
                        List.of("工程名", "", "", "EC"),
                        List.of("機械名", "", "", "EC機 湖南"),
                        List.of("工程名(通称)", "", "", ""),
                        List.of("基本必要人数", "", "", "2"),
                        List.of("余力時追加人数", "", "", "0"),
                        List.of("", "依頼NO条件", "備考", ""));
        assertEquals(
                List.of(
                        List.of("工程名", "", "", "EC"),
                        List.of("機械名", "", "", "EC機 湖南"),
                        List.of("基本必要人数", "", "", "2"),
                        List.of("余力時追加人数", "", "", "0")),
                MasterDispatchSheetEditRules.displayRows(
                        MasterDispatchSheetEditRules.SheetKind.NEED, need));

        List<List<String>> speed =
                List.of(
                        List.of("工程名", "", "", "EC"),
                        List.of("機械名", "", "", "EC機 湖南"),
                        List.of("工程名(通称)", "", "", ""),
                        List.of("基本速度", "", "", "20"));
        assertEquals(
                List.of(
                        List.of("工程名", "", "", "EC"),
                        List.of("機械名", "", "", "EC機 湖南"),
                        List.of("基本速度", "", "", "20")),
                MasterDispatchSheetEditRules.displayRows(
                        MasterDispatchSheetEditRules.SheetKind.SPEED, speed));
        List<String> titles =
                MasterDispatchSheetEditRules.columnTitles(
                        MasterDispatchSheetEditRules.SheetKind.SPEED, speed, 4);
        assertEquals(
                List.of(
                        List.of("工程名", "", "", "EC"),
                        List.of("機械名", "", "", "EC機 湖南"),
                        List.of("基本速度", "", "", "20")),
                MasterDispatchSheetEditRules.restoreTitleRows(
                        MasterDispatchSheetEditRules.SheetKind.SPEED,
                        titles,
                        List.of(List.of("基本速度", "", "", "20"))));
    }

    @Test
    void comboRowStyle_sameEquipmentProcessKeyIsStable() {
        String a = MasterDispatchSheetEditRules.comboRowStyle("巻返し", "機1");
        String b = MasterDispatchSheetEditRules.comboRowStyle("巻返し", "機1");
        String c = MasterDispatchSheetEditRules.comboRowStyle("スリット", "機1");
        assertFalse(a.isEmpty());
        assertEquals(a, b);
        assertNotEquals(a, c);
        assertEquals("", MasterDispatchSheetEditRules.comboRowStyle("", "機1"));
    }

    @Test
    void addEquipmentColumn_appendsProcessAndMachine_andSkipsDuplicate() {
        List<List<String>> skills =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP1"));
        List<List<String>> added =
                MasterDispatchSheetEditRules.addEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS,
                        skills,
                        "分割",
                        "LAC/EC機");
        assertEquals("分割", added.get(0).get(2));
        assertEquals("LAC/EC機", added.get(1).get(2));
        assertEquals("", added.get(2).get(2));
        List<List<String>> again =
                MasterDispatchSheetEditRules.addEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS,
                        added,
                        "分割",
                        "LAC/EC機");
        assertEquals(added, again);
        assertTrue(
                MasterDispatchSheetEditRules.containsEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, added, "分割", "LAC/EC機"));
        assertFalse(
                MasterDispatchSheetEditRules.containsEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, skills, "分割", "LAC/EC機"));

        List<List<String>> need =
                List.of(
                        List.of("工程名", "", "", "巻返し"),
                        List.of("機械名", "", "", "機1"),
                        List.of("必須人数", "", "", "2"));
        List<List<String>> needAdded =
                MasterDispatchSheetEditRules.addEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.NEED, need, "分割", "スライス機1");
        assertEquals("分割", needAdded.get(0).get(4));
        assertEquals("スライス機1", needAdded.get(1).get(4));
        assertEquals("2", needAdded.get(2).get(3));
    }

    @Test
    void visibilityMask_focusKeysShowOnlyThoseEquipmentColumns() {
        List<String> titles =
                List.of("メンバー", "巻返し\n機1", "分割\nLAC/EC機", "分割\nスライス機1", "");
        boolean[] vis =
                MasterDispatchSheetEditRules.visibilityMask(
                        titles,
                        1,
                        java.util.Set.of(
                                jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader
                                        .normalizedComboKey("分割", "LAC/EC機"),
                                jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader
                                        .normalizedComboKey("分割", "スライス機1")));
        assertTrue(vis[0]);
        assertFalse(vis[1]);
        assertTrue(vis[2]);
        assertTrue(vis[3]);
        assertFalse(vis[4]);
    }

    @Test
    void visibilityMask_withoutFocusHidesEmptyExtraColumnsOnly() {
        List<String> titles = List.of("メンバー", "巻返し\n機1", "", "");
        boolean[] vis = MasterDispatchSheetEditRules.visibilityMask(titles, 1, java.util.Set.of());
        assertTrue(vis[0]);
        assertTrue(vis[1]);
        assertFalse(vis[2]);
        assertFalse(vis[3]);
    }

    @Test
    void dialogLabel_replacesNewlineWithSlash() {
        assertEquals("分割 / LAC/EC機", MasterDispatchSheetEditRules.dialogColumnLabel("分割\nLAC/EC機"));
        assertEquals("メンバー", MasterDispatchSheetEditRules.dialogColumnLabel("メンバー"));
    }

    @Test
    void isSkillsSkillValueCell_falseOnHeaderRows() {
        List<List<String>> rows =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP1"));
        assertFalse(
                MasterDispatchSheetEditRules.isSkillsSkillValueCell(0, 1, rows));
        assertFalse(
                MasterDispatchSheetEditRules.isSkillsSkillValueCell(1, 1, rows));
        assertTrue(
                MasterDispatchSheetEditRules.isSkillsSkillValueCell(2, 1, rows));
        assertFalse(
                MasterDispatchSheetEditRules.isSkillsSkillValueCell(2, 0, rows));
    }

    @Test
    void titleRowKind_processAndMachineRows() {
        List<List<String>> skills =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP1"));
        assertEquals(
                MasterDispatchSheetEditRules.TitleRowKind.PROCESS,
                MasterDispatchSheetEditRules.titleRowKind(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, 0, skills));
        assertEquals(
                MasterDispatchSheetEditRules.TitleRowKind.MACHINE,
                MasterDispatchSheetEditRules.titleRowKind(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, 1, skills));
        assertEquals(
                MasterDispatchSheetEditRules.TitleRowKind.NONE,
                MasterDispatchSheetEditRules.titleRowKind(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, 2, skills));
    }

    @Test
    void skills_locksStructureLabelsAndAllowsMemberAndSkillCells() {
        List<List<String>> rows =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP1"));
        assertFalse(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, 0, 0, rows));
        assertFalse(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, 1, 0, rows));
        assertTrue(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, 0, 1, rows));
        assertTrue(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, 2, 0, rows));
        assertTrue(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, 2, 1, rows));
    }

    @Test
    void skills_validateRejectsInvalidTokenAndDuplicatePriority() {
        List<List<String>> ok =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP1"),
                        List.of("佐藤", "AS2"));
        assertTrue(MasterDispatchSheetEditRules.validateForSave(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, ok)
                .isEmpty());

        List<List<String>> badToken =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "見習"));
        assertFalse(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.SKILLS, badToken)
                        .isEmpty());

        List<List<String>> dup =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP1"),
                        List.of("佐藤", "AS1"));
        assertFalse(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.SKILLS, dup)
                        .isEmpty());

        List<List<String>> opEqualsOp1 =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP"),
                        List.of("佐藤", "AS1"));
        assertFalse(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.SKILLS, opEqualsOp1)
                        .isEmpty());
    }

    @Test
    void skills_normalizeOpToOp1() {
        List<List<String>> rows =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP"));
        List<List<String>> out =
                MasterDispatchSheetEditRules.normalizeOnExtract(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, rows);
        assertEquals("OP1", out.get(2).get(1));
    }

    @Test
    void need_locksLabelsAndRejectsNonIntegerCounts() {
        List<List<String>> rows =
                List.of(
                        List.of("工程名", "条件", "備考", "巻返し"),
                        List.of("機械名", "", "", "機1"),
                        List.of("必須人数", "", "", "2"),
                        List.of("配台時追加人数", "", "", "1"));
        assertFalse(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.NEED, 0, 0, rows));
        assertTrue(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.NEED, 2, 3, rows));
        assertTrue(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.NEED, rows)
                        .isEmpty());
        List<List<String>> bad =
                List.of(
                        List.of("工程名", "条件", "備考", "巻返し"),
                        List.of("機械名", "", "", "機1"),
                        List.of("必須人数", "", "", "abc"));
        assertFalse(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.NEED, bad)
                        .isEmpty());
    }

    @Test
    void speed_rejectsNonNumericBaseAndRatio() {
        List<List<String>> rows =
                List.of(
                        List.of("工程名", "", "", "巻返し"),
                        List.of("機械名", "", "", "機1"),
                        List.of("", "", "", ""),
                        List.of("基本速度", "", "", "20"),
                        List.of("実稼働比率", "", "", "0.7"));
        assertTrue(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.SPEED, rows)
                        .isEmpty());
        List<List<String>> bad =
                List.of(
                        List.of("工程名", "", "", "巻返し"),
                        List.of("機械名", "", "", "機1"),
                        List.of("", "", "", ""),
                        List.of("基本速度", "", "", "速い"),
                        List.of("実稼働比率", "", "", "0.7"));
        assertFalse(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.SPEED, bad)
                        .isEmpty());
    }

    @Test
    void combinations_headerLockedAndProcessMachineSynced() {
        List<List<String>> rows =
                List.of(
                        List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械", "組み合わせ優先度", "必須人数", "メンバー1"),
                        List.of("1", "巻返し", "機1", "", "1", "2", "山田"));
        assertFalse(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, 0, 0, rows));
        assertFalse(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, 1, 3, rows));
        assertTrue(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, 1, 1, rows));
        List<List<String>> out =
                MasterDispatchSheetEditRules.normalizeOnExtract(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, rows);
        assertEquals("巻返し+機1", out.get(1).get(3));
        assertTrue(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, rows)
                        .isEmpty());
        List<List<String>> badPrio =
                List.of(
                        List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械", "組み合わせ優先度", "必須人数"),
                        List.of("1", "巻返し", "機1", "", "x", "2"));
        assertFalse(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, badPrio)
                        .isEmpty());
    }
}
