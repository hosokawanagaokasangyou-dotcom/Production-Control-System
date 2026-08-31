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
        String locked =
                MasterDispatchSheetEditRules.combinationRowStyle("巻返し", "機1", "", false, true);
        String added =
                MasterDispatchSheetEditRules.combinationRowStyle("巻返し", "機1", "", true, false);
        assertTrue(locked.contains(MasterDispatchSheetEditRules.LOCKED_ROW_BG));
        assertTrue(added.contains(MasterDispatchSheetEditRules.ADDED_ROW_BG));
        assertNotEquals(locked, added);
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
        assertEquals("1", needAdded.get(2).get(4));
        List<List<String>> needWithSurplus =
                List.of(
                        List.of("工程名", "", "", "巻返し"),
                        List.of("機械名", "", "", "機1"),
                        List.of("必須人数", "", "", "2"),
                        List.of("配台時追加人数", "", "", "0"));
        List<List<String>> surplusAdded =
                MasterDispatchSheetEditRules.addEquipmentColumn(
                        MasterDispatchSheetEditRules.SheetKind.NEED,
                        needWithSurplus,
                        "分割",
                        "スライス機1");
        assertEquals("1", surplusAdded.get(2).get(4));
        assertEquals("0", surplusAdded.get(3).get(4));
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
    void visibilityMask_needLeadingThreeUsesSameFocusKeys() {
        List<String> titles =
                List.of("項目", "依頼NO条件", "備考", "巻返し\n機1", "分割\nスライス機1", "");
        boolean[] vis =
                MasterDispatchSheetEditRules.visibilityMask(
                        titles,
                        3,
                        java.util.Set.of(
                                jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader
                                        .normalizedComboKey("分割", "スライス機1")));
        assertTrue(vis[0]);
        assertTrue(vis[1]);
        assertTrue(vis[2]);
        assertFalse(vis[3]);
        assertTrue(vis[4]);
        assertFalse(vis[5]);
    }

    @Test
    void focusKeysFromVisibility_subsetAndAllSelectedIsEmptyFocus() {
        List<String> titles =
                List.of("メンバー", "巻返し\n機1", "分割\nLAC/EC機", "分割\nスライス機1", "");
        java.util.Set<String> subset =
                MasterDispatchSheetEditRules.focusKeysFromVisibility(
                        titles, 1, new boolean[] {true, false, true, true, false});
        assertEquals(2, subset.size());
        assertTrue(
                subset.contains(
                        jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader.normalizedComboKey(
                                "分割", "LAC/EC機")));
        assertTrue(
                MasterDispatchSheetEditRules.focusKeysFromVisibility(
                                titles, 1, new boolean[] {true, true, true, true, false})
                        .isEmpty());
    }

    @Test
    void combinationDisplayRowVisible_filtersByProcessAndMachine() {
        List<String> header = List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械");
        java.util.Set<String> focus =
                java.util.Set.of(
                        jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader.normalizedComboKey(
                                "分割", "スライス機1"));
        assertTrue(
                MasterDispatchSheetEditRules.combinationDisplayRowVisible(
                        header, List.of("1", "分割", "スライス機1", "分割+スライス機1"), focus));
        assertFalse(
                MasterDispatchSheetEditRules.combinationDisplayRowVisible(
                        header, List.of("2", "巻返し", "機1", "巻返し+機1"), focus));
        assertTrue(
                MasterDispatchSheetEditRules.combinationDisplayRowVisible(
                        header, List.of("2", "巻返し", "機1", "巻返し+機1"), java.util.Set.of()));
        assertTrue(
                MasterDispatchSheetEditRules.combinationDisplayRowVisible(
                        header, List.of("", "", "", ""), focus));
    }

    @Test
    void combinationHiddenGridRows_hidesNonMatchingDataRowsAfterFilterRow() {
        List<List<String>> rows =
                List.of(
                        List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械"),
                        List.of("1", "分割", "スライス機1", "分割+スライス機1"),
                        List.of("2", "巻返し", "機1", "巻返し+機1"));
        java.util.Set<String> focus =
                java.util.Set.of(
                        jp.co.pm.ai.desktop.io.MasterTeamCombinationTableReader.normalizedComboKey(
                                "分割", "スライス機1"));
        java.util.BitSet hidden =
                MasterDispatchSheetEditRules.combinationHiddenGridRows(rows, 24, 1, focus);
        assertFalse(hidden.get(0));
        assertFalse(hidden.get(1));
        assertTrue(hidden.get(2));
        assertTrue(
                MasterDispatchSheetEditRules.combinationHiddenGridRows(
                                rows, 24, 1, java.util.Set.of())
                        .isEmpty());
    }

    @Test
    void dialogLabel_replacesNewlineWithSlash() {
        assertEquals("分割 / LAC/EC機", MasterDispatchSheetEditRules.dialogColumnLabel("分割\nLAC/EC機"));
        assertEquals("メンバー", MasterDispatchSheetEditRules.dialogColumnLabel("メンバー"));
    }

    @Test
    void isSkillsSkillValueCell_falseOnSpeedBaseSpeedEvenIfLooksLikeData() {
        List<List<String>> speed =
                List.of(
                        List.of("工程名", "", "", "分割"),
                        List.of("機械名", "", "", "LAC/EC機"),
                        List.of("基本速度", "", "", "20"),
                        List.of("実稼働比率", "", "", "0.95"));
        assertTrue(
                MasterDispatchSheetEditRules.isSkillsSkillValueCell(2, 3, speed));
        assertFalse(
                MasterDispatchSheetEditRules.isSkillsSkillValueCell(
                        MasterDispatchSheetEditRules.SheetKind.SPEED, 2, 3, speed));
        assertTrue(
                MasterDispatchSheetEditRules.isSpeedBaseSpeedCell(2, 3, speed));
        assertTrue(
                MasterDispatchSheetEditRules.isSpeedNumericCell(3, 3, speed));
        assertFalse(
                MasterDispatchSheetEditRules.isSpeedBaseSpeedCell(3, 3, speed));
        assertTrue(
                MasterDispatchSheetEditRules.isSpeedRatioCell(3, 3, speed));
        List<List<String>> withSpecial =
                List.of(
                        List.of("工程名", "", "", "分割"),
                        List.of("機械名", "", "", "LAC/EC機"),
                        List.of("基本速度", "", "", "20"),
                        List.of("実稼働比率", "", "", "0.95"),
                        List.of("特別指定1", "", "", "x"));
        assertFalse(
                MasterDispatchSheetEditRules.isSpeedNumericCell(4, 3, withSpecial));
    }

    @Test
    void speedBaseDecimal_allowsZeroToNinetyNineWithOneFractionDigit() {
        assertTrue(MasterDispatchSheetEditRules.isSpeedBaseDecimalValid("0.0"));
        assertTrue(MasterDispatchSheetEditRules.isSpeedBaseDecimalValid("0"));
        assertTrue(MasterDispatchSheetEditRules.isSpeedBaseDecimalValid("20"));
        assertTrue(MasterDispatchSheetEditRules.isSpeedBaseDecimalValid("20.5"));
        assertTrue(MasterDispatchSheetEditRules.isSpeedBaseDecimalValid("99.0"));
        assertFalse(MasterDispatchSheetEditRules.isSpeedBaseDecimalValid("-0.1"));
        assertFalse(MasterDispatchSheetEditRules.isSpeedBaseDecimalValid("99.1"));
        assertFalse(MasterDispatchSheetEditRules.isSpeedBaseDecimalValid("100"));
        assertFalse(MasterDispatchSheetEditRules.isSpeedBaseDecimalValid("20.55"));
        assertFalse(MasterDispatchSheetEditRules.isSpeedBaseDecimalValid("速い"));
        assertEquals("20.5", MasterDispatchSheetEditRules.formatSpeedBaseDecimal("20.50"));
        assertEquals("20.0", MasterDispatchSheetEditRules.formatSpeedBaseDecimal("20"));
        assertTrue(MasterDispatchSheetEditRules.isSpeedRatioDecimalValid("0.95"));
        assertTrue(MasterDispatchSheetEditRules.isSpeedRatioDecimalValid("0.7"));
        assertTrue(MasterDispatchSheetEditRules.isSpeedRatioDecimalValid("1"));
        assertFalse(MasterDispatchSheetEditRules.isSpeedRatioDecimalValid("95"));
        assertFalse(MasterDispatchSheetEditRules.isSpeedRatioDecimalValid("1.01"));
    }

    @Test
    void validateSpeed_baseSpeedRejectsOutOfRangeAndTwoDecimals() {
        List<List<String>> ok =
                List.of(
                        List.of("工程名", "", "", "巻返し"),
                        List.of("機械名", "", "", "機1"),
                        List.of("基本速度", "", "", "20.5"),
                        List.of("実稼働比率", "", "", "0.95"));
        assertTrue(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.SPEED, ok)
                        .isEmpty());
        List<List<String>> bad =
                List.of(
                        List.of("工程名", "", "", "巻返し"),
                        List.of("機械名", "", "", "機1"),
                        List.of("基本速度", "", "", "20.55"),
                        List.of("実稼働比率", "", "", "0.95"));
        assertFalse(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.SPEED, bad)
                        .isEmpty());
        List<List<String>> badRatio =
                List.of(
                        List.of("工程名", "", "", "巻返し"),
                        List.of("機械名", "", "", "機1"),
                        List.of("基本速度", "", "", "20.5"),
                        List.of("実稼働比率", "", "", "95"));
        List<String> ratioErrors =
                MasterDispatchSheetEditRules.validateForSave(
                        MasterDispatchSheetEditRules.SheetKind.SPEED, badRatio);
        assertFalse(ratioErrors.isEmpty());
        assertTrue(ratioErrors.get(0).contains("加工速度"), ratioErrors.toString());
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
        assertFalse(
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
    void skills_validateRejectsInvalidToken_allowsMultipleOpAsWithoutPriority() {
        List<List<String>> ok =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP"),
                        List.of("佐藤", "AS"),
                        List.of("鈴木", "OP1"));
        assertTrue(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.SKILLS, ok)
                        .isEmpty());

        List<List<String>> twoOp =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP"),
                        List.of("佐藤", "OP"));
        assertTrue(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.SKILLS, twoOp)
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
    }

    @Test
    void skills_normalizeStripsUnusedPriorityDigits() {
        List<List<String>> rows =
                List.of(
                        List.of("工程名", "巻返し"),
                        List.of("機械名", "機1"),
                        List.of("山田", "OP4"),
                        List.of("佐藤", "AS 3"),
                        List.of("鈴木", "OP"));
        List<List<String>> out =
                MasterDispatchSheetEditRules.normalizeOnExtract(
                        MasterDispatchSheetEditRules.SheetKind.SKILLS, rows);
        assertEquals("OP", out.get(2).get(1));
        assertEquals("AS", out.get(3).get(1));
        assertEquals("OP", out.get(4).get(1));
    }

    @Test
    void skills_canonicalSkillToken_dropsDigits() {
        assertEquals("OP", MasterDispatchSheetEditRules.canonicalSkillToken("OP4"));
        assertEquals("AS", MasterDispatchSheetEditRules.canonicalSkillToken("as3"));
        assertEquals("OP", MasterDispatchSheetEditRules.canonicalSkillToken("OP"));
        assertEquals("", MasterDispatchSheetEditRules.canonicalSkillToken(""));
        assertEquals("見習", MasterDispatchSheetEditRules.canonicalSkillToken("見習"));
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
        assertFalse(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.NEED, 0, 3, rows));
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
        List<List<String>> decimalPrio =
                List.of(
                        List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械", "組み合わせ優先度", "必須人数"),
                        List.of("1", "巻返し", "機1", "", "0", "2"),
                        List.of("2", "巻返し", "機1", "", "0.5", "2"),
                        List.of("3", "巻返し", "機1", "", "8.6", "2"));
        assertTrue(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, decimalPrio)
                        .isEmpty());
        assertFalse(
                MasterDispatchSheetEditRules.isInvalidValue(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, 1, 4, decimalPrio));
        List<List<String>> negativePrio =
                List.of(
                        List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械", "組み合わせ優先度", "必須人数"),
                        List.of("1", "巻返し", "機1", "", "-0.1", "2"));
        assertFalse(
                MasterDispatchSheetEditRules.validateForSave(
                                MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, negativePrio)
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

    @Test
    void combinationRowLock_blocksEditsExceptLockColumn() {
        List<List<String>> rows =
                List.of(
                        List.of(
                                "組み合わせ行ID",
                                "工程名",
                                "機械名",
                                "工程+機械",
                                "メンバー1",
                                "編集ロック",
                                "追加行"),
                        List.of("1", "巻返し", "機1", "巻返し+機1", "山田", "ロック", ""),
                        List.of("2", "巻返し", "機1", "巻返し+機1", "佐藤", "", ""));
        assertFalse(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, 1, 1, rows));
        assertFalse(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, 1, 4, rows));
        assertTrue(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, 1, 5, rows));
        assertTrue(
                MasterDispatchSheetEditRules.isEditable(
                        MasterDispatchSheetEditRules.SheetKind.COMBINATIONS, 2, 1, rows));
    }

    @Test
    void addCombinationRow_appendsMarkedRow_andSkipsDuplicateEquipment() {
        List<List<String>> rows =
                List.of(
                        List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械", "組み合わせ優先度", "必須人数", "メンバー1"),
                        List.of("1", "巻返し", "機1", "巻返し+機1", "1", "2", "山田"));
        List<List<String>> added =
                MasterDispatchSheetEditRules.addCombinationRow(rows, "分割", "スライス機1");
        assertEquals(3, added.size());
        assertEquals("2", added.get(2).get(0));
        assertEquals("分割", added.get(2).get(1));
        assertEquals("スライス機1", added.get(2).get(2));
        assertEquals("分割+スライス機1", added.get(2).get(3));
        int addedCol =
                MasterDispatchSheetEditRules.headerIndex(added.get(0), MasterDispatchSheetEditRules.COL_ADDED_ROW);
        int lockCol =
                MasterDispatchSheetEditRules.headerIndex(added.get(0), MasterDispatchSheetEditRules.COL_EDIT_LOCK);
        assertTrue(addedCol >= 0);
        assertTrue(lockCol >= 0);
        assertEquals(MasterDispatchSheetEditRules.ADDED_FLAG, added.get(2).get(addedCol));
        assertTrue(MasterDispatchSheetEditRules.isAddedCombinationRow(added, 2));
        assertFalse(MasterDispatchSheetEditRules.isAddedCombinationRow(added, 1));
        List<List<String>> again =
                MasterDispatchSheetEditRules.addCombinationRow(added, "分割", "スライス機1");
        assertEquals(added, again);
        List<List<String>> skipExisting =
                MasterDispatchSheetEditRules.addCombinationRow(rows, "巻返し", "機1");
        assertEquals(2, skipExisting.size());
    }

    @Test
    void deleteCombinationRows_skipsLockedAndRemovesUnlocked() {
        List<List<String>> rows =
                List.of(
                        List.of("組み合わせ行ID", "工程名", "機械名", "工程+機械", "編集ロック"),
                        List.of("1", "巻返し", "機1", "巻返し+機1", "ロック"),
                        List.of("2", "分割", "スライス機1", "分割+スライス機1", ""));
        List<List<String>> out =
                MasterDispatchSheetEditRules.deleteCombinationRows(rows, java.util.Set.of(1, 2));
        assertEquals(2, out.size());
        assertEquals("1", out.get(1).get(0));
    }

    @Test
    void combinationMemberChoices_onlyQualifiedSkillsMembersAsOpAsName() {
        List<List<String>> skills =
                List.of(
                        List.of("工程名", "巻返し", "分割"),
                        List.of("機械名", "機1", "スライス機1"),
                        List.of("山田", "OP1", ""),
                        List.of("佐藤", "AS2", "OP3"),
                        List.of("鈴木", "", "AS1"));
        List<String> choices =
                MasterDispatchSheetEditRules.combinationMemberChoices(
                        skills, "分割", "スライス機1", "");
        assertEquals(List.of("", "OP 佐藤", "AS 鈴木"), choices);
        List<String> withCurrent =
                MasterDispatchSheetEditRules.combinationMemberChoices(
                        skills, "分割", "スライス機1", "山田");
        assertTrue(withCurrent.contains("山田"));
        assertEquals(
                "佐藤", MasterDispatchSheetEditRules.combinationMemberName("OP 佐藤"));
        assertEquals(
                "山田", MasterDispatchSheetEditRules.combinationMemberName("山田"));
    }

    @Test
    void comboRowStyle_addedRowUsesDistinctColor() {
        String grouped =
                MasterDispatchSheetEditRules.combinationRowStyle("分割", "スライス機1", "", false, false);
        String added =
                MasterDispatchSheetEditRules.combinationRowStyle("分割", "スライス機1", "", true, false);
        assertFalse(grouped.isEmpty());
        assertFalse(added.isEmpty());
        assertNotEquals(grouped, added);
        assertTrue(added.contains("#f4c36a"));
    }
}
