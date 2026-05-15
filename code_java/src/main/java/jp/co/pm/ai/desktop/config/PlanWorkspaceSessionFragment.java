package jp.co.pm.ai.desktop.config;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

import jp.co.pm.ai.desktop.MainShellTabId;

/**
 * 配台ワークスペース復元用に {@link DesktopSessionState} から切り出したフィールド（配台入力・段階1・段階2成果物パス・設備ガント
 * 表示・担当バッジ位置など）。
 */
public record PlanWorkspaceSessionFragment(
        String planInputPath,
        String planInputSheet,
        String stage1PreviewPath,
        String stage1PreviewSheet,
        String mainRunStage2ProductionPlan,
        String mainRunStage2MemberSchedule,
        Map<String, Integer> innerTabPartial,
        double equipmentGanttGraphicZoomPercent,
        double equipmentGanttDateColWidth,
        double equipmentGanttMachineColWidth,
        double equipmentGanttProcessColWidth,
        String equipmentGanttBarFontFamily,
        double equipmentGanttBarFontPercent,
        double equipmentGanttRowHeightPercent,
        double equipmentGanttHeaderHeightPercent,
        double equipmentGanttSlotWidthPercent,
        double equipmentGanttShiftWheelHScrollPercent,
        double equipmentGanttPersonBadgeGapPx,
        double equipmentGanttPersonBadgeBandVerticalOffsetPx,
        String equipmentGanttGraphicDataFingerprint,
        Map<String, EquipmentGanttBadgeDragDelta> equipmentGanttBadgeDragDeltas,
        boolean equipmentGanttPersonBadgeDragAdjustEnabled,
        boolean equipmentGanttPersonBadgeEnabled,
        boolean equipmentGanttPersonBadgeWireEnabled,
        String equipmentGanttPersonBadgeWireStrokeHex,
        double equipmentGanttPersonBadgeWireWidthPx,
        String equipmentGanttPersonBadgeWireDashStyleKey,
        double equipmentGanttPersonBadgeWireMaxLengthPx,
        String equipmentGanttPersonBadgeFontFamily,
        double equipmentGanttPersonBadgeFontPercent,
        String equipmentGanttPersonBadgeFillHex,
        String equipmentGanttPersonBadgeTextHex,
        String equipmentGanttPersonBadgeStrokeHex,
        double equipmentGanttPersonBadgeStrokeWidth,
        double equipmentGanttPersonBadgeCornerRadius,
        boolean equipmentGanttPersonBadgePill,
        String equipmentGanttPersonBadgeGlowColorHex,
        double equipmentGanttPersonBadgeGlowRadius,
        double equipmentGanttPersonBadgeGlowSpread,
        double equipmentGanttPersonBadgeOpacity,
        Map<String, PersonBadgeStyle> equipmentGanttPersonBadgeStylesByLabel,
        Map<String, PersonBadgeStyle> equipmentGanttPersonBadgeStylesByMemberKey,
        String equipmentGanttPlanJsonPath) {

    private static final Set<String> INNER_CAPTURE_KEYS =
            Set.of(
                    MainShellTabId.PLAN_INPUT.key(),
                    MainShellTabId.STAGE1_PREVIEW.key(),
                    MainShellTabId.CODE_LOOKUP_TABLES.key(),
                    MainShellTabId.DISPATCH_INTERACTIVE.key(),
                    MainShellTabId.EQUIPMENT_GANTT_GRAPHIC.key(),
                    MainShellTabId.PLAN_RESULT_VIEWER.key(),
                    MainShellTabId.DELIVERY_CALENDAR_VIEW.key(),
                    MainShellTabId.RESULT_DISPATCH.key());

    public PlanWorkspaceSessionFragment {
        innerTabPartial =
                innerTabPartial == null || innerTabPartial.isEmpty()
                        ? Map.of()
                        : Map.copyOf(innerTabPartial);
        equipmentGanttBadgeDragDeltas =
                equipmentGanttBadgeDragDeltas == null || equipmentGanttBadgeDragDeltas.isEmpty()
                        ? Map.of()
                        : Map.copyOf(equipmentGanttBadgeDragDeltas);
        equipmentGanttPersonBadgeStylesByLabel =
                equipmentGanttPersonBadgeStylesByLabel == null
                                || equipmentGanttPersonBadgeStylesByLabel.isEmpty()
                        ? Map.of()
                        : Map.copyOf(equipmentGanttPersonBadgeStylesByLabel);
        equipmentGanttPersonBadgeStylesByMemberKey =
                equipmentGanttPersonBadgeStylesByMemberKey == null
                                || equipmentGanttPersonBadgeStylesByMemberKey.isEmpty()
                        ? Map.of()
                        : Map.copyOf(equipmentGanttPersonBadgeStylesByMemberKey);
    }

    public static PlanWorkspaceSessionFragment fromDesktopSession(DesktopSessionState s) {
        if (s == null) {
            return empty();
        }
        return new PlanWorkspaceSessionFragment(
                s.planInputPath(),
                s.planInputSheet(),
                s.stage1PreviewPath(),
                s.stage1PreviewSheet(),
                s.mainRunStage2ProductionPlan(),
                s.mainRunStage2MemberSchedule(),
                filterInner(s.innerTabSelectedIndexByShellTabKey()),
                s.equipmentGanttGraphicZoomPercent(),
                s.equipmentGanttDateColWidth(),
                s.equipmentGanttMachineColWidth(),
                s.equipmentGanttProcessColWidth(),
                s.equipmentGanttBarFontFamily(),
                s.equipmentGanttBarFontPercent(),
                s.equipmentGanttRowHeightPercent(),
                s.equipmentGanttHeaderHeightPercent(),
                s.equipmentGanttSlotWidthPercent(),
                s.equipmentGanttShiftWheelHScrollPercent(),
                s.equipmentGanttPersonBadgeGapPx(),
                s.equipmentGanttPersonBadgeBandVerticalOffsetPx(),
                s.equipmentGanttGraphicDataFingerprint(),
                s.equipmentGanttBadgeDragDeltas(),
                s.equipmentGanttPersonBadgeDragAdjustEnabled(),
                s.equipmentGanttPersonBadgeEnabled(),
                s.equipmentGanttPersonBadgeWireEnabled(),
                s.equipmentGanttPersonBadgeWireStrokeHex(),
                s.equipmentGanttPersonBadgeWireWidthPx(),
                s.equipmentGanttPersonBadgeWireDashStyleKey(),
                s.equipmentGanttPersonBadgeWireMaxLengthPx(),
                s.equipmentGanttPersonBadgeFontFamily(),
                s.equipmentGanttPersonBadgeFontPercent(),
                s.equipmentGanttPersonBadgeFillHex(),
                s.equipmentGanttPersonBadgeTextHex(),
                s.equipmentGanttPersonBadgeStrokeHex(),
                s.equipmentGanttPersonBadgeStrokeWidth(),
                s.equipmentGanttPersonBadgeCornerRadius(),
                s.equipmentGanttPersonBadgePill(),
                s.equipmentGanttPersonBadgeGlowColorHex(),
                s.equipmentGanttPersonBadgeGlowRadius(),
                s.equipmentGanttPersonBadgeGlowSpread(),
                s.equipmentGanttPersonBadgeOpacity(),
                s.equipmentGanttPersonBadgeStylesByLabel(),
                s.equipmentGanttPersonBadgeStylesByMemberKey(),
                s.equipmentGanttPlanJsonPath());
    }

    private static Map<String, Integer> filterInner(Map<String, Integer> full) {
        if (full == null || full.isEmpty()) {
            return Map.of();
        }
        LinkedHashMap<String, Integer> m = new LinkedHashMap<>();
        for (String k : INNER_CAPTURE_KEYS) {
            Integer v = full.get(k);
            if (v != null) {
                m.put(k, v);
            }
        }
        return Map.copyOf(m);
    }

    public static PlanWorkspaceSessionFragment empty() {
        return fromDesktopSession(DesktopSessionState.empty());
    }

    /**
     * 現在のセッション状態を保ったまま、本フラグメントが持つ配台ワークスペース系フィールドだけを上書きした
     * {@link DesktopSessionState} を返す。
     */
    public DesktopSessionState mergeOnto(DesktopSessionState base) {
        if (base == null) {
            return DesktopSessionState.empty();
        }
        Map<String, Integer> mergedInner = new LinkedHashMap<>(base.innerTabSelectedIndexByShellTabKey());
        for (Map.Entry<String, Integer> e : innerTabPartial.entrySet()) {
            if (e.getKey() != null && !e.getKey().isBlank() && e.getValue() != null) {
                mergedInner.put(e.getKey().trim(), e.getValue());
            }
        }
        return new DesktopSessionState(
                planInputPath != null ? planInputPath : "",
                planInputSheet != null ? planInputSheet : "",
                stage1PreviewPath != null ? stage1PreviewPath : "",
                stage1PreviewSheet != null ? stage1PreviewSheet : "",
                base.excludeRulesPath(),
                base.mainRunWorkbook(),
                base.mainRunScriptDir(),
                base.windowWidth(),
                base.windowHeight(),
                base.windowX(),
                base.windowY(),
                base.uiTheme(),
                base.logFontFamily(),
                base.logFontSize(),
                base.mainRunLogFilter(),
                base.mainRunLogLines(),
                base.mainRunLogScroll(),
                mainRunStage2ProductionPlan != null ? mainRunStage2ProductionPlan : "",
                mainRunStage2MemberSchedule != null ? mainRunStage2MemberSchedule : "",
                base.mainRunStage2WriteExcel(),
                base.mainRunStage2ResultBookFont(),
                base.uiEnvRows(),
                base.mainShellTabOrder(),
                base.mainShellTabLayout(),
                base.mainShellTabTitleAliases(),
                Map.copyOf(mergedInner),
                equipmentGanttGraphicZoomPercent,
                equipmentGanttDateColWidth,
                equipmentGanttMachineColWidth,
                equipmentGanttProcessColWidth,
                equipmentGanttBarFontFamily != null ? equipmentGanttBarFontFamily : "",
                equipmentGanttBarFontPercent,
                equipmentGanttRowHeightPercent,
                equipmentGanttHeaderHeightPercent,
                equipmentGanttSlotWidthPercent,
                equipmentGanttShiftWheelHScrollPercent,
                equipmentGanttPersonBadgeGapPx,
                equipmentGanttPersonBadgeBandVerticalOffsetPx,
                equipmentGanttGraphicDataFingerprint != null ? equipmentGanttGraphicDataFingerprint : "",
                equipmentGanttBadgeDragDeltas,
                equipmentGanttPersonBadgeDragAdjustEnabled,
                equipmentGanttPersonBadgeEnabled,
                equipmentGanttPersonBadgeWireEnabled,
                equipmentGanttPersonBadgeWireStrokeHex != null ? equipmentGanttPersonBadgeWireStrokeHex : "",
                equipmentGanttPersonBadgeWireWidthPx,
                equipmentGanttPersonBadgeWireDashStyleKey != null ? equipmentGanttPersonBadgeWireDashStyleKey : "",
                equipmentGanttPersonBadgeWireMaxLengthPx,
                equipmentGanttPersonBadgeFontFamily != null ? equipmentGanttPersonBadgeFontFamily : "",
                equipmentGanttPersonBadgeFontPercent,
                equipmentGanttPersonBadgeFillHex != null ? equipmentGanttPersonBadgeFillHex : "",
                equipmentGanttPersonBadgeTextHex != null ? equipmentGanttPersonBadgeTextHex : "",
                equipmentGanttPersonBadgeStrokeHex != null ? equipmentGanttPersonBadgeStrokeHex : "",
                equipmentGanttPersonBadgeStrokeWidth,
                equipmentGanttPersonBadgeCornerRadius,
                equipmentGanttPersonBadgePill,
                equipmentGanttPersonBadgeGlowColorHex != null ? equipmentGanttPersonBadgeGlowColorHex : "",
                equipmentGanttPersonBadgeGlowRadius,
                equipmentGanttPersonBadgeGlowSpread,
                equipmentGanttPersonBadgeOpacity,
                equipmentGanttPersonBadgeStylesByLabel,
                equipmentGanttPersonBadgeStylesByMemberKey,
                equipmentGanttPlanJsonPath != null ? equipmentGanttPlanJsonPath : "",
                base.stage1NetworkCacheBadgeLabel(),
                base.stage1NetworkCacheBadgeStyle(),
                base.mainShellTabOrganizerHeaderGlow(),
                base.mainShellTabOrganizerHeaderGlowStrength(),
                base.pushButtonDesignPrefs(),
                base.memoryMonitorEnabled(),
                base.memoryMonitorIntervalSec(),
                base.nextLaunchHeapMaxMiB());
    }
}
