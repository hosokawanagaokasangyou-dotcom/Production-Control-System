package jp.co.pm.ai.desktop.config;

import java.util.List;
import java.util.Map;

/**
 * Paths and fields restored on startup from {@link DesktopSessionStateStore}.
 *
 * @param planInputPath path field on 配台計画_タスク入力 tab
 * @param planInputSheet sheet name on the same tab
 * @param stage1PreviewPath Stage1 preview file path
 * @param stage1PreviewSheet Stage1 preview sheet name
 * @param excludeRulesPath PM_AI_EXCLUDE_RULES_JSON path (editor tab)
 * @param mainRunWorkbook task-input workbook field on run tab
 * @param mainRunPythonExe Python executable field
 * @param mainRunScriptDir code/python directory field
 * @param windowWidth last main window width ({@code 0} if unknown / use default scene size)
 * @param windowHeight last main window height ({@code 0} if unknown)
 * @param windowX last window X ({@link Double#NaN} if unknown / keep toolkit placement)
 * @param windowY last window Y ({@link Double#NaN} if unknown)
 * @param uiTheme persisted UI theme id ({@link DesktopTheme#storedId()}, empty defaults to light)
 * @param logFontFamily run-tab log font family name; empty means default family
 * @param logFontSize run-tab log size in points; {@code 0} means default size
 * @param mainRunLogFilter persisted run-tab log filter enum name ({@code ALL}, {@code ERRORS_ONLY}, ...); empty means ALL
 * @param mainRunLogLines last run-tab log lines (capped when saving)
 * @param mainRunLogScroll vertical scroll position as 0..1 proportion of the scroll bar; {@link Double#NaN} if unknown
 * @param mainRunStage2ProductionPlan last shown stage-2 production_plan xlsx path on run tab (empty if none)
 * @param mainRunStage2MemberSchedule last shown stage-2 member_schedule xlsx path on run tab (empty if none)
 * @param mainRunStage2WriteExcel whether stage-2 writes xlsx deliverables; when false only JSON (run tab)
 * @param mainRunStage2ResultBookFont stage-2 result Excel font family; empty with system default in UI means Python
 *     built-in default
 * @param uiEnvRows persisted 環境変数 tab rows (empty uses bootstrap defaults only)
 * @param mainShellTabOrder ordered {@link jp.co.pm.ai.desktop.MainShellTabId#key()} values for the main window
 *     tab strip; empty restores default FXML order
 * @param equipmentGanttGraphicZoomPercent 設備ガント・グラフィックタブの表示倍率（50〜200、0 は未保存として既定 100）
 * @param equipmentGanttDateColWidth 同タブ左・日付列の幅（px、0 は自動計測）
 * @param equipmentGanttMachineColWidth 同タブ左・機械名列の幅（px、0 は自動計測）
 * @param equipmentGanttProcessColWidth 同タブ左・工程名列の幅（px、0 は自動計測）
 * @param equipmentGanttBarFontFamily 同タブタイムライン・バー内ラベル用フォントファミリ（空はシステム既定）
 * @param equipmentGanttBarFontPercent バー内ラベル文字サイズ（50〜200、100＝既定、0 は未保存として既定 100）
 * @param equipmentGanttRowHeightPercent データ行の高さ調整（50〜200、0 は未保存として既定 100）
 * @param equipmentGanttHeaderHeightPercent 見出し行（日付・機械名・工程名・時刻軸）の高さ（50〜200、0 は未保存として既定 100）
 * @param equipmentGanttSlotWidthPercent 時刻スロット列幅の調整（50〜500、0 は未保存として既定 100）
 * @param equipmentGanttShiftWheelHScrollPercent Shift+ホイール横スクロールの感度（50〜1000、100＝従来相当、0 は未保存として既定 200）
 * @param equipmentGanttPersonBadgeEnabled 設備ガント・担当バッジ表示のオンオフ
 * @param equipmentGanttPersonBadgeFontFamily バッジ文字フォント（空は既定ファミリ）
 * @param equipmentGanttPersonBadgeFontPercent バッジ文字サイズ（行ラベル基準の%、0 は未保存として既定 85）
 * @param equipmentGanttPersonBadgeFillHex バッジ背景色（#RRGGBB）
 * @param equipmentGanttPersonBadgeTextHex バッジ文字色
 * @param equipmentGanttPersonBadgeStrokeHex バッジ枠色
 * @param equipmentGanttPersonBadgeStrokeWidth バッジ枠の太さ（px 相当）
 * @param equipmentGanttPersonBadgeCornerRadius 角丸（ピルでないとき）
 * @param equipmentGanttPersonBadgePill カプセル形状
 * @param equipmentGanttPersonBadgeGlowColorHex グロー（DropShadow）の色
 * @param equipmentGanttPersonBadgeGlowRadius グロー半径
 * @param equipmentGanttPersonBadgeGlowSpread DropShadow の spread（0〜1）
 * @param equipmentGanttPersonBadgeStylesByLabel バッジ表示文字のみの旧キー（後方互換・読込のみ参照し得る）
 * @param equipmentGanttPersonBadgeStylesByMemberKey skills メンバー名（正規化キー）ごとの見た目
 * @param stage1NetworkCacheBadgeLabel 段階1付近バッジの表示文言（ネットワークソースがキャッシュのとき）
 * @param stage1NetworkCacheBadgeStyle 同バッジの {@link PersonBadgeStyle}
 */
public record DesktopSessionState(
        String planInputPath,
        String planInputSheet,
        String stage1PreviewPath,
        String stage1PreviewSheet,
        String excludeRulesPath,
        String mainRunWorkbook,
        String mainRunPythonExe,
        String mainRunScriptDir,
        double windowWidth,
        double windowHeight,
        double windowX,
        double windowY,
        String uiTheme,
        String logFontFamily,
        double logFontSize,
        String mainRunLogFilter,
        List<String> mainRunLogLines,
        double mainRunLogScroll,
        String mainRunStage2ProductionPlan,
        String mainRunStage2MemberSchedule,
        boolean mainRunStage2WriteExcel,
        String mainRunStage2ResultBookFont,
        List<UiEnvRowSnapshot> uiEnvRows,
        List<String> mainShellTabOrder,
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
        boolean equipmentGanttPersonBadgeEnabled,
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
        Map<String, PersonBadgeStyle> equipmentGanttPersonBadgeStylesByLabel,
        Map<String, PersonBadgeStyle> equipmentGanttPersonBadgeStylesByMemberKey,
        String stage1NetworkCacheBadgeLabel,
        PersonBadgeStyle stage1NetworkCacheBadgeStyle) {

    public DesktopSessionState {
        equipmentGanttPersonBadgeStylesByLabel =
                equipmentGanttPersonBadgeStylesByLabel == null || equipmentGanttPersonBadgeStylesByLabel.isEmpty()
                        ? Map.of()
                        : Map.copyOf(equipmentGanttPersonBadgeStylesByLabel);
        equipmentGanttPersonBadgeStylesByMemberKey =
                equipmentGanttPersonBadgeStylesByMemberKey == null
                                || equipmentGanttPersonBadgeStylesByMemberKey.isEmpty()
                        ? Map.of()
                        : Map.copyOf(equipmentGanttPersonBadgeStylesByMemberKey);
    }

    /**
     * セッション値と {@link PersonBadgeStyle#defaultStyle()} をマージした実効スタイル。
     */
    public PersonBadgeStyle resolvedPersonBadgeStyle() {
        PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
        return new PersonBadgeStyle(
                nz(equipmentGanttPersonBadgeFontFamily(), d.fontFamily()),
                equipmentGanttPersonBadgeFontPercent() > 0 && equipmentGanttPersonBadgeFontPercent() <= 300
                        ? equipmentGanttPersonBadgeFontPercent()
                        : d.fontPercent(),
                nz(equipmentGanttPersonBadgeFillHex(), d.fillHex()),
                nz(equipmentGanttPersonBadgeTextHex(), d.textHex()),
                nz(equipmentGanttPersonBadgeStrokeHex(), d.strokeHex()),
                equipmentGanttPersonBadgeStrokeWidth() >= 0
                        ? equipmentGanttPersonBadgeStrokeWidth()
                        : d.strokeWidth(),
                equipmentGanttPersonBadgeCornerRadius() >= 0
                        ? equipmentGanttPersonBadgeCornerRadius()
                        : d.cornerRadius(),
                equipmentGanttPersonBadgePill(),
                nz(equipmentGanttPersonBadgeGlowColorHex(), d.glowColorHex()),
                equipmentGanttPersonBadgeGlowRadius() >= 0
                        ? equipmentGanttPersonBadgeGlowRadius()
                        : d.glowRadius(),
                equipmentGanttPersonBadgeGlowSpread() >= 0 && equipmentGanttPersonBadgeGlowSpread() <= 1
                        ? equipmentGanttPersonBadgeGlowSpread()
                        : d.glowSpread());
    }

    /**
     * 担当者キー（バッジに表示する文字列）に紐づくスタイル。未登録キーは {@link #resolvedPersonBadgeStyle()}。
     */
    public PersonBadgeStyle resolvedPersonBadgeStyleForLabel(String badgeLabel) {
        String k = PersonBadgeStyle.normalizeLabelKey(badgeLabel);
        if (!k.isEmpty()) {
            PersonBadgeStyle per = equipmentGanttPersonBadgeStylesByLabel().get(k);
            if (per != null) {
                return per;
            }
        }
        return resolvedPersonBadgeStyle();
    }

    private static String nz(String s, String def) {
        return s != null && !s.isBlank() ? s.strip() : def;
    }

    public static DesktopSessionState empty() {
        PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
        return new DesktopSessionState(
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                0d,
                0d,
                Double.NaN,
                Double.NaN,
                "",
                "",
                0d,
                "",
                List.of(),
                Double.NaN,
                "",
                "",
                true,
                "",
                List.of(),
                List.of(),
                0d,
                0d,
                0d,
                0d,
                "",
                0d,
                0d,
                0d,
                0d,
                0d,
                true,
                "",
                d.fontPercent(),
                d.fillHex(),
                d.textHex(),
                d.strokeHex(),
                d.strokeWidth(),
                d.cornerRadius(),
                d.pill(),
                d.glowColorHex(),
                d.glowRadius(),
                d.glowSpread(),
                Map.of(),
                Map.of(),
                "",
                PersonBadgeStyle.networkSourceCacheBadgeDefault());
    }
}
