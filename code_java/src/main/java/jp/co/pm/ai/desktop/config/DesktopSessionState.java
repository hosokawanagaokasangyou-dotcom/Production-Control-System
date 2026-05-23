package jp.co.pm.ai.desktop.config;

import java.time.LocalDate;
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
 * @param mainRunScriptDir code/python directory field（編集は環境変数タブの {@code PM_AI_CODE_PYTHON_DIR}）
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
 * @param mainRunStage2SkipTodayDispatch when true, stage-2 skips dispatch on the data-extraction calendar day
 *     (UI checkbox is on 配台計画_タスク入力 tab; session key name unchanged)
 * @param planInputStage2InProgressNextDayPrompt when true, stage-2 shows the in-progress next-day dispatch dialog
 *     (配台計画_タスク入力 tab; default on)
 * @param mainRunStage2ResultBookFont stage-2 result Excel font family; empty with system default in UI means Python
 *     built-in default
 * @param mainRunSkipGeminiApi when true, skip Gemini generateContent calls (development; run tab checkbox)
 * @param mainRunStage1MarkAllExcludeAfterRun when true, after successful stage 1 mark all plan-input rows 配台不要=yes (development)
 * @param uiEnvRows persisted 環境変数 tab rows (empty uses bootstrap defaults only)
 * @param mainShellTabOrder ordered {@link jp.co.pm.ai.desktop.MainShellTabId#key()} values for the main window
 *     tab strip; empty restores default FXML order（{@link #mainShellTabLayout()} が空のときのみ有効）
 * @param mainShellTabLayout メインシェルタブの入れ子構成・色（空は未使用として従来のフラット＋{@link #mainShellTabOrder()}）
 * @param mainShellTabTitleAliases メイン作業タブ見出しの表示エイリアス（キーは {@link jp.co.pm.ai.desktop.MainShellTabId#key()}、空値は保存しない）
 * @param innerTabSelectedIndexByShellTabKey メインシェル直下の子 {@link javafx.scene.control.TabPane} の選択インデックス（キーは
 *     {@link jp.co.pm.ai.desktop.MainShellTabId#key()}。対応タブのみ）
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
 * @param equipmentGanttPrepTimeLabelsEnabled 設備ガント・準備時間系バーラベル（日次始業準備／依頼切替準備／休憩再開準備）の表示
 * @param equipmentGanttPersonBadgeGapPx 担当バッジの横方向の固定間隔（px、隣接ピル左端同士の追加距離、0〜48 程度を想定）
 * @param equipmentGanttPersonBadgeBandVerticalOffsetPx 担当バッジブロックをタスク帯に対して縦方向へずらす量（px、正で下方向）
 * @param equipmentGanttGraphicDataFingerprint 設備ガント表示データの内容フィンガープリント（SHA-256 16 進）。JSON 等が変わると無効化される
 * @param equipmentGanttBadgeDragDeltas データ同一時のみ有効な担当バッジのドラッグずれ（キーはバッジ安定 ID）
 * @param equipmentGanttPersonBadgeDragAdjustEnabled 担当バッジをマウスドラッグで移動するモード（データ同一ならずれはセッションに保存される）
 * @param equipmentGanttPersonBadgeEnabled 設備ガント・担当バッジ表示のオンオフ
 * @param equipmentGanttPersonBadgeWireEnabled 担当バッジとチャートバーをワイヤーで結ぶ（バッジ表示時のみ有効）
 * @param equipmentGanttPersonBadgeWireStrokeHex ワイヤー色（#RRGGBB / #RRGGBBAA、空はテーマのバー枠色＋既定不透明度）
 * @param equipmentGanttPersonBadgeWireWidthPx ワイヤー太さ（px、{@code 0} または非正はズームに応じた自動）
 * @param equipmentGanttPersonBadgeWireDashStyleKey 線種（{@code SOLID}|{@code DASHED}|{@code DOTTED}|{@code DASH_DOT}、空は SOLID）
 * @param equipmentGanttPersonBadgeWireMaxLengthPx ワイヤー表示時のバッジ中心とアンカー間の距離（px）。正の値は放射配置の半径かつドラッグ時の距離上限。{@code 0} は無制限（横並び初期配置）
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
 * @param equipmentGanttPersonBadgeOpacity バッジの不透明度（0〜1、{@code -1} は未保存として既定を使用）
 * @param equipmentGanttPersonBadgeStylesByLabel バッジ表示文字のみの旧キー（後方互換・読込のみ参照し得る）
 * @param equipmentGanttPersonBadgeStylesByMemberKey skills メンバー名（正規化キー）ごとの見た目
 * @param equipmentGanttPlanJsonPath 設備ガント・グラフィックタブの計画 JSON パス（空は未保存）
 * @param stage1NetworkCacheBadgeLabel 段階1付近バッジの表示文言（ネットワークソースがキャッシュのとき）
 * @param stage1NetworkCacheBadgeStyle 同バッジの {@link PersonBadgeStyle}
 * @param mainShellTabOrganizerHeaderGlow メインシェル「タブの並び」で指定した見出し色にグロー（dropshadow）を付けるか
 * @param mainShellTabOrganizerHeaderGlowStrength 見出しグローの強さ（0.0〜1.0、1.0 が従来既定の見え方）
 * @param pushButtonDesignPrefs プッシュボタン見た目のユーザー上書き
 * @param memoryMonitorEnabled メモリ設定タブのヒープ監視（トレンドグラフ）を有効にするか
 * @param memoryMonitorIntervalSec 監視間隔（秒、1〜3600）
 * @param nextLaunchHeapMaxMiB 次回 JVM 起動時に希望するヒープ上限（MiB、{@code 0} は未設定として UI で現在値を参照）
 * @param equipmentStatusDashboardActualDate ダッシュボード実績表示日（{@code yyyy-MM-dd}、空は起動時当日）
 * @param equipmentStatusDashboardPlanDate ダッシュボード予定表示日（{@code yyyy-MM-dd}、空は起動時当日）
 * @param equipmentStatusDashboardActualDayOffset 旧セッション互換（読込のみ・日付未保存時）
 * @param equipmentStatusDashboardPlanDayOffset 旧セッション互換（読込のみ・日付未保存時）
 * @param equipmentStatusDashboardAutoRefreshEnabled ダッシュボード自動更新（既定 ON）
 * @param equipmentStatusDashboardShowAladdinPlans ダッシュボードでアラジン予定を表示
 * @param equipmentStatusDashboardShowDispatchPlans ダッシュボードで配台予定を表示
 * @param equipmentStatusDashboardAppearance ダッシュボードカードの見た目設定
 */
public record DesktopSessionState(
        String planInputPath,
        String planInputSheet,
        String stage1PreviewPath,
        String stage1PreviewSheet,
        String excludeRulesPath,
        String mainRunWorkbook,
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
        boolean mainRunStage2SkipTodayDispatch,
        boolean planInputStage2InProgressNextDayPrompt,
        String mainRunStage2ResultBookFont,
        boolean mainRunSkipGeminiApi,
        boolean mainRunStage1MarkAllExcludeAfterRun,
        List<UiEnvRowSnapshot> uiEnvRows,
        List<String> mainShellTabOrder,
        List<MainShellTabLayoutNode> mainShellTabLayout,
        Map<String, String> mainShellTabTitleAliases,
        Map<String, Integer> innerTabSelectedIndexByShellTabKey,
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
        boolean equipmentGanttPrepTimeLabelsEnabled,
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
        String equipmentGanttPlanJsonPath,
        String stage1NetworkCacheBadgeLabel,
        PersonBadgeStyle stage1NetworkCacheBadgeStyle,
        boolean mainShellTabOrganizerHeaderGlow,
        double mainShellTabOrganizerHeaderGlowStrength,
        PushButtonDesignPrefs pushButtonDesignPrefs,
        boolean memoryMonitorEnabled,
        long memoryMonitorIntervalSec,
        long nextLaunchHeapMaxMiB,
        String equipmentStatusDashboardActualDate,
        String equipmentStatusDashboardPlanDate,
        int equipmentStatusDashboardActualDayOffset,
        int equipmentStatusDashboardPlanDayOffset,
        boolean equipmentStatusDashboardAutoRefreshEnabled,
        boolean equipmentStatusDashboardShowAladdinPlans,
        boolean equipmentStatusDashboardShowDispatchPlans,
        EquipmentStatusDashboardAppearancePrefs equipmentStatusDashboardAppearance) {

    /** 設備ガント・担当バッジ横方向固定間隔（px）の既定、およびスライダー上限の目安。 */
    public static final double DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX = 4.0;

    public static final double MAX_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX = 48.0;

    /** 帯に対するバッジブロックの縦オフセット（px）の既定。 */
    public static final double DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX = 0.0;

    /** 帯に対する縦オフセットのスライダー範囲（px）。 */
    public static final double MIN_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX = -48.0;

    public static final double MAX_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX = 48.0;

    /** 設備ガント・準備時間系バーラベル表示の既定（既定 OFF）。 */
    public static final boolean DEFAULT_EQUIPMENT_GANTT_PREP_TIME_LABELS_ENABLED = false;

    /** 設備ガント・バッジワイヤー表示の既定（視認性向上のため既定 ON）。 */
    public static final boolean DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_ENABLED = true;

    /** ワイヤー色が未指定のときテーマのバー枠色に乗せる不透明度。 */
    public static final double DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_THEME_OPACITY = 0.45;

    /** ワイヤー太さが {@code 0} のときの自動太さの係数（{@code max(下限, 係数 * zoom)}）。 */
    public static final double DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_AUTO_WIDTH_FACTOR = 0.65;

    public static final double DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_AUTO_WIDTH_MIN_PX = 0.75;

    /** ワイヤー太さの手動指定時の上限（px）。 */
    public static final double MAX_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_WIDTH_PX = 12.0;

    /** ワイヤー色・線種の既定（空はテーマ／SOLID）。 */
    public static final String DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_STROKE_HEX = "";

    /** {@code 0} は {@link #DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_AUTO_WIDTH_FACTOR} による自動太さ。 */
    public static final double DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_WIDTH_PX = 0d;

    public static final String DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_DASH_STYLE_KEY = "SOLID";

    /**
     * ワイヤー長の既定（px）。ワイヤー表示時はバッジ中心がアンカーからこの距離の円周上に置かれる（放射配置の半径）。
     * {@code 0} は無制限（従来の横並び初期配置＋距離クランプなし）。
     */
    public static final double DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX = 25d;

    /** UI スライダーおよび保存値のワイヤー長上限（px）。 */
    public static final double MAX_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX = 1200d;

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
        equipmentStatusDashboardAppearance =
                equipmentStatusDashboardAppearance != null
                        ? equipmentStatusDashboardAppearance
                        : EquipmentStatusDashboardAppearancePrefs.defaults();
        mainShellTabLayout =
                mainShellTabLayout == null || mainShellTabLayout.isEmpty()
                        ? List.of()
                        : List.copyOf(mainShellTabLayout);
        mainShellTabTitleAliases =
                mainShellTabTitleAliases == null || mainShellTabTitleAliases.isEmpty()
                        ? Map.of()
                        : Map.copyOf(mainShellTabTitleAliases);
        innerTabSelectedIndexByShellTabKey =
                innerTabSelectedIndexByShellTabKey == null || innerTabSelectedIndexByShellTabKey.isEmpty()
                        ? Map.of()
                        : Map.copyOf(innerTabSelectedIndexByShellTabKey);
        equipmentGanttGraphicDataFingerprint =
                equipmentGanttGraphicDataFingerprint != null
                        ? equipmentGanttGraphicDataFingerprint
                        : "";
        equipmentGanttBadgeDragDeltas =
                equipmentGanttBadgeDragDeltas == null || equipmentGanttBadgeDragDeltas.isEmpty()
                        ? Map.of()
                        : Map.copyOf(equipmentGanttBadgeDragDeltas);
        equipmentGanttPlanJsonPath =
                equipmentGanttPlanJsonPath != null ? equipmentGanttPlanJsonPath.strip() : "";
        equipmentGanttPersonBadgeWireStrokeHex =
                equipmentGanttPersonBadgeWireStrokeHex != null
                        ? equipmentGanttPersonBadgeWireStrokeHex.strip()
                        : "";
        equipmentGanttPersonBadgeWireDashStyleKey =
                equipmentGanttPersonBadgeWireDashStyleKey != null
                        ? equipmentGanttPersonBadgeWireDashStyleKey.strip()
                        : "";
        equipmentGanttPersonBadgeWireMaxLengthPx =
                Double.isFinite(equipmentGanttPersonBadgeWireMaxLengthPx)
                                && equipmentGanttPersonBadgeWireMaxLengthPx >= 0
                        ? Math.min(
                                equipmentGanttPersonBadgeWireMaxLengthPx,
                                MAX_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX)
                        : DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX;
        equipmentStatusDashboardActualDate =
                equipmentStatusDashboardActualDate != null
                        ? equipmentStatusDashboardActualDate.strip()
                        : "";
        equipmentStatusDashboardPlanDate =
                equipmentStatusDashboardPlanDate != null
                        ? equipmentStatusDashboardPlanDate.strip()
                        : "";
    }

    /** セッション保存日付（{@code yyyy-MM-dd}）または旧オフセットから実績表示日を復元。 */
    public LocalDate resolveEquipmentStatusDashboardActualDate(LocalDate anchorToday) {
        LocalDate parsed = parseIsoLocalDate(equipmentStatusDashboardActualDate);
        if (parsed != null) {
            return parsed;
        }
        if (anchorToday != null && equipmentStatusDashboardActualDayOffset != 0) {
            return anchorToday.plusDays(equipmentStatusDashboardActualDayOffset);
        }
        return anchorToday != null ? anchorToday : LocalDate.now();
    }

    /** セッション保存日付（{@code yyyy-MM-dd}）または旧オフセットから予定表示日を復元。 */
    public LocalDate resolveEquipmentStatusDashboardPlanDate(LocalDate anchorToday) {
        LocalDate parsed = parseIsoLocalDate(equipmentStatusDashboardPlanDate);
        if (parsed != null) {
            return parsed;
        }
        if (anchorToday != null && equipmentStatusDashboardPlanDayOffset != 0) {
            return anchorToday.plusDays(equipmentStatusDashboardPlanDayOffset);
        }
        return anchorToday != null ? anchorToday : LocalDate.now();
    }

    private static LocalDate parseIsoLocalDate(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        try {
            return LocalDate.parse(raw.strip());
        } catch (Exception ex) {
            return null;
        }
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
                        : d.glowSpread(),
                equipmentGanttPersonBadgeOpacity() >= 0.0 && equipmentGanttPersonBadgeOpacity() <= 1.0
                        ? equipmentGanttPersonBadgeOpacity()
                        : d.opacity());
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
                false,
                true,
                "",
                false,
                false,
                List.of(),
                List.of(),
                List.of(),
                Map.of(),
                Map.of(),
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
                DEFAULT_EQUIPMENT_GANTT_PREP_TIME_LABELS_ENABLED,
                DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX,
                DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX,
                "",
                Map.of(),
                false,
                true,
                DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_ENABLED,
                DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_STROKE_HEX,
                DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_WIDTH_PX,
                DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_DASH_STYLE_KEY,
                DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX,
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
                -1d,
                Map.of(),
                Map.of(),
                "",
                "",
                PersonBadgeStyle.networkSourceCacheBadgeDefault(),
                true,
                1d,
                PushButtonDesignPrefs.inactiveDefaults(),
                false,
                5L,
                0L,
                "",
                "",
                0,
                0,
                true,
                true,
                true,
                EquipmentStatusDashboardAppearancePrefs.defaults());
    }

    /**
     * 工場出荷 UI リセット用: {@code this} をパッケージ既定の UI とみなし、パス・環境タブ・主要実行パスだけ {@code bootstrap}
     * から上書きする。
     */
    public DesktopSessionState withBootstrapFieldsFrom(DesktopSessionState bootstrap) {
        return new DesktopSessionState(
                bootstrap.planInputPath(),
                bootstrap.planInputSheet(),
                bootstrap.stage1PreviewPath(),
                bootstrap.stage1PreviewSheet(),
                bootstrap.excludeRulesPath(),
                bootstrap.mainRunWorkbook(),
                bootstrap.mainRunScriptDir(),
                windowWidth(),
                windowHeight(),
                windowX(),
                windowY(),
                uiTheme(),
                logFontFamily(),
                logFontSize(),
                mainRunLogFilter(),
                mainRunLogLines(),
                mainRunLogScroll(),
                bootstrap.mainRunStage2ProductionPlan(),
                bootstrap.mainRunStage2MemberSchedule(),
                bootstrap.mainRunStage2SkipTodayDispatch(),
                bootstrap.planInputStage2InProgressNextDayPrompt(),
                bootstrap.mainRunStage2ResultBookFont(),
                bootstrap.mainRunSkipGeminiApi(),
                bootstrap.mainRunStage1MarkAllExcludeAfterRun(),
                bootstrap.uiEnvRows(),
                mainShellTabOrder(),
                mainShellTabLayout(),
                mainShellTabTitleAliases(),
                innerTabSelectedIndexByShellTabKey(),
                equipmentGanttGraphicZoomPercent(),
                equipmentGanttDateColWidth(),
                equipmentGanttMachineColWidth(),
                equipmentGanttProcessColWidth(),
                equipmentGanttBarFontFamily(),
                equipmentGanttBarFontPercent(),
                equipmentGanttRowHeightPercent(),
                equipmentGanttHeaderHeightPercent(),
                equipmentGanttSlotWidthPercent(),
                equipmentGanttShiftWheelHScrollPercent(),
                equipmentGanttPrepTimeLabelsEnabled(),
                equipmentGanttPersonBadgeGapPx(),
                equipmentGanttPersonBadgeBandVerticalOffsetPx(),
                equipmentGanttGraphicDataFingerprint(),
                equipmentGanttBadgeDragDeltas(),
                equipmentGanttPersonBadgeDragAdjustEnabled(),
                equipmentGanttPersonBadgeEnabled(),
                equipmentGanttPersonBadgeWireEnabled(),
                equipmentGanttPersonBadgeWireStrokeHex(),
                equipmentGanttPersonBadgeWireWidthPx(),
                equipmentGanttPersonBadgeWireDashStyleKey(),
                equipmentGanttPersonBadgeWireMaxLengthPx(),
                equipmentGanttPersonBadgeFontFamily(),
                equipmentGanttPersonBadgeFontPercent(),
                equipmentGanttPersonBadgeFillHex(),
                equipmentGanttPersonBadgeTextHex(),
                equipmentGanttPersonBadgeStrokeHex(),
                equipmentGanttPersonBadgeStrokeWidth(),
                equipmentGanttPersonBadgeCornerRadius(),
                equipmentGanttPersonBadgePill(),
                equipmentGanttPersonBadgeGlowColorHex(),
                equipmentGanttPersonBadgeGlowRadius(),
                equipmentGanttPersonBadgeGlowSpread(),
                equipmentGanttPersonBadgeOpacity(),
                equipmentGanttPersonBadgeStylesByLabel(),
                equipmentGanttPersonBadgeStylesByMemberKey(),
                bootstrap.equipmentGanttPlanJsonPath(),
                stage1NetworkCacheBadgeLabel(),
                stage1NetworkCacheBadgeStyle(),
                mainShellTabOrganizerHeaderGlow(),
                mainShellTabOrganizerHeaderGlowStrength(),
                pushButtonDesignPrefs(),
                memoryMonitorEnabled(),
                memoryMonitorIntervalSec(),
                nextLaunchHeapMaxMiB(),
                equipmentStatusDashboardActualDate(),
                equipmentStatusDashboardPlanDate(),
                equipmentStatusDashboardActualDayOffset(),
                equipmentStatusDashboardPlanDayOffset(),
                equipmentStatusDashboardAutoRefreshEnabled(),
                equipmentStatusDashboardShowAladdinPlans(),
                equipmentStatusDashboardShowDispatchPlans(),
                equipmentStatusDashboardAppearance());
    }
}
