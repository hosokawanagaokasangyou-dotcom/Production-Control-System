package jp.co.pm.ai.planning.stage2.output;

import java.util.List;

/**
 * Python {@code planning_core._core.default_result_task_sheet_column_order(0)} と同一の
 * 「結果_タスク一覧」既定列順（履歴列なし＝28 列）。
 *
 * <p>PassThrough 経路は配台コア未実行のためセル値は限定的だが、列集合・順序を Python 段階2と揃え
 * 計画ブック JSON のスキーマ比較・ゴールデンを進めやすくする。
 */
public final class Stage2ResultTaskListCanonicalHeaders {

    private Stage2ResultTaskListCanonicalHeaders() {}

    /**
     * 既定の結果_タスク一覧見出し（左から）。Python 3.14 / planning_core 2026-05 時点の実出力と一致。
     */
    public static final List<String> DEFAULT_ORDER_NO_HISTORY =
            List.of(
                    "ステータス",
                    "配台状況メモ",
                    "タスクID",
                    "工程名",
                    "機械名",
                    "加工速度",
                    "優先度",
                    "配台試行順番",
                    "必須OP(上書)",
                    "タスク効率",
                    "加工途中",
                    "特別指定あり",
                    "担当OP指定",
                    "回答納期",
                    "指定納期",
                    "計画基準納期",
                    "原反投入日",
                    "原反投入日_試行前",
                    "試行順パターン原反前倒し",
                    "紝期緊急",
                    "加工開始日",
                    "配台済_加工開始",
                    "配台済_加工終了",
                    "納期を満たすか？",
                    "累計加工量",
                    "残加工量",
                    "完了率(実行時点)",
                    "特別指定_AI");
}
