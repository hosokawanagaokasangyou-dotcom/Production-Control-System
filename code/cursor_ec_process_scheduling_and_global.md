# EC process scheduling and global comments
_Exported on 2026/4/3 at 19:02:06 JST from Cursor (3.0.6)_

---

**User**

メイン_グローバル_未適用メモ(AI)	2026/4/4のEC工程には森下と宮島を配台してください。
メイン_再優先特別記載(原文)	2026/4/4 工程名:EC 森下と宮島を配台
メイン_再優先特別記載(AI)	{"ignore_skill_requirements": false, "ignore_need_minimum": false, "abolish_all_scheduling_limits": false, "task_preferred_operators": {}, "interpretation_ja": "2026年4月4日のEC工程に森下と宮島を配台する指示。", "factory_closure_dates": [], "scheduler_notes_ja": "2026/4/4のEC工程には森下と宮島を配台してください。", "global_speed_rules": [], "task_planning_basis_dates": {}}

グローバルコメントの文言が反映されていない理由は？

---

**Cursor**


グローバルコメントと `scheduler_notes_ja` の読み込み・反映経路をコードベースで調べます。




結論から言うと、**グローバルコメントの文言はログ・解析ブロックには反映されている一方、配台ロジックが読むのは「構造化されたJSONの一部」だけ**だからです。

## 1. 「未適用メモ(AI)」の意味

`メイン_グローバル_未適用メモ(AI)` は、内部では **`scheduler_notes_ja`** をそのまま載せています。

```12621:12623:m:\国分工場\国分共有\工程管理AIプロジェクト\----AI------\細川\GoogleAIStudio\テストコード\python\planning_core.py
    _sn = str(global_priority_override.get("scheduler_notes_ja") or "").strip()
    if _sn:
        ai_log_data["メイン_グローバル_未適用メモ(AI)"] = _sn[:2000]
```

設計上、`scheduler_keywords_ja` は **「上記キーに落としきれない補足」** と定義されています（＝**システムが機械適用しない／しきれなかった内容の置き場**）。

```2854:2855:m:\国分工場\国分共有\工程管理AIプロジェクト\----AI------\細川\GoogleAIStudio\テストコード\python\planning_core.py
    - task_planning_basis_dates: 依頼NO→計画基準納期（YYYY-MM-DD）。配台の納期基準として **特別指定_備考AIより優先**して適用する。該当なしは {{}}。
    - scheduler_notes_ja: 上記に落としきれない補足や運用メモ（速度は可能なら global_speed_rules も併記）。
```

表示ラベルが「未適用メモ」なので、**「コメントが無視された」のではなく、「自動配台には使わない欄に分類された」** と読むのが正確です。

## 2. 森下・宮島の指名が配台に効かない理由

実際に担当OPを上書きするのは **`task_preferred_operators`（依頼NO → 主担当）** だけです。

```6005:6010:m:\国分工場\国分共有\工程管理AIプロジェクト\----AI------\細川\GoogleAIStudio\テストコード\python\planning_core.py
        gpo = global_priority_override or {}
        gop_name = _global_override_preferred_operator_for_task(
            gpo.get("task_preferred_operators"), task_id
        )
        if gop_name is not None:
            preferred_operator_raw = gop_name
```

提示の JSON では **`task_preferred_operators` が `{}`** なので、**グローバルからの OP 指名は一度も走りません**。  
「日付＋工程名（EC）＋複数名」は、AI が **依頼NO単位の1名指名** に落とし切れず、`scheduler_notes_ja` に逃がした、という状態です。

`scheduler_notes_ja` を配台が読む処理はなく（ログ・配台シートの解析表示・速度ルールの限定的な補完など以外）、**この文字列だけでは誰をどのタスクに付けるかは決まりません**。

## 3. まとめ

| 期待 | 実際 |
|------|------|
| グローバルコメントが「記録」として残る | `メイン_再優先特別記載(原文)` と `メイン_再優先特別記載(AI)`、`未適用メモ(AI)` に残る（ここは**反映されている**） |
| 文言どおりに配台結果が変わる | **変わるのは** `factory_closure_dates`・各種ブール・`task_preferred_operators`・`global_speed_rules`・`task_planning_basis_dates` など**構造化フィールド**だけ。今回の主担当指名は `{}` なので**効かない** |

運用上は、**対象の依頼NOが分かるなら**グローバルコメントを「依頼NOごとの指名」表現に寄せる、または**タスク行の担当OP／特別指定**で指名する、といった形で **`task_preferred_operators` が埋まる**ようにすると配台に乗ります。

---

**User**

今回のグローバルコメントは、
2026/4/4のEC工程には森下と宮島を配台してください。

配台に適用する「日」
配台に適用する「工程名」
配台に適用する「メンバー名」
がはいっています

適用できるようになりませんか？


---

**Cursor**


日付・工程名・複数メンバーからグローバルコメントを配台に適用する拡張を実装します。まず `git status` を確認し、既存の担当指名・チーム割当の流れを調べます。




実装します。`global_day_process_operator_rules` を JSON に追加し、日付・工程名・複数メンバーをチームに必須含めします。まず `_resolve_preferred_op_to_member` を共通化してから、`_assign_one_roll_trial_order_flow` とレガシーループの両方に組み込みます。




共通化関数と `global_day_process_operator_rules` の coerce／適用を追加します。




## 対応内容

**グローバルコメント**を Gemini が **`global_day_process_operator_rules`** に落とすと、次が配台に効くようにしました。

- **`date`** … その**稼働日**にロールを切るときだけ適用（`current_date` と一致）
- **`process_contains`** … タスクの**工程名**（`machine` 列相当）への**部分一致**（例: `EC`）
- **`operator_names`** … 列挙した全員を**同一チームに必ず含める**（名前解決は従来の**担当OP指名**と同じく、当日の **OP/AS 候補**へ照合）

**試行順優先フロー**（`_assign_one_roll_trial_order_flow`）と**従来の組合せ探索**の両方で、プリセット／候補チームをこの拘束で絞り込みます。必須人数は **`req_num` を必須メンバー数以上**に繰り上げます。同一日・連続ロールの**継続チーム短絡**も、必須メンバーが揃っていないと使われません。

実装・プロンプト・配台シート右端の解析ブロックに **「日付×工程チーム指名」** 行を追加し、`配台ルール.md` / `配台ルール.html` も更新済みです。

---

## 使い方の注意

1. **依頼NOが分かる1名の主担当**は従来どおり **`task_preferred_operators`**。  
   **「◯月◯日の△工程に A と B」** のような **日付＋工程＋複数名** は、AI に **`global_day_process_operator_rules`** へ出させる想定です（プロンプトに手順を追記済み）。

2. **キャッシュ** … 既存の Gemini キャッシュだけだと新キーが無い JSON になりルールが空になるため、**キャッシュ接頭辞を `GLOBAL_PRIO_v8` に上げ**、次回実行で**再解釈**が走るようにしました。

3. **指名が当日スキル候補に解決できない**場合は **警告ログ**のみで、その名は必須から外れます（従来の担当OP指名と同じ扱い）。

---

## Git

- **be856ff** … `planning_core.py`／`配台ルール.md`／`配台ルール.html`（本機能）
- **623542f** … `GLOBAL_PRIORITY_OVERRIDE_CACHE_PREFIX` を v8 に変更（キャッシュ無効化）

**`git push` は実行していません。**

---

**User**

git comit& remote push実行

---

**Cursor**


作業ツリーを確認し、未コミットがあればコミットしてから `git push` します。




未コミットの変更が複数あります。指示どおりすべてステージしてコミットし、その後リモートへプッシュします。




作業を反映しました。

## コミット
- **ハッシュ**: `e6e685a`
- **メッセージ**: `git comit& remote push実行`
- **内容**: `git add -A` でステージした結果、**8 ファイル**（`.cursorignore`、`gemini_usage_cumulative.json`、`ai_remarks_cache.json`、`_verify_master_structure.py`、`配台ロジック詳細フロー.html`、`----AI------.code-workspace` ほか）がコミットに含まれました。  
  （初回 `git status` にあった `master.xlsm` / `.xlsm` / `__pycache__` などは、このリポジトリでは未追跡または無視のためコミット対象になっていません。）

## リモートプッシュ
- **リモート**: `https://github.com/kamoshika9999/----AI------.git`
- **ブランチ**: `xlwings積極適用`
- **結果**: 成功（`e2d959a..e6e685a` をプッシュ）
- **作業ツリー**: 現在はクリーンです。

