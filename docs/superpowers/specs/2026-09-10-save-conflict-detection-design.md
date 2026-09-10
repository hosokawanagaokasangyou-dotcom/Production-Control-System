# 保存前ファイル競合検知（明示保存画面横断）設計

日付: 2026-09-10  
状態: 実装済み（共通基盤＋明示保存画面へ適用）

## 目的

明示的な「保存」操作の直前に、ディスク上の正本（JSON / 関連 Excel 等）が **画面読込時から外部変更されていないか** を検知する。  
変更があれば業務語の差分要約付きダイアログを出し、ユーザーが再読込／上書き／キャンセルを選ぶ。

当初検討した「タブ表示中の常時ファイルロック」は採用しない。アプリ起動中ずっとロックされる状態を避け、保存時の楽観的競合検知に寄せる。

## 要件（確定）

| 項目 | 内容 |
|------|------|
| ロック | **常時 OS 排他ロックはしない** |
| 検知タイミング | 明示保存の直前 |
| 検知方法 | 読込時に記録した **内容 SHA-256** と、保存直前のファイル内容ハッシュを比較 |
| 関連ファイル | 正本と関連 Excel 等は **セット指紋**（いずれか不一致で競合） |
| 競合時 UI | 強調ヘッダ + スクロール可能な業務語要約。ボタン: **再読込して更新** / **強制上書き** / **キャンセル** |
| 再読込して更新 | UI をディスク内容で更新。未保存編集は破棄。baseline 更新 |
| 強制上書き | 画面内容で保存を続行。成功後に baseline 更新 |
| キャンセル | 保存中止。編集は保持 |
| 範囲 | 明示保存がある画面を横断（下記）。自動保存は対象外 |
| 実装方針 | 共通ガード + 画面別 `ConflictDiffSummarizer` プラグイン |

## 対象外

- `session-state.json` 等の debounce／自動保存
- 列順・バッジ・タブ配置など UI 微調整の自動永続
- 操作者ユーザー `factory-operator-users.bin` の即時永続（必要なら後続タスク）
- 保存処理そのものの短い書き込み排他（既存どおり可。常時ロックとは別）

## アーキテクチャ

```
[各 TabController / 保存入口]
   │ 読込成功時
   ▼
FileContentFingerprint.capture(path…)  →  メモリ上の baseline（パスごと SHA-256）
   │
   │ 明示保存
   ▼
SaveConflictGuard.beforeSave(paths, baseline, summarizer)
   │ 一致 → Proceed
   │ ハッシュ計算失敗 → Abort + エラー（上書き選択肢なし）
   │ 不一致
   ▼
ConflictDiffSummarizer.summarize(...)  ※画面別プラグイン
   │
   ▼
ConflictSaveDialog（強調ヘッダ + スクロール要約）
   ├── 再読込して更新 → reload UI, refresh baseline, Abort save
   ├── 強制上書き → Proceed save; 成功後 refresh baseline
   └── キャンセル → Abort save（編集保持）
```

### 共通コンポーネント（新規・想定パッケージ）

| クラス | 責務 |
|--------|------|
| `FileContentFingerprint` | ファイル内容の SHA-256。欠落は sentinel（例: `ABSENT`）。I/O 失敗は例外 |
| `SaveConflictGuard` | baseline と現ディスクの比較。複数パスのセット判定 |
| `ConflictDiffSummarizer` | interface。画面／ドメインごとに実装 |
| `ConflictSaveDialog` | JavaFX ダイアログ（案 B）。結果 enum: RELOAD / OVERWRITE / CANCEL |

競合の **有無** はハッシュだけで判定する。業務語要約のため、読込成功時に **比較用スナップショット**（パース済みモデル、または読込バイト列）も保持する。ハッシュだけ残して内容を捨てると Summarizer が作れない。

配置の目安: `code_java/.../jp/co/pm/ai/desktop/io/conflict/`（既存の慣習に合わせて調整可）。

### コントローラ側の契約

1. **読込成功時**（および再読込成功時）に対象パスの fingerprint を baseline として保持する。
2. **明示保存の入口**で、実書き込みの前に `SaveConflictGuard` を呼ぶ（非同期保存ならバックグラウンド開始前、または書込スレッド先頭で UI スレッドにダイアログを戻す）。
3. 保存成功後に baseline を更新する。失敗時は baseline を変えない。
4. 同一正本を複数タブが持つ場合（例: 会社カレンダーとメンバー勤怠 → `attendance-data.json`）は **タブごとに独立 baseline**。一方が保存成功すれば、他方は次回保存で競合 → 再読込を促す（意図どおり）。

## 対象画面とファイル

| 画面 | 正本 | 関連（セット指紋） | 保存入口（既存） |
|------|------|-------------------|------------------|
| 会社カレンダー | `attendance-data.json` | 勤怠・機械カレンダー.xlsx | `saveEditsAsync` |
| メンバー勤怠 | 同上 | 同上 | `saveEditsAsync` |
| 機械カレンダー | `machine-calendar-data.json` | 同上 xlsx | `saveEditsAsync` |
| 配台マスタ | `master-dispatch-sheets.json` | `master.xlsm` | `saveCurrentGridsToDisk` |
| 配台計画（タスク入力） | `PM_AI_PLAN_INPUT_PATH` | （パスが xlsx/xlsm ならそれ自体） | `onSaveButtonAction` |
| 配台不要ルール | exclude rules JSON | なし | `onSaveButtonAction` |
| 特別ルール JSON / ビルダー / 工程優先 | `dispatch_special_rules.json` | なし | 各 `onSave*` / `saveToDisk` |
| 材料・製品種類 | サマリ同フォルダ `.txt` 群 | なし（保存対象ファイルをセット） | 各 `saveToDisk` |
| 設備ガント担当 | 計画 JSON + 契約 JSON | 同期する計画 xlsx があれば含める | `beginSaveAssignmentChanges` |
| 依頼書入力（設定） | `request_form_input_settings.json` | — | 設定の明示保存 |
| 依頼書入力（受注転記） | 受注 Excel | — | 転記・更新入口 |
| 後加工商品マスタ編集 | アップロード用 xlsx | — | `saveUpload` |
| ユーザープロファイル | `user-profiles/*.json` | — | `onSaveAction` |
| グローバル設定パッケージ既定 | `init_setting/session_defaults_*.json` 等 | — | `onSavePackageDefaultsAction` |
| 環境変数・Gemini 資格情報 | 認証 JSON / encrypted | — | 暗号化保存入口（要約は秘匿） |

パス解決は既存の `AppPaths` / 各コントローラの解決ロジックに従う。

## 業務語差分（Summarizer）

各画面の Summarizer は、ディスク上の内容と「画面が読込時に持っていた内容（または baseline 取得時のスナップショット）」を比較し、**その画面の用語**で短く列挙する。

例（メンバー勤怠）:

- メンバーの追加・削除（氏名）
- 該当年月の出勤区分セル変更件数・代表例（氏名と日付）
- 関連 Excel のみ変化した場合はその旨

例（会社カレンダー）: 公休・特別休暇・出勤日の日付差分。  
例（Gemini 資格情報）: 中身は出さず「認証ファイルが更新／置換されています」程度。

要約生成に失敗した場合でもダイアログは表示し、要約欄は「詳細差分を生成できませんでした」＋ファイル名＋ハッシュ先頭数桁とする。

## ダイアログ UI（案 B）

- 強調ヘッダ（競合であることが一目で分かる色／文言）
- 「変更の要約」ラベル + スクロール可能な本文
- 注記: 再読込すると未保存編集は破棄される
- ボタン: **再読込して更新** / **強制上書き** / **キャンセル**
- 既存の勤怠未保存確認など他ダイアログと文言・トーンを揃える

## エラー・境界

| 状況 | 挙動 |
|------|------|
| 読込時ファイル無し | baseline = ABSENT。保存時にファイルが存在すれば競合 |
| 保存直前にファイル消失 | 競合（要約: ディスク上に無い）。上書きで新規作成可 |
| ハッシュ計算の I/O・権限失敗 | 保存中止 + エラー。上書き選択肢なし |
| 要約失敗 | ダイアログは表示。フォールバック文言 |
| 上書き後の書込失敗 | 既存の保存エラー処理。baseline 未更新 |

## テスト

- **単体**: `FileContentFingerprint`（同一／変更／欠落）、`SaveConflictGuard`（単一・セット）、代表 Summarizer 数種（勤怠・ルール JSON 等）
- **結合／手動**: 一時ファイルを外部改変して保存 → 再読込／上書き／キャンセルの分岐

## 完了条件

1. 上記の明示保存画面がすべて Guard 経由である
2. 各画面に業務語 Summarizer がある（Gemini は秘匿要約）
3. ダイアログは強調ヘッダ + スクロール要約
4. 常時ファイルロックを導入していない

## 非目標

- 自動マージやセル単位の競合解消 UI
- 他プロセスへのリアルタイム通知
- タブ選択中のファイルロックバッジ表示（方針変更により不要）

## 実装計画へのメモ

対象画面が多いため、実装計画では **共通部品 → 勤怠系 → その他明示保存画面** の順でタスク分割してよい。完了条件は全画面 Guard 経由のまま変えない。
