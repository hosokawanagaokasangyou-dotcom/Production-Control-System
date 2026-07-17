# リモートサポート用ログ（remote_log）

**日付:** 2026-07-17  
**状態:** 実装完了

## 背景

特定工場 PC で段階2後に成果物が欠ける不具合など、リモートで実行ログを確認したい。  
サマリ Excel（`サマリ_AI配台.xlsx`）と同じ共有フォルダへ、操作者別にログを世代管理で残す。

## 要件（確定）

| 項目 | 内容 |
|------|------|
| ルート | `{サマリ Excel 親}/remote_log/{操作者名}/` |
| 内容 | 実行・ログタブ全文 + Python `execution_log.txt`（あれば） |
| 保持 | **3日**（古い世代フォルダを削除） |
| タイミング | 段階1／2／2.1 **終了時**（成功・失敗とも） |
| ユーザー | セッション操作者名（未選択時はスキップ） |
| 無効化 | `PM_AI_REMOTE_LOG=0/false/off` でオフ（空＝有効） |

## 世代フォルダ

```
remote_log/
  細川/
    20260717-104108_stage2/
      ui_run_log.txt
      execution_log.txt   # 存在時のみ
      meta.json
```

- フォルダ名: `yyyyMMdd-HHmmss_{stageId}`（`stage1` / `stage2` / `stage2.1`）
- `meta.json`: 工場・操作者・ホスト名・exitCode・時刻・パス要約（秘密情報なし）

## アーキテクチャ

```
MainShellController.completeStageRunOnFx
  └─ RemoteSupportLogArchive.archiveAfterStageAsync(...)
       ├─ AppPaths.resolveRemoteLogRoot(ui)
       ├─ OperatorUserPaths.sanitizeOperatorDirName
       ├─ ui_run_log.txt / execution_log.txt / meta.json 書込
       └─ 3日超の世代フォルダ削除
```

## 非対象

- 段階3 の自動アーカイブ
- 共有書き込み失敗時のリトライ UI
- ログ内容の暗号化
