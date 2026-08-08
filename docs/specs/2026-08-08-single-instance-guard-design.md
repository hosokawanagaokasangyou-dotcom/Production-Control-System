# PmAiFxApp 二重起動抑制（SingleInstanceGuard）

**日付:** 2026-08-08  
**状態:** 実装完了

## 背景

工程管理 AI デスクトップ（`PmAiFxApp`）は現状、複数 JVM を並行起動できる。セッション・ロック・配台実行の競合を避けるため、同一マシンでは本体を 1 インスタンスに制限する。

## 要件（確定）

| 項目 | 内容 |
|------|------|
| 対象 | `PmAiFxApp` のみ |
| 非対象 | `RemoteDesktopFxApp` / `ReconciliationApp` / GPU プローブ子 JVM |
| 2つ目起動時 | 既存主窓を前面化し、ダイアログなしで即終了（exit 0） |
| 無効化 | `-Dpm.ai.singleInstance=false` |
| ポート上書き | `-Dpm.ai.singleInstance.port=…`（既定 `47821`） |

## 方式

ローカルソケット（`127.0.0.1` のみ）。

1. **Primary**: 固定ポートで listen。受理後に 1 行プロトコル `ACTIVATE\n` を読み、`OK\n` を返し、FX スレッドで主 `Stage` を前面化（`setIconified(false)` → `toFront()` → `requestFocus()`）。
2. **Secondary**: 同ポートへ接続し `ACTIVATE` を送る。成功したら `System.exit(0)`。
3. 呼び出し点: `PmAiFxApp.main` の早期（headless 判定の直後〜`configurePrismAfterProbe` の前）。プローブ JVM を無駄に立てない。

## エラー方針

| 状況 | 動作 |
|------|------|
| 接続成功＋応答 | Secondary として終了 |
| 接続失敗／短タイムアウト（例 300ms） | Primary 候補として bind 試行 |
| bind 失敗（他プロセス占有など） | ガードを諦めて通常起動。`StartupCrashLog` に警告 |
| アプリ終了 | `ServerSocket` を閉じる（shutdown hook または主窓 close） |

## 構成

- 新規: `jp.co.pm.ai.desktop.runtime.SingleInstanceGuard`（実装時に既存 `runtime` パッケージへ配置）
- 変更: `PmAiFxApp`（main 早期フック、主 Stage 登録、終了時解放）

## テスト

**単体（ヘッドレス）**

- Primary listen 中にクライアントが `ACTIVATE` → コールバック 1 回
- `pm.ai.singleInstance=false` でガード無効
- ポート上書きが効く

**手動**

- 通常二重起動 → 既存窓前面＋2つ目即終了
- `-Dpm.ai.singleInstance=false` で二重起動可

## 非対象（明示）

- Alert ダイアログ
- 環境変数タブ／TSV へのキー追加（システムプロパティのみ）
- 他 JavaFX 入口アプリの単一化
