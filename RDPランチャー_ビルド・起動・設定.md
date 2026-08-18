# RDPランチャー — ビルド・起動・設定

リモートデスクトップ専用ポータブル（`PmAiRpaLuncher.exe`）と、接続先 PC 向け C# ランチャー（`PmAiRdpRemoteLauncher.exe`）の **ビルド・起動・設定** をまとめる。

**関連ドキュメント（設計・データフロー詳細）**: [`リモートデスクトップ・ランチャー整理.md`](リモートデスクトップ・ランチャー整理.md)

**実装の正本**: `AppPaths.java` / `FactoryOperatorUserStore.java` / `RequestFormRemoteDesktopPane.java` / `tools/pm-ai-rdp-remote-launcher/`

---

## 概要（2 系統のポータブル）

| 製品 | エントリ | ローカルビルド出力 | 共有正本（版アップ） | 用途 |
|------|----------|-------------------|---------------------|------|
| **配台システム（PMD）** | `PMD.exe` → `PmAiFxApp` | `pm-ai-package-release\` | `\\192.168.0.101\共有フォルダ\湖南工場\…\pm-ai-package-release` | 配台・多数タブ（リモートデスクトップタブ含む） |
| **RDP ランチャー専用** | `PmAiRpaLuncher.exe` → `RemoteDesktopFxApp` | `rpa_luncher_release\` | `\\192.168.0.101\共有フォルダ\掲示板\rpa_luncher\RDPランチャ` | リモートデスクトップ UI のみ |

**第 3 の exe（ポータブルではない）**

| exe | 動く場所 | 役割 |
|-----|----------|------|
| `PmAiRdpRemoteLauncher.exe` | RDP **接続先 PC** | ログオン後に ini を読み、アラジン RPA を起動 |

旧名 `PmAiRdpLauncher.exe` は配台 PMD と混同しやすいため **`PmAiRpaLuncher.exe`（操作者 PC）** / **`PmAiRdpRemoteLauncher.exe`（接続先 PC）** に分離済み。

### 主な違い（PMD vs RDP 専用）

| 項目 | 配台 PMD | RDP 専用（PmAiRpaLuncher） |
|------|----------|---------------------------|
| PC ローカル設定 | `%USERPROFILE%\.pm-ai-desktop` | `%USERPROFILE%\.pm-ai-desktop-rdp` |
| 操作者 bin | `factory-operator-users.bin`（工場別） | `rdp-launcher-operator-users.bin`（`DATA\`） |
| 操作者スコープ | KONAN / KOKUBU | `FactorySite.RDP_LAUNCHER`（部署別） |

---

## ビルド方法

### 前提

- **Windows 上で PowerShell 5.1+** から実行（OpenJFX は win 分類器、jpackage は Windows JDK）。
- WSL からは `powershell.exe` 経由で同じスクリプトを呼ぶ。
- 初回は Temurin JDK zip・JavaFX・Maven 依存のダウンロードがあり時間がかかる。

### エントリポイント

| スクリプト | 説明 |
|-----------|------|
| `fast_package_rdp_launcher.ps1` | リポジトリ直下から呼ぶ **推奨エントリ**（内部で `code_java\package_rdp_launcher_app.ps1` を実行） |
| `code_java\package_rdp_launcher_app.ps1` | 実処理（Maven → jpackage → release 反映 → ZIP → 共有デプロイ） |

### WSL からのビルド（コピペ用）

**フルビルド（初回・依存更新後）**

```bash
cd /mnt/c/工程管理AIプロジェクト_JAVA
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./fast_package_rdp_launcher.ps1
```

**キャッシュ済みの再ビルド（JDK / JavaFX / C# スタブをスキップ）**

```bash
cd /mnt/c/工程管理AIプロジェクト_JAVA
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./fast_package_rdp_launcher.ps1 \
  -SkipJdkPrepare -SkipJavaFxPrepare -SkipCsLauncherBuild
```

**共有フォルダへの自動コピーをスキップ（オフライン / WSL で UNC 不可のとき）**

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./fast_package_rdp_launcher.ps1 -SkipCanonicalDeploy
```

**正本フォルダを上書き指定**

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./fast_package_rdp_launcher.ps1 \
  -CanonicalDeployDir '\\192.168.0.101\共有フォルダ\掲示板\rpa_luncher\RDPランチャ'
```

環境変数 `PM_AI_RDP_CANONICAL_DEPLOY_DIR` でも同様に指定可能。

### ビルドで行われること

1. （省略可）`scripts/build-rdp-remote-launcher.ps1` … `PmAiRdpRemoteLauncher.exe`
2. （省略可）`scripts/build-rdp-desktop-launcher.ps1` … `PmAiRpaLuncher.exe`（Unicode パス対応 C# スタブ）
3. `mvnw.cmd clean package -DskipTests`
4. jpackage `app-image` … `code_java\dist\PmAiRpaLuncher\`
5. `launcher-deploy-seed\` … 接続先向け exe 等を同梱
6. `rpa_luncher_release\PmAiRpaLuncher\` へミラー
7. ZIP 生成・（既定）共有 `RDPランチャ` へ `version.txt` / ZIP / `build-manifest.json` をコピー

### 出力物

| パス | 内容 |
|------|------|
| `rpa_luncher_release\PmAiRpaLuncher\PmAiRpaLuncher.exe` | 操作者 PC 用ポータブル（**ダブルクリック起動**） |
| `rpa_luncher_release\PmAiRpaLuncher_portable.zip` | フォルダ一式 ZIP |
| `rpa_luncher_release\PmAiRpaLuncher_version_upgrade.zip` | 版アップ用 ZIP |
| `rpa_luncher_release\version.txt` / `build-manifest.json` | 版比較・整合性 |
| `code_java\dist\PmAiRpaLuncher\launcher-deploy-seed\` | 接続先共有へコピーする種 |

ポータブルフォルダ内の `launch-pm-ai-rpa-launcher.bat` は、Defender が `runtime\bin\java.exe` を削除した場合の **フォールバック**。

### 所要時間の目安

| 条件 | 目安 |
|------|------|
| 初回（JDK + JavaFX ダウンロード + Maven + jpackage + C# ×2） | **15〜30 分**（回線・PC 性能依存） |
| キャッシュあり（`-SkipJdkPrepare -SkipJavaFxPrepare -SkipCsLauncherBuild`） | **3〜8 分** |
| ソース小変更のみ（上記スキップ + Maven 増分が効く場合） | **2〜5 分** |

配台 PMD のビルドは別系統: `fast_package_app.ps1` → `pm-ai-package-release\`（本ドキュメントの対象外）。

---

## 起動方法

### 本番（操作者 PC）— ポータブル

1. 共有またはローカルに展開した **`PmAiRpaLuncher.exe` があるフォルダ** を開く。
2. **`PmAiRpaLuncher.exe` をダブルクリック**（`app\` / `runtime\` と **同じフォルダに exe を置いたまま** 運用する）。
3. 起動時: **部署選択** → **操作者 + PIN**（ゲストは RDP 接続不可）。
4. リモートデスクトップタブでプロファイル選択 → **接続**（内部で ini 更新・資格情報 JSON 同期 → `mstsc`）。

版アップ: 環境変数 `PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR` の正本（既定 `…\rpa_luncher\RDPランチャ`）とローカル `version.txt` を比較し、新しければ `PmAiRpaLuncher_version_upgrade.zip` を適用。

### 開発（操作者 PC）— Maven exec

WSL から:

```bash
cd /mnt/c/工程管理AIプロジェクト_JAVA/code_java
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./run-pm-ai-remote-desktop.ps1
```

ヒープ変更例:

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./run-pm-ai-remote-desktop.ps1 -MaxHeap 3g
```

Windows コンソールから `code_java` で:

```powershell
.\run-pm-ai-remote-desktop.ps1
```

### 接続先 PC — PmAiRdpRemoteLauncher

接続先 Windows に **`PmAiRdpRemoteLauncher.exe`** を配置（運用上の正本は共有 `rpa_luncher\` 直下。詳細は後述）。

**手動起動例（操作者 細川）**

```text
\\192.168.0.101\共有フォルダ\掲示板\rpa_luncher\PmAiRdpRemoteLauncher.exe 細川
```

→ 同階層の `細川_RPA設定.ini` を参照。

**本番運用**: ログオン後 **タスクスケジューラ** で上記と同様に **操作者名を第 1 引数** に含める。

`起動プログラム番号=0` のときは RPA 起動・サインアウトともに行わない（二重起動抑止）。

---

## 設定方法

### 環境変数（RDP 専用シェルの環境変数タブ）

`RemoteDesktopEnvRows` が載せる主なキー。空欄時の既定は `RemoteDesktopEnvRows.applyRdpLauncherEmptyDefaults` / `ui_ref_env_defaults.json` に準拠。

| キー | 用途 | 既定（空のとき） |
|------|------|------------------|
| `PM_AI_REQUEST_FORM_RDP_PROFILE` | 接続用 `.rdp` | Windows 既定 `Default.rdp` を探索 |
| `PM_AI_RDP_LAUNCHER_DEPLOY_DIR` | 接続先 exe・ini・起動プロファイル JSON の配備先 | `\\192.168.0.101\…\rpa_luncher\RDPランチャ` ※ |
| `PM_AI_RDP_LAUNCHER_EXE` | ランチャー exe フルパス上書き | 配備先 + `PmAiRdpRemoteLauncher.exe` |
| `PM_AI_RDP_LAUNCHER_INI` | ini フルパス上書き | 配備先 + `{操作者}_RPA設定.ini` |
| `PM_AI_RDP_PORTABLE_BUNDLE_SOURCE_DIR` | ポータブル版アップ正本 | `…\rpa_luncher\RDPランチャ` |
| `PM_AI_RDP_OPERATOR_USERS_STORE_DIR` | 操作者 bin フォルダ | `…\rpa_luncher\DATA` |
| `PM_AI_RDP_LAUNCH_PROFILE_NUMBER` | 最後に使った起動プロファイル番号 1〜9 | 1 |
| `PM_AI_RDP_LAUNCHER_AUTO_DEPLOY` | UI から接続先へ exe 自動再配備 | 有効（`0`/`false`/`off` で無効） |
| `PM_AI_RDP_FULLSCREEN` / `WIDTH` / `HEIGHT` | RDP 全画面・解像度 | ウィンドウ 1920×1080 |
| `PM_AI_RDP_EMBED_STARTUP_IN_PROFILE` | `.rdp` へ alternate shell 書込 | off（既定はタスクスケジューラ + ini） |
| `PM_AI_OPERATOR_USER` | セッション操作者名（子プロセス env） | 起動時ダイアログで選択 |
| `PM_AI_FACTORY_SITE` | C# が JSON 内の工場ブロックを選ぶキー | `KONAN` |

※ **コード上の既定** `PM_AI_RDP_LAUNCHER_DEPLOY_DIR` は `RDPランチャ` サブフォルダ。**運用上 exe / ini を `rpa_luncher\` 直下に置く場合**は、環境変数タブで配備先を **`\\192.168.0.101\共有フォルダ\掲示板\rpa_luncher`** に明示設定する。

設定の保存先（PC ローカル）: `%USERPROFILE%\.pm-ai-desktop-rdp\session-state.json`

### RPA設定.ini（操作者別）

| 条件 | ファイル |
|------|----------|
| 操作者 **細川** | `{配備先}\細川_RPA設定.ini` |
| 操作者未設定 | `{配備先}\RPA設定.ini` |
| `PM_AI_RDP_LAUNCHER_INI` 指定 | そのフルパス |

**配備先** = `PM_AI_RDP_LAUNCHER_DEPLOY_DIR`（未設定時は上表の既定）。

レガシー読取フォールバック: `RAP設定.ini`、`DATA\` 配下の同名ファイル（**新規は `RPA設定.ini` を正**）。

**Java 側（接続直前）** が ini に書き込む主な項目:

- `起動プログラム番号`（1〜9、接続中は 0 でスケジューラ抑止）
- 選択スロットの RPA exe / `--scenario`
- `操作者=`（セッション操作者名）

**C# 側（接続先）** は ini を読み、RPA 起動後に `起動プログラム番号=0` へ戻す。

### 資格情報の流れ（bin → JSON → C#）

```
rdp-launcher-operator-users.bin（正本・DATA\）
        ↓ 保存 / 接続直前 sync
operator-aladdin-credentials.launcher.json（ini と同じフォルダ）
        ↓ ini の 操作者= + PM_AI_FACTORY_SITE
PmAiRdpRemoteLauncher → Aladdin RPA (--id / --password)
```

| データ | 場所 | 編集 UI |
|--------|------|---------|
| **操作者・PIN・アラジン ID/PW（正本）** | `DATA\rdp-launcher-operator-users.bin` | ユーザー管理者タブ |
| **C# 向けキャッシュ JSON** | `{配備先}\operator-aladdin-credentials.launcher.json` | Java が自動生成（手編集不要） |
| **接続時の操作者名** | ini の `操作者=` 行 | 接続直前に Java が更新 |

**「資格情報を保存」ボタン**（リモートデスクトップタブ）:

1. 起動時に操作者を選択済みであること。
2. アラジン loginId / password を bin に保存。
3. `syncLauncherCredentialsJsonToDeployDir` で JSON を ini と同じ親フォルダへ書き出す。

#### KONAN と RDP_LAUNCHER（トラブル多発ポイント）

JSON 構造:

```json
{
  "schemaVersion": 1,
  "factories": {
    "KONAN": { "操作者名": { "loginId": "...", "password": { ... } } },
    "RDP_LAUNCHER": { "操作者名": { "loginId": "...", "password": { ... } } }
  }
}
```

| 側 | 挙動 |
|----|------|
| **Java（PmAiRpaLuncher）** | 部署別操作者を **`RDP_LAUNCHER` ブロック** に書く（PMD 用 `KONAN`/`KOKUBU` とは別） |
| **C#（PmAiRdpRemoteLauncher）** | 環境変数 `PM_AI_FACTORY_SITE`（未設定時 **`KONAN`**）のブロックを先に参照 |
| **C# フォールバック** | 主キーに無い場合 **`RDP_LAUNCHER` ブロック** から解決（ログに明示） |

**接続先 PC のタスクスケジューラ** で `PM_AI_FACTORY_SITE=KONAN` のまま運用し、JSON が `RDP_LAUNCHER` のみ更新されていると資格情報未設定になる。**対処**: 接続先の env を未設定（KONAN ブロックにデータを置く）か `RDP_LAUNCHER` に合わせる、または Java 側で両ブロックに同期された JSON を維持する。

### 操作者 bin（ユーザー管理者）

| 項目 | 値 |
|------|-----|
| ファイル名 | `rdp-launcher-operator-users.bin` |
| 既定パス | `\\192.168.0.101\共有フォルダ\掲示板\rpa_luncher\DATA\` |
| バックアップ | 同 `DATA\rdp-launcher-operator-users-backups\` |
| 前回操作者 | PC ローカル `%USERPROFILE%\.pm-ai-desktop-rdp\last-rdp-launcher-operator.txt`（部署別ファイルあり） |

配台 PMD の `factory-operator-users.bin` とは **別ファイル**。RDP 専用アプリは部署単位で操作者を管理する。

### 更新（ポータブル版）

1. 開発 PC で `fast_package_rdp_launcher.ps1` を実行。
2. 正本 `RDPランチャ\` に `version.txt`・`PmAiRpaLuncher_version_upgrade.zip`・`build-manifest.json` が配置される（`-SkipCanonicalDeploy` 時は手動コピー）。
3. 各操作者 PC の `PmAiRpaLuncher` が起動時に版比較し、新しければ ZIP を適用（同梱 `rdp-apply-portable-update.ps1`）。

接続先 exe の更新: `launcher-deploy-seed\` から `rpa_luncher\` 直下へコピー、または UI の自動再配備（`PM_AI_RDP_LAUNCHER_AUTO_DEPLOY`）。

---

## 正本フォルダとファイル配置

運用上の UNC 正本（掲示板共有）:

```
\\192.168.0.101\共有フォルダ\掲示板\rpa_luncher\
│
├── PmAiRdpRemoteLauncher.exe          … 接続先ランチャー（接続先 PC / タスクスケジューラ）
├── PmAiRdpRemoteLauncher.version.txt
├── RPA設定.ini                        … 操作者未指定時
├── 細川_RPA設定.ini                   … 操作者別（例）
├── RDP起動プロファイル.json           … プロファイル名称・説明（Java UI 用）
├── operator-aladdin-credentials.launcher.json … C# 向け資格情報キャッシュ
├── launcher-logs\                     … 接続先ランチャー日次ログ（共有ミラー）
│   └── launcher-yyyyMMdd.log
│
├── DATA\
│   ├── rdp-launcher-operator-users.bin
│   └── rdp-launcher-operator-users-backups\
│
└── RDPランチャ\                        … ポータブル版アップ正本
    ├── version.txt
    ├── PmAiRpaLuncher_version_upgrade.zip
    ├── PmAiRpaLuncher_portable.zip
    └── build-manifest.json
```

| 用途 | 正本パス |
|------|----------|
| ポータブル bundle（版アップ ZIP 等） | `\\192.168.0.101\共有フォルダ\掲示板\rpa_luncher\RDPランチャ` |
| 接続先 exe + ini + 資格情報 JSON | `\\192.168.0.101\共有フォルダ\掲示板\rpa_luncher\`（**ルート**） |
| 操作者 bin | `\\192.168.0.101\共有フォルダ\掲示板\rpa_luncher\DATA` |
| ローカルビルド出力 | リポジトリ直下 `rpa_luncher_release\PmAiRpaLuncher\` |

**ini / JSON / exe は同じ「配備先」フォルダに揃える**（C# は ini の親ディレクトリから JSON を探す）。配備先がルートなら上記ルート、コード既定の `RDPランチャ` だけに置く運用なら環境変数を変更しない。

---

## トラブルシューティング

### 資格情報（KONAN / RDP_LAUNCHER）

| 症状 | 確認 |
|------|------|
| ログ `factory=KONAN` で未設定 | JSON に `factories.KONAN.操作者名` があるか。無ければ `RDP_LAUNCHER` ブロックを確認（C# はフォールバックあり） |
| Java で保存したが C# が読めない | `operator-aladdin-credentials.launcher.json` が **ini と同じフォルダ** にあるか |
| 操作者不一致 | ini の `操作者=` と JSON のキー名が一致するか（全角・空白） |
| 接続先 only | タスクスケジューラの `PM_AI_FACTORY_SITE` / `PM_AI_OPERATOR_USER` |

接続前に Java UI の「資格情報を保存」を実行し、共有フォルダ上の JSON 更新時刻を確認する。

### ログ場所

| コンポーネント | パス |
|----------------|------|
| **PmAiRpaLuncher（Java）** | `%USERPROFILE%\.pm-ai-desktop-rdp\`（起動失敗時 bat が案内） |
| **PmAiRdpRemoteLauncher（C#）** | 接続先 `%TEMP%\PM-AI-RDP-Launcher\launcher-yyyyMMdd.log` |
| **C# 共有ミラー** | ini 配備先の `launcher-logs\launcher-yyyyMMdd.log`（無ければ自動作成） |
| **Java UI プレビュー** | リモートデスクトップタブから共有ログを参照可能 |

### ビルド・起動その他

| 症状 | 対処 |
|------|------|
| WSL から UNC デプロイ失敗 | `-SkipCanonicalDeploy` でローカル `rpa_luncher_release\` までビルドし、手動コピー |
| ポータブル exe 起動不可 | `runtime\bin\java.exe` が Defender に削除されていないか。`launch-pm-ai-rpa-launcher.bat` を試す |
| 日本語パスで jpackage 失敗 | 非 ASCII リポジトリパス時は自動的に `%TEMP%` にステージング（`PM_AI_JPACKAGE_DEST` でも指定可） |
| ini が読まれない | `PmAiRdpRemoteLauncher.exe 操作者名` 引数と `{操作者}_RPA設定.ini` の存在を確認。レガシー `DATA\` のみの配置はフォールバック |

---

## 関連スクリプト・ソース（開発者向け）

| ファイル | 内容 |
|----------|------|
| `fast_package_rdp_launcher.ps1` | リポジトリ直下ビルドエントリ |
| `code_java/package_rdp_launcher_app.ps1` | jpackage・release・ZIP・canonical deploy |
| `code_java/run-pm-ai-remote-desktop.ps1` | 開発起動 |
| `scripts/build-rdp-remote-launcher.ps1` | 接続先 C# exe |
| `scripts/build-rdp-desktop-launcher.ps1` | 操作者 PC 用 C# スタブ |
| `scripts/rdp-apply-portable-update.ps1` | ポータブル版アップ適用 |
| `リモートデスクトップ・ランチャー整理.md` | データフロー・引数・シーケンス図 |
