---
name: fast_package_app の ZIP に大物デバッグ成果物が混入する問題の除外再設計
overview: code_java/ 直下の JVM ヒープダンプ（*.hprof, *.hprof.p[0-9] 等）が package_workspace_copy.ps1 の filesystem mirror で pm-ai-data/code_java/ にコピーされ、Step 8 で生成する PMD_initial_install.zip / PMD_version_upgrade.zip がそれぞれ約 6.1 GB に膨張する事象への対処。.gitignore は無視されるため、明示的なファイル名パターン除外を追加し、過去プラン（portable_bundle_ui_exclude_defaults.plan.md）と同形式で経緯と方針を残す。
todos:
  - id: add-hprof-exclude
    content: package_workspace_copy.ps1 の $excludedFileNamePatterns に *.hprof, *.hprof.* を追加し、code_java/ 直下の JVM ヒープダンプが pm-ai-data に紛れ込まないようにする
    status: completed
  - id: add-misc-junk-exclude
    content: 同パターンに *.heapsnapshot / *.dump / *.tmp / Thumbs.db / desktop.ini を追加し、Windows / IDE が落とす一時ゴミも一律で外す
    status: completed
  - id: readme-portable-sync
    content: Copy-BundleToDist が書く pm-ai-data/README_PORTABLE.txt の「Excluded files (all profiles)」一文に新パターンを追記する
    status: completed
  - id: verify-rebuild
    content: 次回 fast_package_app.ps1 実行時に、pm-ai-package-release/PMD_initial_install.zip / PMD_version_upgrade.zip が再び ~230 MB 帯に戻ることを確認する。残った 6.1 GB ZIP と中間フォルダ pm-ai-package-release/PMD_initial_install/ pm-ai-package-release/PMD_version_upgrade/ は手動で削除する
    status: pending
  - id: keep-source-hprof
    content: code_java/ 直下の *.hprof は解析用にローカル保持を継続する（.gitignore は既存。`.gitignore` で git からは除外済み・filesystem からは消さない）
    status: completed
  - id: review-code_java-mirror-scope
    content: pm-ai-data/code_java/ にソース一式（src/, pom.xml 他）を入れる必要があるかを後日見直す。実行に不要なら除外して ZIP をさらに縮小できる（過去 5/9 ZIP では 1.8 MB / 202 ファイル）
    status: pending
isProject: false
---

# fast_package_app の ZIP に大物デバッグ成果物が混入する問題の除外再設計

## 背景・観測値

- 2026-05-09 時点の `PMD_initial_install.zip`: **240 MB**、`PMD_version_upgrade.zip`: **236 MB**（両方とも 10,856 / 10,831 件、ファイル群は通常）。
- 2026-05-10 22:19 のリリース実行で生成された ZIP: **6,085 MB / 6,086 MB**（974 件）。
- 中身を `python -m zipfile` で内訳した結果、`pm-ai-data/code_java/java_pid*.hprof` 系 9 ファイルだけで raw 34 GB / zip 5,711 MB を占めることが判明。

| エントリ | raw MB | zip MB |
|---|---:|---:|
| `pm-ai-data/code_java/java_pid8348.hprof` | 6,685.6 | 1,214.7 |
| `pm-ai-data/code_java/java_pid9152.hprof` | 6,200.7 | 1,233.8 |
| `pm-ai-data/code_java/java_pid5328.hprof` | 5,822.7 | 739.7 |
| `pm-ai-data/code_java/java_pid15412.hprof` | 5,822.4 | 757.0 |
| `pm-ai-data/code_java/java_pid10440.hprof` | 2,757.5 | 519.8 |
| `pm-ai-data/code_java/java_pid9152.hprof.p3` 他 `.p1` `.p2` `.p3` | 計 約 6.7 GB | 計 約 1.2 GB |

`code_java/` 直下に物理ファイルとしては 9 件（合計 34 GB）残っており、これは `-XX:+HeapDumpOnOutOfMemoryError` によって Java/Maven 実行で生成されたもの。**ローカル解析用に保持する方針**（ユーザー指示）。

## 原因

- `code_java/.gitignore` ・ ルート `.gitignore` には既に `*.hprof` / `*.hprof.*` が登録され git からは外れている。
- しかし `code_java/package_workspace_copy.ps1` の `Copy-WorkspaceTreeWithExplicitExclusions` は **filesystem walk による mirror** であり `.gitignore` を見ない。`*.log` と `~$*` 以外の拡張子フィルタが無く、hprof が透過してコピーされていた。

## 実装サマリ（コード参照）

### パッケージ（PowerShell）

- **ファイル**: [`code_java/package_workspace_copy.ps1`](../../code_java/package_workspace_copy.ps1) の `$excludedFileNamePatterns`
- **追加パターン**:
  - `*.hprof` / `*.hprof.*` … JVM ヒープダンプ（本件の主犯）
  - `*.heapsnapshot` … Chromium 系のヒープスナップショット
  - `*.dump` / `*.mdmp` … 任意ダンプ・Windows ミニダンプ
  - `*.tmp` … 一般一時ファイル（`~$*` で拾えないものの保険）
  - `Thumbs.db` / `desktop.ini` … Windows Explorer 由来のメタファイル
- **判定**: 既存の `Test-IsExcludedFileLeaf` が `-like` で評価する。`*.hprof.*` で `java_pid10440.hprof.p1` 等にもマッチする。

### README_PORTABLE.txt（バンドルに同梱されるテキスト）

- **ファイル**: リポジトリ直下 [`fast_package_app.ps1`](../../fast_package_app.ps1) の `Copy-BundleToDist`
- **修正**: `Excluded files (all profiles): *.log, ~$*` を新パターン込みの一文に置換。

## 検証

1. `code_java/` 直下に hprof を残したまま、`./fast_package_app.ps1` を再実行する。
2. `pm-ai-package-release/PMD_initial_install.zip` と `PMD_version_upgrade.zip` の **ファイルサイズが ~230 MB 帯**に戻ることを確認する。
3. ZIP 内エントリを `python -m zipfile -l` で覗き、`java_pid*.hprof*` が **存在しない**ことを確認する。
4. 過去ビルドで残った `pm-ai-package-release/PMD_initial_install/` と `pm-ai-package-release/PMD_version_upgrade/` の中間フォルダ・古い 6.1 GB の ZIP は **手動削除**で良い（コードからは触らない）。

## 副次の検討

- 過去 5/9 の ZIP では `pm-ai-data/code_java/` は **1.8 MB / 202 ファイル**しかなく、ソース一式は `src/main/java/...` の小サイズで収まっている。実行に必要なものは jpackage 出力 (`PMD.exe` + `app/` + `runtime/`) と `pm-ai-data/runtime/python-embed` で揃うため、**`pm-ai-data/code_java/` 自体を将来的に除外**する余地がある（todo `review-code_java-mirror-scope`）。

## 関連コミット（参照用・履歴）

- `eaf4f6be perf: パッケージZIPを並列化しInitialの不要成果物を除外`（直前。output/plan/plans 除外と Step 8 並列化）
- 本プランの実装はこれに続く同一ファイル群の小さな差分で入る想定。

## エンコーディング（プレビュー文字化け対策）

- 本ファイルは **UTF-8（改行 LF）** で保存する。`/mnt/c/` 上で Shift_JIS として保存すると Markdown プレビューで文字化けするため、エディタの「文字コード: UTF-8」を確認すること。
