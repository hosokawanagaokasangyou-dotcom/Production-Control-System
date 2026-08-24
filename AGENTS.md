# AGENTS.md

このリポジトリは工場の生産管理・配台（ディスパッチ）システムです。構成は 2 つ:

- `code_java/` … JavaFX デスクトップアプリ（本体 UI）。Maven ビルド。メインクラス `jp.co.pm.ai.desktop.PmAiFxApp`。
- `code/python/` … 配台計画ロジック（`planning_core`）。常駐サービスではなく、Java 本体が `PythonProcessRunner` で子プロセスとして起動する CLI/ライブラリ。

コーディング規約・配台ロジックの正本などは `.cursor/rules/*.mdc` と `.cursorrules` を参照（本ファイルはそれらを置き換えない）。

## Cursor Cloud specific instructions

このセクションは Cloud Agent 向けの起動時の非自明な注意点のみを記す。依存インストール自体は起動時の update script（`uv` による venv 再同期）で行われる。

### ツールチェーン（スナップショットに含まれる想定・`~/.bashrc` で設定済み）

- **JDK 26**: `/opt/java/jdk26`（`JAVA_HOME`）。`pom.xml` は `maven.compiler.release=26` なので **JDK 21 では compile/test できない**。`~/.bashrc` で `JAVA_HOME` と `PATH` を設定済み。
- **Python 3.14**: `uv`（`~/.local/bin/uv`）で導入。仮想環境は **`/workspace/.venv`**（`pyproject.toml` は `requires-python >=3.14`）。
- Maven は wrapper（`code_java/mvnw`）を使う。**Gradle は無い**（`.cursor/rules/code-java-maven-build.mdc`）。依存は `~/.m2` にキャッシュ済み。

### ビルド / テスト

- 一括テスト: リポジトリ直下の `./test.sh`（`code_java` の `./mvnw test` と `code/python` の `pytest` を順に実行）。
- Java 単体: `cd code_java && ./mvnw -q compile` / `./mvnw test`。
- Python 単体: `/workspace/.venv/bin/python -m pytest tests/ -q`（`code/python` で実行）。`pytest` は `requirements.txt` に無いため update script で別途入れている。
- **既知の pre-existing 失敗（環境要因ではなく Linux では常に失敗、コード修正不要）**:
  - Java 7 件: `AgentDebugLogOverlayTest` 等は `C:\repo` などの **Windows パス前提**、`RdpRemoteLauncherDeployerTest`（RDP ランチャー）、`FactorySiteLogoSupportTest`（未同梱 png）、`RequestFormPreviewPdfFontsTest`（pdfbox 版差）。
  - Python 2 件: `test_stage3_input_builder.py` の複数日枝番テストは、ビルダが親単位で集約するため決定的に 2→1 になり失敗（日付非依存）。

### デスクトップアプリの起動（GUI）

- **表示先は `DISPLAY=:1`**（Xvfb 上の X サーバが用意されている）。ヘッドレスだと `PmAiFxApp` は exit code 2 で終了する。
- 起動コマンド（`code_java` で実行）:
  ```bash
  JAVA_HOME=/opt/java/jdk26 DISPLAY=:1 PM_AI_PYTHON=/workspace/.venv/bin/python \
    ./mvnw -q validate compile exec:exec@pm-ai-desktop -Dpm.ai.javafx.prism.skipGpuProbe=true
  ```
  - `exec:exec@pm-ai-desktop` は `validate` フェーズで OpenJFX の module-path プロパティを組むため、**`validate` を同一起動に含める**必要がある（`exec` 単体では module-path が空になる）。
  - GPU が無いため **software rendering（`prism.order=sw`）** で動かす。`-Dpm.ai.javafx.prism.skipGpuProbe=true` で GPU プローブ子プロセスを省略すると安定して起動する。
- **`PM_AI_PYTHON`** を `/workspace/.venv/bin/python` に設定すると、Java 本体が配台処理を回す際に正しいインタプリタで `planning_core` を起動できる（`StagePythonExecutable` の解決順の先頭）。
- Gemini（`google-genai`）による AI 備考解析は API キー未設定なら自動スキップされる（警告ログのみ、配台自体は動く）。任意機能。
