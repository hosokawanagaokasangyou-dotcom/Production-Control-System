# Power Query (M) ソース（チャット出典）

出典: Cursor チャットでユーザーが貼付した M を、エージェントトランスクリプトから抽出して `plan/*.m` に保存しています。Excel ブック内の最新 M と差分がある場合は **ブックを正**として `.m` を更新してください。

| ファイル | 内容（要約） |
|----------|----------------|
| [01_加工計画DATA_単一ファイル.m](01_加工計画DATA_単一ファイル.m) | 生産計画問合せ・`Date created` 先頭1ファイル |
| [02__q加工実績明細DATA.m](02__q加工実績明細DATA.m) | 加工実績明細DATA・`Date accessed` 先頭1件・累積等 |
| [03__q加工計画DATA_実績比較用.m](03__q加工計画DATA_実績比較用.m) | 同一フォルダ・先頭20ファイル結合（M 内コメントと件数は要確認） |
| [04__q結果_配台表.m](04__q結果_配台表.m) | 名前「フォルダパス」+ `結果_配台表.xlsx`・`_t結果_配台表` |

JavaFX 移行プランの「Power Query 正本」と併せて参照してください。

**ETL 実装時**: 共有フォルダからの「最新」選定ユーティリティは **§実装前に固める取り決め（L0〜L8）の L4**（本リポジトリでは [excelからjavafxへUI移行.plan.md](excelからjavafxへUI移行.plan.md)）およびプラン正本の **「本ドキュメントの保守手順」手順 5** と整合させる。

## JavaFX 移行プラン（全文）

- **[excelからjavafxへUI移行.plan.md](excelからjavafxへUI移行.plan.md)** … **リポジトリ上の正本**（Git で共有）。プロジェクトルートの **[`.cursor/plans/excelからjavafxへui移行_e2f0b739.plan.md`](../../.cursor/plans/excelからjavafxへui移行_e2f0b739.plan.md)** に **同期（マージ）**し、Cursor のプラン UI と整合させる。
- **運用**: 編集は **`plan/excelからjavafxへUI移行.plan.md` を正**とし、変更後に上記 `.cursor/plans` へコピーする（またはその逆を一度決めて固定）。両者がずれたときは **Git に載っている `plan/` 側を優先**して `.cursor/plans` を上書きする。
