段階2 Java エンジン（PM_AI_STAGE2_ENGINE=java）と Python 子プロセスの回帰比較用メモ

1. 同一の PM_AI_* 環境で Python 段階2を一度実行し、出力フォルダの 計画*.xlsx / 計画*.json（および人員*）を基準として保存する。
2. PM_AI_STAGE2_ENGINE=java で同じ入力を実行し、成果物の format_version・シート集合・主要列の行数を比較する。
3. 本リポジトリの JUnit では `Stage2JavaEngineSmokeTest` が最小マスタ＋CSV で入出力の足場を検証する（Python 基準ファイルの同梱は任意）。
