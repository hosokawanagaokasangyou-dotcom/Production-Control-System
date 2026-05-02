// 出典: Cursor チャットでユーザーが貼付した Power Query (M)
// PQ: _q加工計画DATA_実績比較用

let
    // 1. フォルダから全ファイルを取得し、新しい順に並べ替え
    ソース = Folder.Files("\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\生産管理システム\管理システム\●DATA\生産計画問合せ"),
    並べ替えられたファイル = Table.Sort(ソース,{{"Date created", Order.Descending}}),
    
    // 最新の8ファイルに限定して処理を高速化
    対象ファイル = Table.FirstN(並べ替えられたファイル, 20),

    #"フィルター選択された非表示の File1" = Table.SelectRows(対象ファイル, each [Attributes]?[Hidden]? <> true),

    // 2. 各ファイルに対して個別にデータクレンジングと見出し変換を行う（ループ処理）
    個別処理 = Table.AddColumn(#"フィルター選択された非表示の File1", "処理済みテーブル", each 
        let
            現在の抽出時間 = [Date created],
            現在のテーブル = #"ファイルの変換 (3)"([Content]),
            
            削除された最初の行 = Table.Skip(現在のテーブル,3),
            置き換えられた値 = Table.ReplaceValue(削除された最初の行,"",null,Replacer.ReplaceValue,{"倉庫       : 520201 湖南工場01本倉庫", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43"}),

            // ▼================ 修正：フィル対象を見出しエリアだけに限定 ================▼
            // 見出しの構成に必要な最初の2行（日付行とラベル行）だけを切り出す
            ヘッダーエリア = Table.FirstN(置き換えられた値, 2),
            データエリア = Table.Skip(置き換えられた値, 2),
            
            フィル対象列 = {"倉庫       : 520201 湖南工場01本倉庫", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43"},
            
            // ヘッダーエリア内でのみ「上へフィル」を実行
            限定フィル = Table.FillUp(ヘッダーエリア, フィル対象列),
            
            // フィルしたヘッダーと、触れていないデータエリアを再結合
            結合 = Table.Combine({限定フィル, データエリア}),
            昇格されたヘッダー数 = Table.PromoteHeaders(結合, [PromoteAllScalars=true]),
            // ▲========================================================================▲
            
            旧ヘッダー = Table.ColumnNames(昇格されたヘッダー数),
            先頭行レコード = Table.First(昇格されたヘッダー数),

            受注日リスト = try 昇格されたヘッダー数[受注日] otherwise {},
            有効な受注日 = List.RemoveNulls(List.Transform(受注日リスト, each try Number.From(_) otherwise null)),
            基準日 = if List.Count(有効な受注日) > 0 
                    then Date.AddDays(#date(1899, 12, 30), List.Min(有効な受注日)) 
                    else Date.From(現在の抽出時間),
            基準年 = Date.Year(基準日),
            基準月 = Date.Month(基準日),

            新ヘッダー = List.Transform(旧ヘッダー, each 
                let
                    パーツ = Text.Split(_, "_"),
                    末尾 = List.Last(パーツ),
                    は連番 = (try Number.From(末尾) <> null otherwise false),
                    本来の見出し = if List.Count(パーツ) > 1 and は連番 
                                  then Text.Combine(List.RemoveLastN(パーツ, 1), "_") 
                                  else _,
                    
                    先頭行の値 = if 先頭行レコード = null then "" 
                                 else let v = Record.Field(先頭行レコード, _) in 
                                      if v = null then "" else Text.From(v),
                    
                    結合済み = if 本来の見出し = 先頭行の値 or 先頭行の値 = "" then 
                                 本来の見出し 
                               else 
                                 本来の見出し & "_" & 先頭行の値,
                    
                    アンダーバー分割 = Text.Split(結合済み, "_"),
                    見出し前半 = アンダーバー分割{0},
                    
                    日付候補 = Text.Split(見出し前半, "("){0},
                    スラッシュ分割 = Text.Split(日付候補, "/"),
                    
                    新見出し = if List.Count(スラッシュ分割) = 2 and (try Number.From(スラッシュ分割{0}) <> null otherwise false) then
                        let
                            月 = Number.From(スラッシュ分割{0}),
                            日 = Number.From(スラッシュ分割{1}),
                            
                            対象年 = if 基準月 >= 11 and 月 <= 2 then 基準年 + 1
                                    else if 基準月 <= 2 and 月 >= 11 then 基準年 - 1
                                    else 基準年,
                            
                            新日付 = Text.From(対象年) & "/" & Text.PadStart(Text.From(月), 2, "0") & "/" & Text.PadStart(Text.From(日), 2, "0"),
                            
                            残り = if List.Count(アンダーバー分割) > 1 then "_" & Text.Combine(List.Skip(アンダーバー分割, 1), "_") else ""
                        in
                            新日付 & 残り
                    else
                        結合済み
                in
                    新見出し
            ),
            
            見出しの更新 = Table.RenameColumns(昇格されたヘッダー数, List.Zip({旧ヘッダー, 新ヘッダー})),
            不要な先頭行を削除 = if 先頭行レコード = null then 見出しの更新 else Table.Skip(見出しの更新, 1),

            // 回答納期の空白対応
            回答納期のクリーンアップ = Table.ReplaceValue(不要な先頭行を削除, "", null, Replacer.ReplaceValue, {"回答納期"}),

            変更された型1 = Table.TransformColumnTypes(回答納期のクリーンアップ,{{"受注日", type date}, {"原反投入日", type date}, {"出荷日", type date}, {"指定納期", type date}, {"回答納期", type date}, {"加工開始日", type date}, {"加工完了日", type date}}),
            フィルターされた行 = Table.SelectRows(変更された型1, each ([依頼NO] <> null and [依頼NO] <> "")),
            
            すべての列 = Table.ColumnNames(フィルターされた行),
            削除対象の列 = List.Select(すべての列, each Text.Contains(_, "_加工速度") or Text.Contains(_, "_加工時間")),
            不要列の削除 = Table.RemoveColumns(フィルターされた行, 削除対象の列),
            
            抽出時間列の追加 = Table.AddColumn(不要列の削除, "抽出時間", each 現在の抽出時間, type datetime)
        in
            抽出時間列の追加
    ),

    // 3. 全てのファイルのテーブルを縦に結合
    結合されたテーブル = Table.Combine(個別処理[処理済みテーブル]),

    // 4. 重複の削除（Upsert）
    並べ替えられた全データ = Table.Sort(結合されたテーブル,{{"抽出時間", Order.Descending}})
in
    並べ替えられた全データ
</user_query>