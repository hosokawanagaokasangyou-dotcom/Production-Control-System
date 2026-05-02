// 出典: Cursor チャットでユーザーが貼付した Power Query (M)
// PQ: 加工計画DATA（単一ファイル・Folder.Files 先頭1件）

let
    // 1. フォルダから最新の1ファイルのみを特定
    パス = "\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\生産管理システム\管理システム\●DATA\生産計画問合せ",
    ソース = Folder.Files(パス),
    非表示除外 = Table.SelectRows(ソース, each [Attributes]?[Hidden]? <> true),
    並べ替えられたファイル = Table.Sort(非表示除外,{{"Date created", Order.Descending}}),
    最新ファイル = 並べ替えられたファイル{0},
    現在の抽出時間 = 最新ファイル[Date created],
    
    // 2. ファイルの内容をテーブルに展開
    現在のテーブル = ファイルの変換(最新ファイル[Content]),
    
    // 3. データクレンジング（Table.TransformColumnsで列指定を自動化）
    削除された最初の行 = Table.Skip(現在のテーブル, 3),
    全列名 = Table.ColumnNames(削除された最初の行),
    // 空文字をnullに変換する処理を全列に対して一括適用（速度向上）
    置き換えられた値 = Table.TransformColumns(削除された最初の行, 
        List.Transform(全列名, (col) => {col, each if _ = "" then null else _})),

    // 4. フィル対象を見出しエリアだけに限定
    ヘッダーエリア = Table.FirstN(置き換えられた値, 2),
    データエリア = Table.Skip(置き換えられた値, 2),
    限定フィル = Table.FillUp(ヘッダーエリア, 全列名),
    結合 = Table.Combine({限定フィル, データエリア}),
    昇格されたヘッダー数 = Table.PromoteHeaders(結合, [PromoteAllScalars=true]),
    
    // 5. 【高速化の肝】新ヘッダー計算の事前確定
    旧ヘッダー = Table.ColumnNames(昇格されたヘッダー数),
    先頭行レコード = Table.First(昇格されたヘッダー数),

    // 基準日の計算（一回だけ計算して変数に格納）
    受注日リスト = try 昇格されたヘッダー数[受注日] otherwise {},
    有効な受注日 = List.RemoveNulls(List.Transform(受注日リスト, each try Number.From(_) otherwise null)),
    基準日 = if List.Count(有効な受注日) > 0 
             then Date.AddDays(#date(1899, 12, 30), List.Min(有効な受注日)) 
             else Date.From(現在の抽出時間),
    基準年 = Date.Year(基準日),
    基準月 = Date.Month(基準日),

    // ヘッダー変換ロジックを List.Buffer でメモリにロックする
    新ヘッダーリスト = List.Buffer(List.Transform(旧ヘッダー, each 
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
    )),
    
    // 6. 変換の実行
    見出しの更新 = Table.RenameColumns(昇格されたヘッダー数, List.Zip({旧ヘッダー, 新ヘッダーリスト})),
    不要な先頭行を削除 = if 先頭行レコード = null then 見出しの更新 else Table.Skip(見出しの更新, 1),

    // 7. 最終成形
    回答納期のクリーンアップ = Table.ReplaceValue(不要な先頭行を削除, "", null, Replacer.ReplaceValue, {"回答納期"}),
    変更された型1 = Table.TransformColumnTypes(回答納期のクリーンアップ,{{"受注日", type date}, {"原反投入日", type date}, {"出荷日", type date}, {"指定納期", type date}, {"回答納期", type date}, {"加工開始日", type date}, {"加工完了日", type date}}),
    フィルターされた行 = Table.SelectRows(変更された型1, each ([依頼NO] <> null and [依頼NO] <> "")),
    
    すべての列名 = Table.ColumnNames(フィルターされた行),
    削除対象の列 = List.Select(すべての列名, each Text.Contains(_, "_加工速度") or Text.Contains(_, "_加工時間")),
    不要列の削除 = Table.RemoveColumns(フィルターされた行, 削除対象の列),
    
    抽出時間列の追加 = Table.AddColumn(不要列の削除, "抽出時間", each 現在の抽出時間, type datetime)
in
    抽出時間列の追加

@code/UI参照用_生産管理_AI配台(RC1).xlsx 現状のUIを参照するブックとなる
</user_query>