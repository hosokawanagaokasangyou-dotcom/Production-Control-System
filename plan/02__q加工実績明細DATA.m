// 出典: Cursor チャットでユーザーが貼付した Power Query (M)
// PQ: _q加工実績明細DATA

let
    // ご指定いただいたデータソースのパス
    ソース = Folder.Files("\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\002  加工G\●検査表作成\加工実績明細DATA"),
    並べ替えられた行 = Table.Sort(ソース,{{"Date accessed", Order.Descending}}),
    保存された先頭行 = Table.FirstN(並べ替えられた行,1),

    // ★追加: 読み込んだファイルの「更新日時(Date modified)」を抽出して記憶しておく
    ファイル更新日時 = 保存された先頭行{0}[Date modified],

    #"フィルター選択された非表示の File1" = Table.SelectRows(保存された先頭行, each [Attributes]?[Hidden]? <> true),
    カスタム関数の呼び出し1 = Table.AddColumn(#"フィルター選択された非表示の File1", "ファイルの変換 (2)", each #"ファイルの変換 (2)"([Content])),
    #"名前が変更された列 1" = Table.RenameColumns(カスタム関数の呼び出し1, {"Name", "Source.Name"}),
    削除された他の列1 = Table.SelectColumns(#"名前が変更された列 1", {"Source.Name", "ファイルの変換 (2)"}),
    #"展開された ファイルの変換 (2)" = Table.ExpandTableColumn(削除された他の列1, "ファイルの変換 (2)", {"Name", "Data", "Item", "Kind", "Hidden"}, {"Name", "Data", "Item", "Kind", "Hidden"}),
    #"展開された Data" = Table.ExpandTableColumn(#"展開された ファイルの変換 (2)", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Column22", "Column23", "Column24", "Column25", "Column26", "Column27", "Column28", "Column29", "Column30", "Column31", "Column32", "Column33", "Column34", "Column35", "Column36", "Column37", "Column38", "Column39", "Column40", "Column41", "Column42", "Column43", "Column44", "Column45", "Column46", "Column47", "Column48", "Column49", "Column50", "Column51", "Column52", "Column53", "Column54", "Column55", "Column56", "Column57", "Column58", "Column59", "Column60", "Column61", "Column62", "Column63", "Column64", "Column65", "Column66", "Column67", "Column68", "Column69", "Column70", "Column71", "Column72", "Column73", "Column74", "Column75", "Column76", "Column77", "Column78", "Column79", "Column80", "Column81", "Column82", "Column83", "Column84", "Column85", "Column86", "Column87", "Column88", "Column89", "Column90", "Column91", "Column92", "Column93", "Column94", "Column95", "Column96", "Column97", "Column98", "Column99", "Column100", "Column101", "Column102", "Column103", "Column104", "Column105", "Column106", "Column107", "Column108", "Column109", "Column110", "Column111", "Column112", "Column113", "Column114", "Column115", "Column116", "Column117", "Column118", "Column119", "Column120", "Column121", "Column122", "Column123", "Column124", "Column125", "Column126", "Column127", "Column128", "Column129", "Column130", "Column131", "Column132", "Column133", "Column134", "Column135", "Column136", "Column137", "Column138", "Column139", "Column140", "Column141", "Column142", "Column143", "Column144", "Column145", "Column146", "Column147", "Column148", "Column149", "Column150", "Column151", "Column152", "Column153", "Column154", "Column155", "Column156", "Column157", "Column158", "Column159", "Column160", "Column161", "Column162", "Column163", "Column164", "Column165", "Column166", "Column167", "Column168", "Column169", "Column170", "Column171", "Column172", "Column173", "Column174", "Column175", "Column176", "Column177", "Column178", "Column179", "Column180", "Column181", "Column182", "Column183", "Column184", "Column185", "Column186", "Column187", "Column188", "Column189", "Column190", "Column191", "Column192", "Column193", "Column194", "Column195", "Column196", "Column197", "Column198", "Column199", "Column200", "Column201", "Column202", "Column203", "Column204", "Column205", "Column206", "Column207", "Column208", "Column209", "Column210", "Column211", "Column212", "Column213", "Column214", "Column215", "Column216", "Column217", "Column218", "Column219", "Column220", "Column221", "Column222", "Column223", "Column224", "Column225", "Column226", "Column227", "Column228", "Column229", "Column230", "Column231", "Column232", "Column233", "Column234", "Column235", "Column236", "Column237", "Column238", "Column239", "Column240", "Column241", "Column242", "Column243", "Column244", "Column245", "Column246", "Column247", "Column248", "Column249", "Column250", "Column251", "Column252", "Column253", "Column254", "Column255", "Column256", "Column257", "Column258", "Column259", "Column260", "Column261", "Column262", "Column263", "Column264", "Column265", "Column266", "Column267", "Column268", "Column269", "Column270", "Column271"}),
    フィルターされた行 = Table.SelectRows(#"展開された Data", each ([Column4] <> null)),
    昇格されたヘッダー数 = Table.PromoteHeaders(フィルターされた行, [PromoteAllScalars=true]),
    変更された型 = Table.TransformColumnTypes(昇格されたヘッダー数,{{"加工日", type date}, {"原反投入日1", type date}, {"加工完了予定日", type date}, {"加工開始予定日", type date}}),
    
    // ここから「検査備考」を削除しました
    削除された他の列 = Table.SelectColumns(変更された型,{"対象NO", "機械名", "工程名", "加工内容名", "原反製品判定区分", "加工日", "依頼NO", "加工内容", "開始時間", "開始分", "終了時間", "終了分", "停機時間", "停機分", "停機時間分", "休憩時間", "製造条件(内訳)", "加工担当者名1", "加工担当者名2", "加工担当者名3", "加工担当者名4", "加工担当者名5", "換算数量", "加工予定数", "製品NO1", "製品NO2", "原反NO1", "実加工数", "false"}),
    
    追加された加工開始日時 = Table.AddColumn(削除された他の列, "加工開始日時", each 
        #datetime(Date.Year([加工日]), Date.Month([加工日]), Date.Day([加工日]), Number.From([開始時間]), Number.From([開始分]), 0), 
        type datetime),
    
    追加された加工終了日時 = Table.AddColumn(追加された加工開始日時, "加工終了日時", each 
        #datetime(Date.Year([加工日]), Date.Month([加工日]), Date.Day([加工日]), Number.From([終了時間]), Number.From([終了分]), 0), 
        type datetime),

    追加された加工時間 = Table.AddColumn(追加された加工終了日時, "加工時間", each 
        Duration.TotalMinutes([加工終了日時] - [加工開始日時]) - (if [休憩時間] = null or [休憩時間] = "" then 0 else Number.From([休憩時間])), 
        type number),

    フィルターされた行1 = Table.SelectRows(追加された加工時間, each ([#"製造条件(内訳)"] = "長さ")),
    削除された他の列2 = Table.SelectColumns(フィルターされた行1,{"機械名", "工程名", "依頼NO", "停機時間分", "加工担当者名1", "加工担当者名2", "加工担当者名3", "加工担当者名4", "加工担当者名5", "換算数量", "加工予定数", "実加工数", "加工開始日時", "加工終了日時", "加工時間"}),
    
    追加された停機時間分_変換後 = Table.AddColumn(削除された他の列2, "停機時間分(変換後)", each 
        if [停機時間分] = null or [停機時間分] = "" then 0 
        else 
            let 
                split = Text.Split(Text.From([停機時間分]), ":"),
                h = Number.From(split{0}?),
                m = Number.From(if List.Count(split) > 1 then split{1} else 0)
            in 
                (h * 60) + m, 
        type number),

    // ★追加箇所: 停機時間を加算した日時列の作成
    追加された加算後日時 = Table.AddColumn(追加された停機時間分_変換後, "加工開始日時(停機時間加算後)", each 
        [加工開始日時] + #duration(0, 0, [#"停機時間分(変換後)"], 0), 
        type datetime),

    並べ替えられた列 = Table.ReorderColumns(追加された加算後日時,{"依頼NO", "機械名", "工程名", "加工開始日時", "加工開始日時(停機時間加算後)", "加工終了日時", "加工時間", "停機時間分(変換後)", "実加工数", "換算数量", "加工予定数", "加工担当者名1", "加工担当者名2", "加工担当者名3", "加工担当者名4", "加工担当者名5"}),

    // --- ここから累積計算 ---
    追加された日付 = Table.AddColumn(並べ替えられた列, "加工日付", each DateTime.Date([加工開始日時]), type date),
    日次集計 = Table.Group(追加された日付, {"依頼NO", "工程名", "加工日付"}, {{"日次実績", each List.Max([実加工数]), type number}}),
    累積計算 = Table.Group(日次集計, {"依頼NO", "工程名"}, {
        {"Data", each 
            let 
                並べ替え = Table.Sort(_, {{"加工日付", Order.Ascending}}),
                インデックス追加 = Table.AddIndexColumn(並べ替え, "Index", 1, 1, Int64.Type),
                累積列追加 = Table.AddColumn(インデックス追加, "累積実績", (row) => List.Sum(List.FirstN(インデックス追加[日次実績], row[Index])), type number)
            in 
                累積列追加
        }
    }),
    展開された累積 = Table.ExpandTableColumn(累積計算, "Data", {"加工日付", "累積実績"}, {"加工日付", "累積実績"}),
    マージされたクエリ = Table.NestedJoin(追加された日付, {"依頼NO", "工程名", "加工日付"}, 展開された累積, {"依頼NO", "工程名", "加工日付"}, "累積データ", JoinKind.LeftOuter),
    展開されたマージ = Table.ExpandTableColumn(マージされたクエリ, "累積データ", {"累積実績"}, {"累積実績"}),
    追加された完了率 = Table.AddColumn(展開されたマージ, "累積完了率", each if [換算数量] = null or [換算数量] = 0 then 0 else [累積実績] / [換算数量], type number),
    型の変更_完了率 = Table.TransformColumnTypes(追加された完了率,{{"累積完了率", Percentage.Type}}),
    削除された作業列 = Table.RemoveColumns(型の変更_完了率,{"加工日付"}),
    追加されたデータ抽出時間 = Table.AddColumn(削除された作業列, "データ抽出時間", each ファイル更新日時, type datetime),

    // 最終的な並べ替え
    並べ替えられた列1 = Table.ReorderColumns(追加されたデータ抽出時間,{"依頼NO", "機械名", "工程名", "停機時間分(変換後)", "加工開始日時", "加工開始日時(停機時間加算後)", "加工終了日時", "加工時間", "実加工数", "換算数量", "加工予定数", "加工担当者名1", "加工担当者名2", "加工担当者名3", "加工担当者名4", "加工担当者名5", "累積実績", "累積完了率", "データ抽出時間"})
in
    並べ替えられた列1
</user_query>