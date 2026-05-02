// 出典: Cursor チャットでユーザーが貼付した Power Query (M)
// PQ: _q結果_配台表

let
    // 1. Excelで定義した名前「フォルダパス」からパスを取得
    現在のパス = Excel.CurrentWorkbook(){[Name="フォルダパス"]}[Content]{0}[Column1],
    
    // 2. 取得したパスとファイル名を結合
    ファイルパス = 現在のパス & "結果_配台表.xlsx",

    // 3. ファイルパス変数を使用してソースを読み込む
    ソース = Excel.Workbook(File.Contents(ファイルパス), null, true),
    
    _t結果_配台表_Table = ソース{[Item="_t結果_配台表",Kind="Table"]}[Data],
    
    変更された型 = Table.TransformColumnTypes(_t結果_配台表_Table,{
        {"工程名", type text}, 
        {"機械名", type text}, 
        {"受注日", type any}, 
        {"受注NO", type any}, 
        {"依頼NO", type text}, 
        {"品名(原反)", type any}, 
        {"使用原反", type text}, 
        {"原反数", type any}, 
        {"品名(製品)", type any}, 
        {"製品名", type text}, 
        {"換算数量", Int64.Type}, 
        {"実加工数", Int64.Type}, 
        {"加工内容", type text}, 
        {"在庫場所", type text}, 
        {"原反投入日", Int64.Type}, 
        {"指定納期", Int64.Type}, 
        {"回答納期", Int64.Type}, 
        {"加工完了日", type any}, 
        {"加工完了区分", type text}, 
        {"実出来高", Int64.Type}, 
        {"計画合計", type any}, 
        {"原反投入場所", type any}, 
        {"配台日", Int64.Type}, 
        {"当日配台数量", Int64.Type}
    })
in
    変更された型
</user_query>