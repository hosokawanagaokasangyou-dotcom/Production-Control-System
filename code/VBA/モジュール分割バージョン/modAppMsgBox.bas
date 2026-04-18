Option Explicit

' MsgBox の代替。Public 関数 AppMsgBox を呼び出す。
' ※標準モジュール名を AppMsgBox にすると関数 AppMsgBox と衝突し「モジュールではなく…」コンパイルエラーになるため、モジュール名は modAppMsgBox とする。
' ユーザーフォーム UF_LargeMessage を同じ VBA プロジェクトにインポート済みであること（UF_LargeMessage.frm + .frx）。

Public Function AppMsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = vbNullString) As VbMsgBoxResult
    Dim grp As Long
    grp = Buttons And &H7&
    If grp <> 0 And grp <> vbOKCancel And grp <> vbYesNo Then
        AppMsgBox = MsgBox(Prompt, Buttons, Title)
        Exit Function
    End If
    UF_LargeMessage.ApplySetup Prompt, Buttons, Title
    UF_LargeMessage.Show vbModal
    AppMsgBox = UF_LargeMessage.DialogResult
    Unload UF_LargeMessage
End Function
