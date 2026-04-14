Attribute VB_Name = "AppMsgBox"
Option Explicit

' MsgBox の代替。ユーザーフォーム UF_LargeMessage で本文を大きめフォント表示する。
' ブックに UF_LargeMessage をインポート済みであること（同フォルダの UF_LargeMessage.frm）。

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
