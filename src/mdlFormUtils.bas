Attribute VB_Name = "mdlFormUtils"
Option Explicit
' フォームコントロールに関するモジュールを定義します.
Private Const MODULE_NAME As String = "mdlFormUtils"


'
' チェックボックスのチェック状態を取得します.
'
' @param sheetName シート名
' @param cbName チェックボックスのコントロール名
' @return チェックボックスのチェック状態
'
Public Function GetCheckBoxValue(ByRef sheetName As String, ByRef cbName As String) As Boolean
    On Error GoTo ErrHandler

    Dim onOff As Integer
    Dim checked As Boolean

    onOff = ThisWorkbook.Worksheets(sheetName).CheckBoxes(cbName).Value

    If onOff = xlOn Then
        checked = True
    Else
        checked = False
    End If

    GetCheckBoxValue = checked

    Exit Function
ErrHandler:
    ThrowError Err, MODULE_NAME, "GetCheckBoxValue"
End Function
