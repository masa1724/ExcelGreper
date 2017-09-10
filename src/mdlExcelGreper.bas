Attribute VB_Name = "mdlExcelGreper"
Option Explicit
' ExcelGreperに関するモジュール、定数を定義します.
Private Const MODULE_NAME As String = "mdlExcelGreper"


'------------------------------------------------
' オブジェクト種類
'------------------------------------------------
' セル
Public Const OBJECT_TYPE_CELL As String = "TYPE_CELL"
' シェイプ
Public Const OBJECT_TYPE_SHAPE As String = "TYPE_SHAPE"

' Grep結果のカラム数
Public Const GREP_RESULT_COLUMN_COUNT As Integer = 5

'
' オブジェクト種類からオブジェクト名を取得します.
'
' @param objType オブジェクト種類
' @return オブジェクト名
'
Function ObjectTypeToName(ByRef objType As String) As String
    Dim objName As String

    If objType = OBJECT_TYPE_CELL Then
        objName = "セル"
    ElseIf objType = OBJECT_TYPE_SHAPE Then
        objName = "シェイプ"
    Else
        objName = ""
    End If

    ObjectTypeToName = objName
End Function

