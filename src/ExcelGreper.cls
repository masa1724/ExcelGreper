VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelGreper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Excelファイルを対象にGrep検索を行います.
Private Const MODULE_NAME As String = "ExcelGreper"


' 検索キーワード
Private m_Keyword As String

' 検索対象ファイルパス
Private m_FilePaths As Collection

' 大文字小文字を区別するか
Private m_IsIgnoreCase As Boolean

' 正規表現の使用有無
Private m_IsUseRegExp As Boolean

' 正規表現
Private m_RegExp As RegExp

' InStr
Private m_Compare As Integer

'
' コンストラクタ
'
Private Sub Class_Initialize()
    m_Keyword = ""
    Set m_FilePaths = New Collection
    m_IsUseRegExp = False
End Sub

'
' デストラクタ
'
Private Sub Class_Terminate()
    Set m_FilePaths = Nothing
    Set m_RegExp = Nothing
End Sub

'
' インスタンスの初期化処理を行います.
'
' @param keyword 検索キーワード
' @param filePaths 検索対象のファイルパスリスト
' @param isIgnoreCase 大文字小文字を区別するか
' @param isUseRegExp 正規表現の使用有無 デフォルト値はfalse
'
Public Sub Initialize(ByRef keyword As String, ByRef filePaths As Collection, ByVal isIgnoreCase As Boolean, ByVal isUseRegExp As Boolean)
    On Error GoTo ErrHandler

    m_Keyword = keyword
    Set m_FilePaths = filePaths
    m_IsIgnoreCase = isIgnoreCase
    m_IsUseRegExp = isUseRegExp

    If m_IsUseRegExp Then
        Set m_RegExp = New RegExp

        With m_RegExp
            .Pattern = m_Keyword
            .IgnoreCase = m_IsIgnoreCase
            .Global = True
        End With
    End If

    If m_IsIgnoreCase Then
        m_Compare = vbTextCompare
    Else
        m_Compare = vbBinaryCompare
    End If

    Exit Sub
ErrHandler:
    ThrowError Err, MODULE_NAME, "Initialize"
End Sub

'
' Grep検索を実行します.
'
' @return resultList Grep結果リスト
'
Public Function Execute() As Collection
    On Error GoTo ErrHandler

    Dim book As Workbook
    Dim sheet As Worksheet
    Dim resultList As Collection
    Dim FilePath As Variant
    Dim bk As Workbook

    Set resultList = New Collection

    For Each FilePath In m_FilePaths
        Set book = Workbooks.Open(FilePath)

        For Each sheet In book.Worksheets
            Call SearchCells(sheet, resultList, FilePath)
            Call SearchShapes(sheet, resultList, FilePath)
        Next

        book.Close
        Set book = Nothing
    Next

    Set Execute = resultList

    Exit Function
ErrHandler:
    If Not book Is Nothing Then
        For Each bk In Workbooks
            If bk.Name = book.Name Then
                book.Close
                Exit For
            End If
        Next
    End If

    ThrowError Err, MODULE_NAME, "Execute"
End Function

'
' シート内のセルを対象に検索を行います.
' 検索結果はresultListに格納されます.
'
' @param sheet 検索対象のシート
' @param resultList Grep結果を格納するリスト
' @param filePath Grepを行ったファイルのパス
'
Private Sub SearchCells(ByRef sheet As Worksheet, ByRef resultList As Collection, ByRef FilePath As Variant)
    On Error GoTo ErrHandler

    Dim usedRngValues As Variant
    Dim i As Integer
    Dim j As Integer
    Dim rowCnt As Long
    Dim colCnt As Long
    Dim Text As String
    Dim result As GrepResult
    Dim CellAddress As String

    usedRngValues = sheet.UsedRange
    rowCnt = UBound(usedRngValues, 1)
    colCnt = UBound(usedRngValues, 2)

    For i = 1 To rowCnt
        For j = 1 To colCnt
            Text = usedRngValues(i, j)

            ' キーワードが含まれているか判定
            If IsHit(Text) Then
                CellAddress = sheet.Name & "!" & sheet.Cells(i, j).Address(False, False)

                Set result = New GrepResult
                result.objectType = OBJECT_TYPE_CELL
                result.FilePath = FilePath
                result.CellAddress = CellAddress
                result.Text = Text
                resultList.Add result
            End If
        Next
    Next

    Exit Sub
ErrHandler:
    ThrowError Err, MODULE_NAME, "SearchCells"
End Sub

'
' シート内のシェイプを対象に検索を行います.
' 検索結果はresultListに格納されます.
'
' @param sheet 検索対象のシート
' @param resultList Grep結果を格納するリスト
' @param filePath Grepを行ったファイルのパス
'
Private Sub SearchShapes(ByRef sheet As Worksheet, ByRef resultList As Collection, ByRef FilePath As Variant)
    On Error GoTo ErrHandler

    Dim shape As shape
    Dim Text As String
    Dim result As GrepResult
    Dim CellAddress As String

    For Each shape In sheet.Shapes
        Text = shape.TextFrame.Characters.Text

        If IsHit(Text) Then
            ' シェイプが配置されている絶対座標(左上)からセル座標を取得
            CellAddress = sheet.Name & "!" & sheet.Range(shape.TopLeftCell, shape.BottomRightCell).Address(False, False)

            Set result = New GrepResult
            result.objectType = OBJECT_TYPE_SHAPE
            result.FilePath = FilePath
            result.CellAddress = shape.Name & " " & CellAddress
            result.Text = Text

            resultList.Add result
        End If
    Next

    Exit Sub
ErrHandler:
    ThrowError Err, MODULE_NAME, "SearchShapes"
End Sub

'
' テキスト内にキーワードが含まれるか判定します.
'
' @param text 検索キーワード
' @return true:検索キーワードが含まれる, false:検索キーワードが含まれない
Private Function IsHit(ByRef Text As String) As Boolean
    If Text = "" Then
        IsHit = False
        Exit Function
    End If

    If m_IsUseRegExp Then
        If m_RegExp.Test(Text) Then
            IsHit = True
            Exit Function
        End If
    Else
        If InStr(1, Text, m_Keyword, m_Compare) <> 0 Then
            IsHit = True
            Exit Function
        End If
    End If

    IsHit = False

    Exit Function
ErrHandler:
    ThrowError Err, MODULE_NAME, "IsHit"
End Function
