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
' Excel�t�@�C����Ώۂ�Grep�������s���܂�.
Private Const MODULE_NAME As String = "ExcelGreper"


' �����L�[���[�h
Private m_Keyword As String

' �����Ώۃt�@�C���p�X
Private m_FilePaths As Collection

' �啶������������ʂ��邩
Private m_IsIgnoreCase As Boolean

' ���K�\���̎g�p�L��
Private m_IsUseRegExp As Boolean

' ���K�\��
Private m_RegExp As RegExp

' InStr
Private m_Compare As Integer

'
' �R���X�g���N�^
'
Private Sub Class_Initialize()
    m_Keyword = ""
    Set m_FilePaths = New Collection
    m_IsUseRegExp = False
End Sub

'
' �f�X�g���N�^
'
Private Sub Class_Terminate()
    Set m_FilePaths = Nothing
    Set m_RegExp = Nothing
End Sub

'
' �C���X�^���X�̏������������s���܂�.
'
' @param keyword �����L�[���[�h
' @param filePaths �����Ώۂ̃t�@�C���p�X���X�g
' @param isIgnoreCase �啶������������ʂ��邩
' @param isUseRegExp ���K�\���̎g�p�L�� �f�t�H���g�l��false
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
' Grep���������s���܂�.
'
' @return resultList Grep���ʃ��X�g
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
' �V�[�g���̃Z����ΏۂɌ������s���܂�.
' �������ʂ�resultList�Ɋi�[����܂�.
'
' @param sheet �����Ώۂ̃V�[�g
' @param resultList Grep���ʂ��i�[���郊�X�g
' @param filePath Grep���s�����t�@�C���̃p�X
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

            ' �L�[���[�h���܂܂�Ă��邩����
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
' �V�[�g���̃V�F�C�v��ΏۂɌ������s���܂�.
' �������ʂ�resultList�Ɋi�[����܂�.
'
' @param sheet �����Ώۂ̃V�[�g
' @param resultList Grep���ʂ��i�[���郊�X�g
' @param filePath Grep���s�����t�@�C���̃p�X
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
            ' �V�F�C�v���z�u����Ă����΍��W(����)����Z�����W���擾
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
' �e�L�X�g���ɃL�[���[�h���܂܂�邩���肵�܂�.
'
' @param text �����L�[���[�h
' @return true:�����L�[���[�h���܂܂��, false:�����L�[���[�h���܂܂�Ȃ�
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