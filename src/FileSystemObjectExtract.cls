VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemObjectExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' FileSystemObjectの拡張メソッドを定義します.
Private Const MODULE_NAME As String = "FileSystemObjectExtract"


' FileSystemObject
Private m_FSO As FileSystemObject

'
' コンストラクタ
'
Private Sub Class_Initialize()
End Sub

'
' デストラクタ
'
Private Sub Class_Terminate()
    Set m_FSO = Nothing
End Sub

'
' インスタンスの初期化処理を行います.
'
' @param fso FileSystemObject
'
Public Sub Initialize(ByRef fso As FileSystemObject)
    Set m_FSO = fso
End Sub

Public Function GetFilePaths(ByVal dirPath As String, _
                             ByRef extentions As Collection, _
                             ByVal isRecursiveSearch, _
                             Optional ByRef excludeFiles As Collection = Nothing) As Collection
    On Error GoTo ErrHandler

    Dim tmpList As Collection: Set tmpList = New Collection
    Dim dirList As Collection: Set dirList = New Collection
    Dim fileList As Collection: Set fileList = New Collection
    Dim d As Variant
    Dim f As Variant
    Dim ex As Variant
    Dim e As Variant

    If excludeFiles Is Nothing Then
        Set excludeFiles = New Collection
        excludeFiles.Add ".*"
    End If

    If isRecursiveSearch Then
        Set dirList = GetSubFolder(dirPath, tmpList)
    Else
        dirList.Add dirPath
    End If

    Dim exFlg As Boolean: exFlg = True
    For Each d In dirList
        For Each f In m_FSO.GetFolder(d).files
            For Each ex In excludeFiles
                If Not f.Name Like ex Then
                    exFlg = False
                    Exit For
                End If
            Next
            If exFlg Then Exit For

            For Each e In extentions
                If m_FSO.GetExtensionName(f) = e Then
                    fileList.Add f.path
                End If
            Next
        Next
    Next

    Set GetFilePaths = fileList

    Exit Function
ErrHandler:
    ThrowError Err, MODULE_NAME, "GetFilePaths"
End Function

'
' ディレクトリパス.
'
' @param dirPath ファイルパス
' @param col ファイルパスリスト
' @return true:ファイルが開かれている, false:ファイルが開かれていない
'
Private Function GetSubFolder(ByVal dirPath As String, ByRef col As Collection) As Collection
    On Error GoTo ErrHandler

    Dim f As Variant

    ' ルートディレクトリ用
    If col.count = 0 Then
        col.Add dirPath
    End If

    For Each f In m_FSO.GetFolder(dirPath).SubFolders
        col.Add f
        GetSubFolder f, col
    Next f

    Set GetSubFolder = col

    Exit Function
ErrHandler:
    ThrowError Err, MODULE_NAME, "GetSubFolder"
End Function

'
' ファイルが既に開かれているか判定します.
'
' @param filePath ファイルパス
' @return true:ファイルが開かれている, false:ファイルが開かれていない
'
Public Function CheckAlreadyFileOpen(ByVal FilePath As String) As Boolean
    On Error Resume Next
    Open FilePath For Append As #1
    Close #1
    If Err.Number > 0 Then
        CheckAlreadyFileOpen = True
        Exit Function
    End If

    CheckAlreadyFileOpen = False
End Function

'
' ファイルパスからファイル種類を取得します.
'
' @param filePathOrDirPath ファイルパス/ディレクトリパス
' @return ファイル種類
'
Public Function GetFilePathType(ByVal filePathOrDirPath As String) As String
    On Error GoTo ErrHandler

    Dim isDir As Boolean
    Dim isFile As Boolean
    Dim filePathType As String

    isDir = m_FSO.FolderExists(filePathOrDirPath)
    isFile = m_FSO.FileExists(filePathOrDirPath)

    ' ディレクトリ、ファイル共に存在する場合はファイルと判断
    If isDir And isFile Then
        filePathType = FILE_PATH_TYPE_FILE
    ElseIf isDir Then
        filePathType = FILE_PATH_TYPE_DIRECTORY
    ElseIf isFile Then
        filePathType = FILE_PATH_TYPE_FILE
    Else
        filePathType = FILE_PATH_TYPE_NOTFOUND
    End If

    GetFilePathType = filePathType

    Exit Function
ErrHandler:
    ThrowError Err, MODULE_NAME, "GetFilePathType"
End Function
