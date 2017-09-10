VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GrepResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Grep結果を保持します.


' キーワードが見つかったオブジェクトの種類(セル、シェイプ等)
Public objectType As String

' キーワードが見つかったファイルのパス
Public FilePath As String

' キーワードが見つかったセルのアドレス
' シェイプの場合は、シェイプ名 + シェイプが配置されている左上のセルのアドレスが設定されます.
Public CellAddress As String

' キーワードが見つかったセル、シェイプのテキスト
Public Text As String
