Attribute VB_Name = "mdlExcelGreper"
Option Explicit
' ExcelGreper�Ɋւ��郂�W���[���A�萔���`���܂�.
Private Const MODULE_NAME As String = "mdlExcelGreper"


'------------------------------------------------
' �I�u�W�F�N�g���
'------------------------------------------------
' �Z��
Public Const OBJECT_TYPE_CELL As String = "TYPE_CELL"
' �V�F�C�v
Public Const OBJECT_TYPE_SHAPE As String = "TYPE_SHAPE"

' Grep���ʂ̃J������
Public Const GREP_RESULT_COLUMN_COUNT As Integer = 5

'
' �I�u�W�F�N�g��ނ���I�u�W�F�N�g�����擾���܂�.
'
' @param objType �I�u�W�F�N�g���
' @return �I�u�W�F�N�g��
'
Function ObjectTypeToName(ByRef objType As String) As String
    Dim objName As String

    If objType = OBJECT_TYPE_CELL Then
        objName = "�Z��"
    ElseIf objType = OBJECT_TYPE_SHAPE Then
        objName = "�V�F�C�v"
    Else
        objName = ""
    End If

    ObjectTypeToName = objName
End Function

