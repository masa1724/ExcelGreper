Attribute VB_Name = "mdlFormUtils"
Option Explicit
' �t�H�[���R���g���[���Ɋւ��郂�W���[�����`���܂�.
Private Const MODULE_NAME As String = "mdlFormUtils"


'
' �`�F�b�N�{�b�N�X�̃`�F�b�N��Ԃ��擾���܂�.
'
' @param sheetName �V�[�g��
' @param cbName �`�F�b�N�{�b�N�X�̃R���g���[����
' @return �`�F�b�N�{�b�N�X�̃`�F�b�N���
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
