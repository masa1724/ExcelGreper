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
' Grep���ʂ�ێ����܂�.


' �L�[���[�h�����������I�u�W�F�N�g�̎��(�Z���A�V�F�C�v��)
Public objectType As String

' �L�[���[�h�����������t�@�C���̃p�X
Public FilePath As String

' �L�[���[�h�����������Z���̃A�h���X
' �V�F�C�v�̏ꍇ�́A�V�F�C�v�� + �V�F�C�v���z�u����Ă��鍶��̃Z���̃A�h���X���ݒ肳��܂�.
Public CellAddress As String

' �L�[���[�h�����������Z���A�V�F�C�v�̃e�L�X�g
Public Text As String