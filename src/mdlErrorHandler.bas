Attribute VB_Name = "mdlErrorHandler"

Public Sub ThrowError(ByRef preErr As ErrObject, _
                      ByRef moduleName As String, _
                      ByRef methodName As String)
    Dim preSource As String
    Dim preNumber As Integer
    Dim preDescription As String
    Dim preHelpContext As String
    Dim preHelpFile As String
    'Dim preLastDllError As Integer

    preSource = preErr.Source
    preNumber = preErr.Number
    preDescription = preErr.Description
    preHelpContext = preErr.HelpContext
    preHelpFile = preErr.HelpFile
    'preLastDllError = preErr.LastDllError

    Err.Clear

    Err.Raise Source:=moduleName & "." & methodName & vbCrLf & preSource, _
              Number:=preNumber, _
              Description:=preDescription, _
              HelpContext:=preHelpContext, _
              HelpFile:=preHelpFile
              'LastDllError:=preLastDllError
End Sub
