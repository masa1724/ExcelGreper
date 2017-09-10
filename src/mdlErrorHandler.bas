Attribute VB_Name = "mdlErrorHandler"

    Public Sub ThrowError(ByRef preSource As String, _
            ByRef preDescription As String, _
            ByRef preHelpContext As String, _
            ByRef preHelpFile As String, _
                      ByRef moduleName As String, _
                      ByRef methodName As String)
    preError.Clear

    Error.Raise Source:=moduleName & "." & methodName & vbCrLf & preSource, _
                Description:=preDescription, _
                HelpContext:=preHelpContext, _
                HelpFile:=preHelpFile
End Sub
