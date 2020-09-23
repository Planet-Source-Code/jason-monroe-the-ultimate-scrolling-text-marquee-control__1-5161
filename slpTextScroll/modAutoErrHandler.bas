Attribute VB_Name = "modAutoErrHandler"
Option Explicit

'This module has been added to your project by Auto Error Handler Beta. The software has
'placed an Error Handler in one or more of your forms, classes, modules, etc. and the
'global routine is here.
Public Sub ProjectErrorHandler(MyMethod As String)
    
    Dim sErrStr As String
    Dim uResult As VbMsgBoxResult
    
    sErrStr = "Error " & Trim$(Str$(Err)) & " in " & MyMethod    
    If Erl Then
        sErrStr = sErrStr & " (Line #: " & Erl & ")"
    End If
        
    sErrStr = sErrStr & vbCrLf & "while running " & App.EXEName & ".exe v" & Format$(App.Major, "#") & "." & Format$(App.Minor, "0#")
    sErrStr = sErrStr & " (Build " & Format$(App.Revision, "#0") & ")" & vbCrLf & vbCrLf & "Error = '" & Error$ & "'"
    
    uResult = MsgBox(sErrStr, vbOKOnly + vbCritical + vbApplicationModal, "Error")
    
End Sub
