Attribute VB_Name = "modOther"
Option Explicit

Public Function CalledByPtr(ByVal Caller As Long) As Long
    
    If Caller = 0 Then
        CalledByPtr = MsgBox("I'm called by pointer ;)", vbInformation Or vbYesNoCancel, "ThunVB sample 1")
    Else
        CalledByPtr = MsgBox("I'm called From C code ;)", vbInformation Or vbYesNoCancel, "ThunVB sample 1")
    End If
    
End Function
