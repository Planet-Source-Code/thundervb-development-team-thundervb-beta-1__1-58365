Attribute VB_Name = "modSubClasser"
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]


'Revision history:
'13/9/2004[dd/mm/yyyy] : Created by Raziel
'This is where all the subclassing init/update to curect window , message forwarding is done
'here is all the code that may crash something , and to keep that seperate the code that
'actualy does the message proccessing (witch is safe) is on a different module..
'This code acts as a abraction layer
'Works for MDI olny..
'
'1/10/2004[dd/mm/yyyy] : Edited by Raziel
'Works on SDI and MDI
'
'
'Notes
'hWnd In Window class is a hiden member .. to see it press f2
'and then right click , show hiden mebmers
'
'
'6/10/2004[dd/mm/yyyy] : Edited by Raziel
'Some heavy corection and fixes for mdi mode
'now , both MDIclient and child windows are hooked
'

Option Explicit


Public MainhWnd As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Public Const GWL_WNDPROC = (-4)

Type SubClassed
    
    hWnd As Long
    Previous As Long
    
End Type
Dim nCount As Long, sWind() As SubClassed, TimerId As Long

'if any new code windows are oppened then subclass them
Sub CheckStatus()
Dim temp As Long
    
    If App.LogMode = 0 Then Exit Sub
    If VBI Is Nothing Then GoTo t2
    If VBI.MainWindow Is Nothing Then GoTo t2
    

    If MainhWnd <> 0 Then
    
        temp = FindWindow("VbaWindow", vbNullString)
        If temp = 0 Then GoTo t1
        GetOldAddress temp
        Exit Sub
t1:
        temp = FindWindowEx(MainhWnd, 0, "MDIClient", vbNullString)
        If temp = 0 Or (temp = MainhWnd) Then GoTo nhp
        GetOldAddress temp
        GetOldAddress GetParent(temp)
nhp:
        temp = CodeWindowHWnd(VBI)
        If temp = 0 Then Exit Sub
        GetOldAddress temp
        
    Else
t2:
        temp = FindWindow("VbaWindow", vbNullString)
        If temp = 0 Then Exit Sub
        GetOldAddress temp
    End If

    
End Sub

Function KillAllSubClasses() As Boolean
Dim i As Long, Error As Boolean
    LogMsg "Un initing Subclassed windows", "modSubClasser", "KillAllSubClasses"
           
    For i = 0 To nCount - 1
    
    If sWind(i).Previous = 0 Then
    
        ErrorBox "Subclassed window " & sWind(i).hWnd & ",id=" & i & vbNewLine & _
                 "Original wndproc is 0.This may cause a crash after addin unload", _
                 "modSubClasser", "CheckStatus"
                 
        Error = True
        
    Else
        
        LogMsg "Subclassed window " & i & " hWnd:" & sWind(i).hWnd & " RestoreTo : " & sWind(i).Previous, "modSubClasser", "KillAllSubClasses"
        
        SetWindowLong sWind(i).hWnd, GWL_WNDPROC, sWind(i).Previous
        
    End If
    
    Next i

End Function

Function GetOldAddress(hWnd As Long) As Long
Dim i As Long
    
    For i = 0 To nCount - 1
    
    If sWind(i).hWnd = hWnd Then
        
        GetOldAddress = sWind(i).Previous
        Exit Function
        
    End If
    
    Next i
    
    ReDim Preserve sWind(nCount)
    sWind(nCount).Previous = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
    If sWind(nCount).Previous = 0 Then Exit Function
    sWind(nCount).hWnd = hWnd
    GetOldAddress = sWind(nCount).Previous
    LogMsg "Subclassed " & sWind(nCount).hWnd & "(id:" & nCount & ")" & _
           "OldProc is " & sWind(nCount).Previous & " new is " & GetAddr(AddressOf WindowProc), "modSubClasser", "GetOldAddress"
    nCount = nCount + 1
    
End Function

Private Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If VBI Is Nothing Then GoTo def
    If VBI.ActiveWindow Is Nothing Then GoTo def
    If VBI.ActiveWindow.Type <> 0 Then GoTo def
    Dim skipVB As Boolean, skipAft As Boolean
    
    WindowProcBef hWnd, uMsg, wParam, lParam, GetOldAddress(hWnd), skipVB, skipAft
    
    If skipVB = False Then
        WindowProc = CallWindowProc(GetOldAddress(hWnd), hWnd, uMsg, wParam, lParam)
    End If
    
    If skipAft = False Then
        WindowProcAft hWnd, uMsg, wParam, lParam, GetOldAddress(hWnd), WindowProc
    End If
    
    Exit Function
    
def:
    WindowProc = CallWindowProc(GetOldAddress(hWnd), hWnd, uMsg, wParam, lParam)
    
End Function

Function GetAddr(temp As Long) As Long

    GetAddr = temp

End Function
Public Sub ApiTimer(ByVal aStop As Boolean)

    If App.LogMode = 0 Then Exit Sub

    If aStop Then
        If TimerId Then
            KillTimer 0, TimerId
            TimerId = 0
        End If
    Else
        If TimerId Then
            KillTimer 0, TimerId
            TimerId = 0
        End If
        TimerId = SetTimer(0, 0, 200, AddressOf CheckStatus)
    End If

End Sub

' return the handle of the current
' code window, 0 if none
Function CodeWindowHWnd(VBInstance As VBIDE.VBE) As Long
Dim hWnd As Long, caption As String
    caption = VBInstance.ActiveCodePane. _
    Window.caption
    If VBInstance.DisplayModel = _
        vbext_dm_MDI Then
        ' get the handle of the main window
        ' hWnd is a hidden, undocumented
        ' property
        hWnd = VBInstance.MainWindow.hWnd
        ' in MDI mode there is an
        ' intermediate window, of class\
        ' MDIClient
        hWnd = FindWindowEx(hWnd, 0, _
        "MDIClient", vbNullString)
        ' finally we can get the hWnd of
        ' the code window
        CodeWindowHWnd = _
        FindWindowEx(hWnd, 0, _
        "VbaWindow", caption)
    Else
        ' no intermediate window in SDI mode
        CodeWindowHWnd = _
        FindWindow("VbaWindow", caption)
    End If
End Function
