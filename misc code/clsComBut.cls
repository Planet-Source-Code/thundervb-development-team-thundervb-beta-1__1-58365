VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComBut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Public WithEvents cbar As CommandBarEvents
Attribute cbar.VB_VarHelpID = -1
Public cbarObj As Office.CommandBarControl
Dim con As Connect

Dim c_id As Long

Function init(ctrl As CommandBarControl, from As Connect, IDc As Long) As Long

        c_id = IDc
        Set con = from
        Set cbarObj = ctrl
        Set cbar = VBI.Events.CommandBarEvents(cbarObj)
        init = c_id
        
End Function

Private Sub cbar_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    con.cmd_Click c_id, CommandBarControl, handled, CancelDefault
    
End Sub

Private Sub Class_Initialize()
    'nothing
End Sub

Private Sub Class_Terminate()
    
    If cbarObj Is Nothing Then Exit Sub
    
    cbarObj.Delete
    
End Sub
