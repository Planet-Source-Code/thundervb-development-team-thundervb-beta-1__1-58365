VERSION 5.00
Begin VB.UserControl ExportList 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   282
   Begin VB.ListBox lstExp 
      Height          =   3660
      ItemData        =   "ExportList.ctx":0000
      Left            =   0
      List            =   "ExportList.ctx":0007
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "ExportList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]

Option Explicit


'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Raziel
'Control created , intial version

Property Let SelectedExports(data As String)
Dim temp() As String
    
    temp = Split(data, "**@@split_@@**")
    enumerate VBI, lstExp, temp
    
End Property

Private Sub UserControl_Resize()

    lstExp.Width = UserControl.Width / 15
    lstExp.Height = UserControl.Height / 15
    UserControl.Width = lstExp.Width * 15
    UserControl.Height = lstExp.Height * 15
    
End Sub

Property Get SelectedExports() As String
Dim data() As String, intTemp As Long

    ReDim Preserve data(1)
    For intTemp = 0 To lstExp.ListCount - 1
        If lstExp.Selected(intTemp) = True Then
            data(UBound(data)) = Split(lstExp.list(intTemp), " ")(0)
            ReDim Preserve data(UBound(data) + 1)
        End If
    Next
    
    ReDim Preserve data(UBound(data) - 1)
    SelectedExports = Join$(data, "**@@split_@@**")
    
End Property


Sub enumerate(vb As VBIDE.VBE, ToList As ListBox, exports() As String)
Dim Components As VBComponents
Dim cMembers As Members
Dim strtemp As String
Dim intTemp As Long
Dim cObjC As Long
Dim cObjM As Long

    If vb Is Nothing Then Exit Sub
    If vb.ActiveVBProject Is Nothing Then Exit Sub
    If vb.ActiveVBProject.VBComponents Is Nothing Then Exit Sub
    On Error Resume Next
    frmLoading.show
    DoEvents
    ToList.Clear
    
    'enumerate the procedures in every module file within
    'the current project
    Set Components = vb.ActiveVBProject.VBComponents
    For cObjC = 1 To Components.count
    
        If Components(cObjC).Type = vbext_ct_StdModule Then
            Set cMembers = Components(cObjC).codeModule.Members
            With cMembers
                For cObjM = 1 To .count
                    If .item(cObjM).Type = vbext_mt_Method Then
                        ToList.AddItem .item(cObjM).name & " (defined in " & Components(cObjC).name & ")"
                        'check if the procedure is mardked to be exported.
                        'if so, tick the box next to it.
                        For intTemp = 1 To UBound(exports)
                            If exports(intTemp) = .item(cObjM).name Then
                                ToList.Selected(ToList.ListCount - 1) = True
                            End If
                        Next
                    End If
                Next
            End With
        End If
    Next
    frmLoading.hide
    
End Sub
Public Sub Refresh()
Dim nullarr(0) As String

    enumerate VBI, lstExp, nullarr
    
End Sub
Private Sub UserControl_Initialize()
Dim nullarr(0) As String

    UserControl.ScaleMode = 3
    enumerate VBI, lstExp, nullarr
    
End Sub
