VERSION 5.00
Begin VB.UserControl CommonFileDialog 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   ClipBehavior    =   0  'None
   InvisibleAtRuntime=   -1  'True
   MaskPicture     =   "CommonDialog.ctx":0000
   Picture         =   "CommonDialog.ctx":02E2
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   154
   ToolboxBitmap   =   "CommonDialog.ctx":0F24
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "CommonFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private decs
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private cdlg As OPENFILENAME
Private LastFileName As String
Private Type OPENFILENAME
  lStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String
End Type

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

Public Enum DialogFlags
  ALLOWMULTISELECT = OFN_ALLOWMULTISELECT
  CREATEPROMPT = OFN_CREATEPROMPT
  ENABLEHOOK = OFN_ENABLEHOOK
  ENABLETEMPLATE = OFN_ENABLETEMPLATE
  ENABLETEMPLATEHANDLE = OFN_ENABLETEMPLATEHANDLE
  EXPLORER = OFN_EXPLORER
  EXTENSIONDIFFERENT = OFN_EXTENSIONDIFFERENT
  FILEMUSTEXIST = OFN_FILEMUSTEXIST
  HIDEREADONLY = OFN_HIDEREADONLY
  LONGNAMES = OFN_LONGNAMES
  NOCHANGEDIR = OFN_NOCHANGEDIR
  NODEREFERENCELINKS = OFN_NODEREFERENCELINKS
  NOLONGNAMES = OFN_NOLONGNAMES
  NONETWORKBUTTON = OFN_NONETWORKBUTTON
  NOREADONLYRETURN = OFN_NOREADONLYRETURN
  NOTESTFILECREATE = OFN_NOTESTFILECREATE
  NOVALIDATE = OFN_NOVALIDATE
  OVERWRITEPROMPT = OFN_OVERWRITEPROMPT
  PATHMUSTEXIST = OFN_PATHMUSTEXIST
  ReadOnly = OFN_READONLY
  SHAREAWARE = OFN_SHAREAWARE
  SHAREFALLTHROUGH = OFN_SHAREFALLTHROUGH
  SHARENOWARN = OFN_SHARENOWARN
  SHAREWARN = OFN_SHAREWARN
  ShowHelp = OFN_SHOWHELP
End Enum


'Public props
Private CFm_CancelError       As Boolean
Private CFm_DialogTitle       As String
Private CFm_DefaultExt        As String
Private CFm_FileName          As String
Private CFm_FileTitle         As String
Private CFm_Filter            As String
Private CFm_Flags             As DialogFlags
Private CFm_InitDir           As String

Public Property Get CancelError() As Boolean
    CancelError = CFm_CancelError
End Property

Public Property Let CancelError(PropVal As Boolean)
    CFm_CancelError = PropVal
End Property

Public Property Get DialogTitle() As String
    DialogTitle = CFm_DialogTitle
End Property

Public Property Let DialogTitle(PropVal As String)
    CFm_DialogTitle = PropVal
End Property

Public Property Get DefaultExt() As String
    DefaultExt = CFm_DefaultExt
End Property

Public Property Let DefaultExt(PropVal As String)
    CFm_DefaultExt = PropVal
End Property

Public Property Get FileName() As String
    FileName = CFm_FileName
End Property

Public Property Let FileName(PropVal As String)
    CFm_FileName = PropVal
End Property

Public Property Get FileTitle() As String
    FileTitle = CFm_FileTitle
End Property

Public Property Let FileTitle(PropVal As String)
    CFm_FileTitle = PropVal
End Property

Public Property Get Filter() As String
    Filter = CFm_Filter
End Property

Public Property Let Filter(PropVal As String)
    CFm_Filter = PropVal
End Property

Public Property Get Flags() As DialogFlags
    Flags = CFm_Flags
End Property

Public Property Let Flags(PropVal As DialogFlags)
    CFm_Flags = PropVal
End Property

Public Property Get InitDir() As String
    InitDir = CFm_InitDir
End Property

Public Property Let InitDir(PropVal As String)
    CFm_InitDir = PropVal
End Property




Private Sub UserControl_Initialize()
UserControl.Height = 32 * 15
UserControl.Width = 32 * 15
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 32 * 15
UserControl.Width = 32 * 15
End Sub

Public Sub ShowOpen()
  Dim i As Integer
  Dim flt As String, idir As String, trez As String
  flt = Replace(Filter, "|", Chr(0))
  If Len(flt) = 0 Then flt = Replace("All Files (*.*)|*.*|", "|", Chr(0))
  If Right(flt, 1) <> Chr(0) Then flt = flt & Chr(0)
  If Len(InitDir) = 0 Then idir = FileName Else idir = InitDir
  cdlg.hwndOwner = UserControl.Parent.hWnd
  cdlg.hInstance = App.hInstance
  cdlg.lpstrFilter = flt
  cdlg.lpstrFile = FileName & String(255 - Len(FileName), Chr(0))
  cdlg.nMaxFile = 256
  cdlg.lpstrFileTitle = String(255, Chr(0))
  cdlg.nMaxFileTitle = 256
  cdlg.lpstrInitialDir = idir
  cdlg.lpstrTitle = DialogTitle
  cdlg.Flags = Flags
  cdlg.lStructSize = Len(cdlg)
  trez = IIf(GetOpenFileName(cdlg), Trim(cdlg.lpstrFile), "")
  If Len(trez) > 0 Then FileName = trez: FileTitle = cdlg.lpstrFileTitle Else If CancelError Then Err.Raise -1, "CDL control", "Cancel"
End Sub

Public Sub ShowSave()
  Dim i As Integer
  Dim flt As String, idir As String, trez As String
  flt = Replace(Filter, "|", Chr(0))
  If Len(flt) = 0 Then flt = Replace("All Files (*.*)|*.*|", "|", Chr(0))
  If Right(flt, 1) <> Chr(0) Then flt = flt & Chr(0)
  If Len(InitDir) = 0 Then idir = FileName Else idir = InitDir
  cdlg.hwndOwner = UserControl.Parent.hWnd
  cdlg.hInstance = App.hInstance
  cdlg.lpstrFilter = flt
  cdlg.lpstrFile = FileName & String(255 - Len(FileName), Chr(0))
  cdlg.nMaxFile = 256
  cdlg.lpstrFileTitle = String(255, Chr(0))
  cdlg.nMaxFileTitle = 256
  cdlg.lpstrInitialDir = idir
  cdlg.lpstrTitle = DialogTitle
  cdlg.Flags = Flags
  cdlg.lStructSize = Len(cdlg)
  trez = IIf(GetSaveFileName(cdlg), Trim(cdlg.lpstrFile), "")
  If Len(trez) > 0 Then FileName = trez: FileTitle = cdlg.lpstrFileTitle Else If CancelError Then Err.Raise -1, "CDL control", "Cancel"
End Sub
