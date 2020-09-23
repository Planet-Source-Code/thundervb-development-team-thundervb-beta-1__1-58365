VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDebugDeleteSettings 
      Caption         =   "Delete Settings"
      Height          =   495
      Left            =   1440
      TabIndex        =   35
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton cmdDebugEnumDefault 
      Caption         =   "Debug - Enum default controls"
      Height          =   495
      Left            =   2400
      TabIndex        =   34
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdDebugEnumLocal 
      Caption         =   "Debug - Enum local controls"
      Height          =   495
      Left            =   360
      TabIndex        =   33
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdDebugSave 
      Caption         =   "Debug - Save"
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdDebugLoad 
      Caption         =   "Debug - Load"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdStorno 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   5760
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdSet 
      Left            =   1680
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab sstSet 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   8
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   485
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmSettings.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "General_HookCompiler"
      Tab(0).Control(1)=   "General_PopUpWindow"
      Tab(0).Control(2)=   "General_AddTLB"
      Tab(0).Control(3)=   "General_HideErrorDialogs"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Paths"
      TabPicture(1)   =   "frmSettings.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblPaths_4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblPaths_3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblPaths_2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblPaths_1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblPaths_5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblPaths_6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdPaths_TextEditor"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdPaths_Packer"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdPaths_MIDL"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdPaths_ML"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Paths_Packer"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Paths_MIDL"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Paths_TextEditor"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Paths_ML"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Paths_CCompiler"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Paths_LIBFiles"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Paths_INCFiles"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cmdPaths_CCompiler"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "cmdPaths_LIBFiles"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmdINCFiles"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Packer"
      TabPicture(2)   =   "frmSettings.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblPacker_1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblPacker_2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Packer_UsePacker"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Packer_CmdLine1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Packer_CmdLine2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Packer_CmdLine3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Packer_Desc1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Packer_Desc3"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "optPacker_1"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "optPacker_2"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "optPacker_3"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Packer_ShowPackerOutPut"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Packer_Desc2"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "ASM"
      TabPicture(3)   =   "frmSettings.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblpreVasm"
      Tab(3).Control(1)=   "ASM_UseASMColors"
      Tab(3).Control(2)=   "Compile_PauseLink"
      Tab(3).Control(3)=   "cmdColor_DefaultASM"
      Tab(3).Control(4)=   "ASM_ASMColors"
      Tab(3).Control(5)=   "cmdColor_AddASM"
      Tab(3).Control(6)=   "cmdColor_DeleteASM"
      Tab(3).Control(7)=   "ASM_FixASMListings"
      Tab(3).Control(8)=   "ASM_CompileASMCode"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Debug"
      TabPicture(4)   =   "frmSettings.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Debug_OutDebLog"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Compile"
      TabPicture(5)   =   "frmSettings.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Compile_PauseAsm"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Compile_ModifyCmdLine"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "StdCall DLL"
      TabPicture(6)   =   "frmSettings.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblDLL_1"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "lblDLL_2"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "DLL_ExportSymbols"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "DLL_BaseAddress"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "DLL_LinkAsDLL"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "DLL_ExpList"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "DLL_EntryPoint"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "cdmDLL_AddDllMain"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).ControlCount=   8
      TabCaption(7)   =   "C"
      TabPicture(7)   =   "frmSettings.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblpreVc"
      Tab(7).Control(1)=   "C_CColors"
      Tab(7).Control(2)=   "cmdColor_DeleteC"
      Tab(7).Control(3)=   "cmdColor_AddC"
      Tab(7).Control(4)=   "cmdColor_DefaultC"
      Tab(7).Control(5)=   "C_UseCColors"
      Tab(7).Control(6)=   "C_CompileCCode"
      Tab(7).ControlCount=   7
      Begin VB.TextBox Packer_Desc2 
         Height          =   285
         Left            =   -72360
         TabIndex        =   74
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox Packer_ShowPackerOutPut 
         Caption         =   "Show packer output"
         Height          =   255
         Left            =   -72360
         TabIndex        =   71
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton optPacker_3 
         Height          =   255
         Left            =   -70800
         TabIndex        =   70
         Top             =   2535
         Width           =   255
      End
      Begin VB.OptionButton optPacker_2 
         Height          =   255
         Left            =   -70800
         TabIndex        =   69
         Top             =   2055
         Width           =   255
      End
      Begin VB.OptionButton optPacker_1 
         Height          =   255
         Left            =   -70800
         TabIndex        =   68
         Top             =   1575
         Width           =   255
      End
      Begin VB.TextBox Packer_Desc3 
         Height          =   285
         Left            =   -72360
         TabIndex        =   67
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Packer_Desc1 
         Height          =   285
         Left            =   -72360
         TabIndex        =   66
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Packer_CmdLine3 
         Height          =   285
         Left            =   -74760
         TabIndex        =   65
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox Packer_CmdLine2 
         Height          =   285
         Left            =   -74760
         TabIndex        =   64
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox Packer_CmdLine1 
         Height          =   285
         Left            =   -74760
         TabIndex        =   63
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton cdmDLL_AddDllMain 
         Caption         =   "Add DllMain"
         Height          =   375
         Left            =   -71520
         TabIndex        =   60
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CheckBox C_CompileCCode 
         Caption         =   "Compile C code"
         Height          =   250
         Left            =   -74760
         TabIndex        =   59
         Tag             =   "def"
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CheckBox ASM_CompileASMCode 
         Caption         =   "Compile ASM code"
         Height          =   250
         Left            =   -74760
         TabIndex        =   58
         Tag             =   "def"
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CheckBox ASM_FixASMListings 
         Caption         =   "* Fix ASM listings"
         Height          =   255
         Left            =   -74760
         TabIndex        =   57
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CheckBox C_UseCColors 
         Caption         =   "* Use C code coloring"
         Height          =   195
         Left            =   -74760
         TabIndex        =   56
         Tag             =   "def"
         Top             =   750
         Width           =   1935
      End
      Begin VB.CommandButton cmdColor_DefaultC 
         Caption         =   "Default"
         Height          =   255
         Left            =   -71520
         TabIndex        =   55
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdColor_AddC 
         Caption         =   "Add"
         Height          =   255
         Left            =   -74760
         TabIndex        =   53
         Top             =   2670
         Width           =   975
      End
      Begin VB.CommandButton cmdColor_DeleteC 
         Caption         =   "Delete"
         Height          =   255
         Left            =   -73800
         TabIndex        =   52
         Top             =   2670
         Width           =   975
      End
      Begin VB.TextBox DLL_EntryPoint 
         Height          =   285
         Left            =   -71520
         TabIndex        =   51
         Tag             =   "*"
         Top             =   3555
         Width           =   1000
      End
      Begin ThunderVB.ExportList DLL_ExpList 
         Height          =   2310
         Left            =   -74760
         TabIndex        =   50
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4075
      End
      Begin VB.CommandButton cmdColor_DeleteASM 
         Caption         =   "Delete"
         Height          =   255
         Left            =   -73680
         TabIndex        =   49
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdColor_AddASM 
         Caption         =   "Add"
         Height          =   255
         Left            =   -74760
         TabIndex        =   48
         Top             =   2520
         Width           =   1095
      End
      Begin ThunderVB.ColorList ASM_ASMColors 
         Height          =   1335
         Left            =   -74760
         TabIndex        =   47
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2566
      End
      Begin VB.CommandButton cmdColor_DefaultASM 
         Caption         =   "Default"
         Height          =   255
         Left            =   -71520
         TabIndex        =   46
         Top             =   690
         Width           =   975
      End
      Begin VB.CommandButton cmdINCFiles 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   45
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton cmdPaths_LIBFiles 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   44
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton cmdPaths_CCompiler 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   43
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox Paths_INCFiles 
         Height          =   315
         Left            =   120
         TabIndex        =   39
         Top             =   4920
         Width           =   4095
      End
      Begin VB.TextBox Paths_LIBFiles 
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   4320
         Width           =   4095
      End
      Begin VB.TextBox Paths_CCompiler 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   3720
         Width           =   4095
      End
      Begin VB.CheckBox DLL_LinkAsDLL 
         Caption         =   "* Create DLL"
         Height          =   375
         Left            =   -74640
         TabIndex        =   32
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox DLL_BaseAddress 
         Height          =   285
         Left            =   -73500
         TabIndex        =   31
         Tag             =   "*"
         Top             =   3550
         Width           =   1000
      End
      Begin VB.CheckBox DLL_ExportSymbols 
         Caption         =   "* Export functions"
         Height          =   375
         Left            =   -72600
         TabIndex        =   29
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Paths_ML 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox Paths_TextEditor 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   2220
         Width           =   4095
      End
      Begin VB.TextBox Paths_MIDL 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   1620
         Width           =   4095
      End
      Begin VB.TextBox Paths_Packer 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   2880
         Width           =   4095
      End
      Begin VB.CommandButton cmdPaths_ML 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   20
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdPaths_MIDL 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   19
         Top             =   1620
         Width           =   375
      End
      Begin VB.CommandButton cmdPaths_Packer 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   18
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdPaths_TextEditor 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   17
         Top             =   2220
         Width           =   375
      End
      Begin VB.CheckBox General_HideErrorDialogs 
         Caption         =   "* Hide error dialogs"
         Height          =   250
         Left            =   -74760
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox General_AddTLB 
         Caption         =   "Add TLB if needed"
         Height          =   250
         Left            =   -74760
         TabIndex        =   10
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox General_PopUpWindow 
         Caption         =   "* PopUp StdCall DLL window when compiling"
         Height          =   250
         Left            =   -74760
         TabIndex        =   9
         Top             =   1440
         Width           =   3975
      End
      Begin VB.CheckBox General_HookCompiler 
         Caption         =   "Hook compiler"
         Height          =   250
         Left            =   -74760
         TabIndex        =   8
         Tag             =   "def"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox Compile_ModifyCmdLine 
         Caption         =   "Modify CmdLine"
         Height          =   375
         Left            =   -74760
         TabIndex        =   7
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox Compile_PauseLink 
         Caption         =   "Pause before linking"
         Height          =   375
         Left            =   -74760
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox Compile_PauseAsm 
         Caption         =   "Pause before assembly"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox Debug_OutDebLog 
         Caption         =   "Enable Output to DebugLog"
         Height          =   375
         Left            =   -74760
         TabIndex        =   4
         Tag             =   "def"
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox ASM_UseASMColors 
         Caption         =   "* Use ASM code coloring"
         Height          =   195
         Left            =   -74760
         TabIndex        =   3
         Tag             =   "def"
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox Packer_UsePacker 
         Caption         =   "* Use packer"
         Height          =   375
         Left            =   -74760
         TabIndex        =   1
         Tag             =   "def"
         Top             =   780
         Width           =   1335
      End
      Begin ThunderVB.ColorList C_CColors 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   54
         Top             =   1110
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2566
      End
      Begin VB.Label lblpreVc 
         AutoSize        =   -1  'True
         Caption         =   "'#c' int i = 0x80 // a sample line"
         Height          =   195
         Left            =   -74760
         TabIndex        =   73
         Top             =   3240
         Width           =   2220
      End
      Begin VB.Label lblpreVasm 
         AutoSize        =   -1  'True
         Caption         =   "'#asm' mov eax , 12345 ; an example"
         Height          =   195
         Left            =   -74760
         TabIndex        =   72
         Top             =   3000
         Width           =   2625
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Use"
         Height          =   195
         Left            =   -70868
         TabIndex        =   62
         Top             =   1260
         Width           =   390
      End
      Begin VB.Label lblPacker_2 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   -72360
         TabIndex        =   61
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Path to .INC files"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   4680
         Width           =   1185
      End
      Begin VB.Label lblPaths_6 
         AutoSize        =   -1  'True
         Caption         =   "Path to .LIB files"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   4080
         Width           =   1155
      End
      Begin VB.Label lblPaths_5 
         AutoSize        =   -1  'True
         Caption         =   "Path to C compiler"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   1290
      End
      Begin VB.Label lblDLL_2 
         AutoSize        =   -1  'True
         Caption         =   "* Entry-Point"
         Height          =   195
         Left            =   -72450
         TabIndex        =   36
         Top             =   3600
         Width           =   870
      End
      Begin VB.Label lblDLL_1 
         AutoSize        =   -1  'True
         Caption         =   "* Base address  &&H"
         Height          =   195
         Left            =   -74880
         TabIndex        =   30
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label lblPaths_1 
         AutoSize        =   -1  'True
         Caption         =   "Path to ML.exe"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblPaths_2 
         AutoSize        =   -1  'True
         Caption         =   "Path to Text-Editor"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1980
         Width           =   1320
      End
      Begin VB.Label lblPaths_3 
         AutoSize        =   -1  'True
         Caption         =   "Path to MIDL.exe"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1380
         Width           =   1245
      End
      Begin VB.Label lblPaths_4 
         AutoSize        =   -1  'True
         Caption         =   "Path to Packer"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   1065
      End
      Begin VB.Label lblPacker_1 
         AutoSize        =   -1  'True
         Caption         =   "Packer Command-Line"
         Height          =   195
         Left            =   -74760
         TabIndex        =   2
         Top             =   1260
         Width           =   1605
      End
   End
   Begin VB.CheckBox DLL_UsePreLoader 
      Caption         =   "* Use ""Pre-Loader"""
      Height          =   195
      Left            =   1200
      TabIndex        =   75
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CheckBox DLL_FullLoading 
      Caption         =   "* Full loading"
      Height          =   195
      Left            =   1200
      TabIndex        =   76
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CheckBox DLL_DebugPreLoader 
      Caption         =   "* Debug ""Pre-Loader"""
      Height          =   195
      Left            =   1200
      TabIndex        =   77
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdGeneral_DeleteDebugDir 
      Caption         =   "Delete ""/debug"""
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   78
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdDebug_DeleteAllFiles 
      Caption         =   "Delete all files in ""/debug"""
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   79
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Frame fraDebug_Frame1 
      Caption         =   "Delete when UnLoading"
      Height          =   1455
      Left            =   360
      TabIndex        =   80
      Top             =   3840
      Width           =   2055
      Begin VB.CheckBox Debug_DeleteASM 
         Caption         =   ".ASM files"
         Height          =   375
         Left            =   240
         TabIndex        =   82
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Debug_DeleteLST 
         Caption         =   ".LST files"
         Height          =   375
         Left            =   240
         TabIndex        =   81
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CheckBox Debug_ForceLog 
      Caption         =   "Force log"
      Height          =   375
      Left            =   360
      TabIndex        =   83
      Top             =   2760
      Width           =   975
   End
   Begin VB.CheckBox Debug_OutMapFiles 
      Caption         =   "Output detailed MASM && LINK Map files"
      Height          =   375
      Left            =   360
      TabIndex        =   84
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CheckBox Debug_OutAsmToLog 
      Caption         =   "Output Assembler messages to log"
      Height          =   375
      Left            =   360
      TabIndex        =   85
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CheckBox Debug_DelDebBefCom 
      Caption         =   "Delete DebugLog before compiling"
      Height          =   375
      Left            =   360
      TabIndex        =   86
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CheckBox General_LogThunVBTools 
      Caption         =   "Log ThunderVB tools"
      Height          =   255
      Left            =   360
      TabIndex        =   87
      Top             =   3780
      Width           =   2055
   End
   Begin VB.CheckBox General_AddToMenu 
      Caption         =   "Add to menu"
      Height          =   255
      Left            =   360
      TabIndex        =   88
      Tag             =   "def"
      Top             =   3420
      Width           =   1455
   End
   Begin VB.CheckBox General_SetTopMost 
      Caption         =   "Set TopMost"
      Height          =   250
      Left            =   360
      TabIndex        =   89
      Tag             =   "def"
      Top             =   2100
      Width           =   1575
   End
   Begin VB.CheckBox General_LoadOnStartUp 
      Caption         =   "Load on StartUp"
      Height          =   250
      Left            =   360
      TabIndex        =   90
      Tag             =   "def"
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a part of the ThunderVB project.
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel) or Libor .
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'NOTE , THIS IS AN OLD VERSION RELEASED FOR TESTING
'   IT DATES 13/11/2004 [dd/mm/yyyy]


Option Explicit

'15.8. 2004 - initial version
'16.8. 2004 - basic GUI
'17.8. 2004 - Loading/Saving global settings, Debugging code
'20.8. 2004 - Local/Global/Default Values, added new Settings to StdCall DLL
'22.8. 2004 - StdCall DLL - new settings, Loading/Saving local settings
'25.8. 2004 - Paths, General - new settings, Default, save/load global/local settings code rewritten
'27.8. 2004 - new settings, SaveReg/VBP and ReadReg/VBP moved to this module from modPublic
'30.8. 2004 - many bugs fixed, added Export List, Color lists

'Raziel : fix on form loading , paths

'06.9. 2004 - better saving to VBP, killed bug when uset hits Storno
'07.9. 2004 - load settings in Load event, new Tabs - ASM, C
'12.9. 2004 - changed StdCall DLL tab - removed Patch preloader option, Partial loading changed to Full loading option
'           - added "Add DllMain" button, Settings remember last visible Tab
'13.9. 2004 - better "Add DllMain" button, using several packer command-lines
'16.9. 2004 - added "Show packer output", "Add to menu" option and "Log ThunderVB tools"

'TODO : bug in UserControl

'How to get Settings?
'--------------------
'If you want to load settings from registry and VBP only call function LoadSettings.
'It is in module modPublic.bas.
'
'From now you can use any Get_* function to read Settings
'
'it will try to load setting from VBP, if failed default will be used
'
'Global settings are common for all projects and local settings have to be loaded when
'user open new project
'
'How to load Global and Local settings?
'--------------------------------------
'is not supported
'
'Local (private)/Global (public) settings
'----------------------------------------
'*** CheckBox ***
'If Caption property of CheckBox control begins with LOCAL_VALUE (look at constant) then this item is local.
'Otherwise it is global.
'*** TextBox ***
'If Tag property of TextBox contains LOCAL_VALUE that this item is local.
'Otherwise it is global.
'
'Local settings will be saved in VBP file. Global settings will be saved in registry.

'Default values
'--------------
'*** CheckBox ***
'If Tag property of CheckBox contains DEFAULT_VALUE then this item will be set as default.
'*** TextBox ***
'TextBoxes do not have Default settings. Each TextBox will be cleared.

Private Const MSG_TITLE As String = "Settings"  'message box title
Private Const LOCAL_VALUE As String = "*"       'local setting
Private Const DEFAULT_VALUE As String = "def"   'default value

Private Const NO_STRING_DATA_VBP = "/-\"                 'when saving zero-length string to VBP save this string
Private Const PACKER_CMDLINE As String = "PackerCmdLine" 'name of entry in VBP - store selected cmd-line for Packer

'last selected tab
Private lLastTab As Long

'section in registry where Settings will be saved
Private Const REG_SECTION As String = "Settings"

Public Enum SET_SCOPE
    LOCAL_
    Global_
End Enum

'choose directory dialog
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'----------------------
'--- Control events ---
'----------------------

'set Default Settings
Private Sub cmdDefault_Click()
    Call SetDefaultSettings(Global_, True)
    Call SetDefaultSettings(LOCAL_, True)
End Sub

'save settings
Private Sub cmdOK_Click()

    On Error Resume Next
    
    'store settings to variables
    Call SaveSettingsToVariables

    'save to VBP and registry
    If SaveSettings(Global_) = False Then MsgBox "Error during saving " & Add34("global") & " Settings to registry.", vbExclamation, MSG_TITLE
    If SaveSettings(LOCAL_) = False Then MsgBox "Error during saving " & Add34("local") & " Settings to project file.", vbExclamation, MSG_TITLE
    
    'set visible tab
    lLastTab = sstSet.Tab
    
    Unload Me

End Sub

'Button Storno
Private Sub cmdStorno_Click()

    On Error Resume Next
    Unload Me
    
End Sub

'-------------------
'--- Form Events ---
'-------------------

Private Sub Form_Initialize()
    LogMsg "Loading " & Add34(Me.caption) & " window", Me.name, "Form_Initialize", True, True
End Sub

Private Sub Form_Terminate()
    LogMsg "Unloading " & Add34(Me.caption) & " window", Me.name, "Form_Terminate", True, True
End Sub

Private Sub Form_Load()
    
    'load settings
    Call LoadSettings(Global_, False, True)
    Call LoadSettings(LOCAL_, False, True)

    'save to variables
    Call SaveSettingsToVariables

    'default tab
    sstSet.Tab = lLastTab

    'common dialog settings
    With cdSet
        .Filter = "executable (*.exe)|*.exe|all (*.*)|*.*"
        .filename = ""
        .InitDir = App.path
        .flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNShareAware
        .CancelError = True
    End With

    'VB IDE - BUG??? always move Compile_PauseLink checkbox to somewhere so I manually move it back
    Rem Compile_PauseLink.Move 240, 1200, 2055, 375

End Sub

'-----------------------
'--- Tab StdCall DLL ---
'-----------------------

'check Base Address
Private Sub DLL_BaseAddress_Validate(Cancel As Boolean)

    'check string
    If Len(DLL_BaseAddress.Text) = 0 Then DLL_BaseAddress.Text = 0
    If IsNumeric("&H" & DLL_BaseAddress.Text) = False Then
        MsgBox "Invalid base address.", vbInformation, MSG_TITLE
        Cancel = True
    End If

End Sub

'show form
Private Sub cdmDLL_AddDllMain_Click()
    frmDllMainTemp.show vbModal, Me
End Sub

Private Sub DLL_EntryPoint_LostFocus()
    DLL_EntryPoint.Text = Trim(DLL_EntryPoint)
End Sub

'-----------------
'--- Tab Paths ---
'-----------------

Private Sub cmdPaths_MIDL_Click()
    Call SetPath("Path to " & Add34("MIDL.exe"), Paths_MIDL, "midl.exe")
End Sub

Private Sub cmdPaths_ML_Click()
    Call SetPath("Path to " & Add34("ML.exe"), Paths_ML, "ml.exe")
End Sub

Private Sub cmdPaths_Packer_Click()
    Call SetPath("Path to packer", Paths_Packer)
End Sub

Private Sub cmdPaths_TextEditor_Click()
    Call SetPath("Path to Text-Editor", Paths_TextEditor)
End Sub

Private Sub cmdPaths_CCompiler_Click()
    Call SetPath("Path to C compiler", Paths_CCompiler)
End Sub

Private Sub cmdPaths_LIBFiles_Click()
    Call SetDirectory("Select .LIB directory", Paths_LIBFiles)
End Sub

Private Sub cmdINCFiles_Click()
    Call SetDirectory("Select .INC directory", Paths_INCFiles)
End Sub

'append "\"
Private Sub Paths_INCFiles_Validate(Cancel As Boolean)
    If Right(Paths_INCFiles.Text, 1) <> "\" Then Paths_INCFiles.Text = Paths_INCFiles.Text & "\"
End Sub

'append "\"
Private Sub Paths_LIBFiles_Validate(Cancel As Boolean)
    If Right(Paths_LIBFiles.Text, 1) <> "\" Then Paths_LIBFiles = Paths_LIBFiles & "\"
End Sub

'-----------------------
'--- Tab ASM/C ---
'-----------------------

'set ASM default colors
Private Sub cmdColor_DefaultASM_Click()
    ASM_ASMColors.SetDefaultsAsm
End Sub

'set C default colors
Private Sub cmdColor_DefaultC_Click()
    C_CColors.SetDefaultsC
End Sub

'change ASM color
Private Sub ASM_ASMColors_ChangeColor(oldCol As Long, newcol As Long, bCancel As Boolean, bHandled As Boolean)

On Error Resume Next
    cdSet.ShowColor
    If Err.Number <> 0 Then
        bCancel = True
        Exit Sub
    End If
On Error GoTo 0
    
    newcol = cdSet.Color
    bCancel = False
    bHandled = True
    
End Sub

'change C color
Private Sub C_CColors_ChangeColor(oldCol As Long, newcol As Long, bCancel As Boolean, bHandled As Boolean)

On Error Resume Next
    
    cdSet.ShowColor
    If Err.Number <> 0 Then
        bCancel = True
        Exit Sub
    End If

On Error GoTo 0
    
    newcol = cdSet.Color
    bCancel = False
    bHandled = True

End Sub

'add new ASM color - BUG
Private Sub cmdColor_AddASM_Click()
    ASM_ASMColors.EditText ASM_ASMColors.AddColor("", vbBlack)
End Sub

'add new C color - BUG
Private Sub cmdColor_AddC_Click()
    C_CColors.EditText C_CColors.AddColor("", vbBlack)
End Sub

'delete ASM color
Private Sub cmdColor_DeleteASM_Click()
Dim oLB As ListBox
    
    Set oLB = ASM_ASMColors.listdata
    If oLB.ListIndex = -1 Then Exit Sub
    ASM_ASMColors.RemoveColor oLB.ListIndex
    
End Sub

'delete C color
Private Sub cmdColor_DeleteC_Click()
Dim oLB As ListBox
    
    Set oLB = C_CColors.listdata
    If oLB.ListIndex = -1 Then Exit Sub
    C_CColors.RemoveColor oLB.ListIndex
    
End Sub

'------------------------
'--- helper functions ---
'------------------------

'set path to the program
'- sDialogTitle - open dialog Title
'-txtTarget     - textbox where path will be stored
'-sAppName      - app name (eg. ml.exe or midl.exe)

Private Sub SetPath(sDialogTitle As String, txtTarget As TextBox, Optional sAppName As String = "")

    'set dialog title
    cdSet.DialogTitle = sDialogTitle
    cdSet.filename = ""
    
    'set new init directory
    If Len(txtTarget.Text) <> 0 Then cdSet.InitDir = Left(txtTarget.Text, InStrRev(txtTarget.Text, "\")) Else cdSet.InitDir = App.path & "\"
    
On Error Resume Next
    
    'select file
    cdSet.ShowOpen
    'cancel was pressed
    If Err.Number = 32755 Then Exit Sub
    
On Error GoTo 0
    
    'check predefined app name
    If Len(sAppName) = 0 Then GoTo 10
    
    'check filename
    If StrComp(Right(cdSet.filename, Len(sAppName)), sAppName, vbTextCompare) <> 0 Then
        MsgBox "Select " & Add34(sAppName) & " file.", vbInformation, "Settings"
    Else
10:
        'store path to the textbox
        txtTarget.Text = cdSet.filename
    End If
    
End Sub

'code from API-GUIDE
'parameters - sPrompt - prompt
'           - txtTarget - textbox, where path to directory will be stored
'KPD-Team 1998
'URL: http://www.allapi.net/
'KPDTeam@Allapi.net
Private Sub SetDirectory(ByVal sPrompt As String, txtTarget As TextBox)
Dim iNull As Integer, lpIDList As Long, lResult As Long
Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hwndOwner = Me.hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat(sPrompt, "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
            'append \
            If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        End If
    End If

    txtTarget.Text = sPath
    
End Sub

'---------------------------
'--- Save/Load functions ---
'---------------------------

'save all global setting to registry
'return - True (OK) /False (error)

Private Function SaveSettings(eScope As SET_SCOPE) As Boolean
Dim oControl As Object

On Error GoTo 10

    LogMsg "Saving settings - " & IIf(eScope = Global_, "global", "local"), Me.name, "SaveSettings", True, True

    'enum all controls
    For Each oControl In Me.Controls
        
        'if it is checkbox
        If TypeOf oControl Is CheckBox Then
            
            'check global flag
            If eScope = Global_ Then
                If Left(oControl.caption, 1) <> LOCAL_VALUE Then Call SaveReg(oControl.name, oControl.value)
            Else
                If Left(oControl.caption, 1) = LOCAL_VALUE Then Call SaveVBP(oControl.name, oControl.value)
            End If
            
        'it is textbox
        ElseIf TypeOf oControl Is TextBox Then
            
            'check global flag
            If eScope = Global_ Then
                If oControl.Tag <> LOCAL_VALUE Then Call SaveReg(oControl.name, Trim(oControl.Text))
            Else
                If oControl.Tag = LOCAL_VALUE Then Call SaveVBP(oControl.name, Trim(oControl.Text))
            End If
            
        End If
        
    Next oControl
    
    '--- other controls that must be saved ---
    
    If eScope = Global_ Then
        '--- global ---
        SaveReg ASM_ASMColors.name, ASM_ASMColors.ColorInfo
        SaveReg C_CColors.name, C_CColors.ColorInfo
    Else
        '--- local ---
        SaveVBP DLL_ExpList.name, DLL_ExpList.SelectedExports
        SaveVBP PACKER_CMDLINE, GetPackerOption
    End If
    
    'when saving to VBP, then upload VBP file
    If eScope = LOCAL_ Then VBI.ActiveVBProject.SaveAs VBI.ActiveVBProject.filename
    
    SaveSettings = True
    Exit Function
    
10:
    'error
    SaveSettings = False

    LogMsg "Error during saving settings - " & IIf(eScope = Global_, "global", "local"), Me.name, "SaveSettings", True, True

End Function

'load all global setting from registry
'parameters - bWarning - if error occurs then show msgbox
'           - bUseDefault - if true then set default
'return - True (OK) / False (error)

Private Function LoadSettings(eScope As SET_SCOPE, Optional bWarning As Boolean = False, Optional bUseDefault As Boolean = False) As Boolean
Dim oControl As Object
    
On Error GoTo 10

    LogMsg "Loading settings - " & IIf(eScope = Global_, "global", "local"), Me.name, "LoadSettings", True, True
    
    'enum all controls
    For Each oControl In Me.Controls
        
        'if it is checkbox
        If TypeOf oControl Is CheckBox Then
            
            'check global flag
            If eScope = Global_ Then
                If Left(oControl.caption, 1) <> LOCAL_VALUE Then oControl.value = ReadReg(oControl.name)
            Else
                If Left(oControl.caption, 1) = LOCAL_VALUE Then oControl.value = ReadVBP(oControl.name)
            End If
            
        ElseIf TypeOf oControl Is TextBox Then
                
            'check global flag
            If eScope = Global_ Then
                If oControl.Tag <> LOCAL_VALUE Then oControl.Text = Trim(ReadReg(oControl.name))
            Else
                If oControl.Tag = LOCAL_VALUE Then oControl.Text = Trim(ReadVBP(oControl.name))
            End If
            
        End If
        
    Next oControl
    
    '--- other controls that must be loaded ---
    
    If eScope = Global_ Then
        '--- global ---
        ASM_ASMColors.ColorInfo = ReadReg(ASM_ASMColors.name)
        C_CColors.ColorInfo = ReadReg(C_CColors.name)
    Else
        '--- local ---
        DLL_ExpList.SelectedExports = ReadVBP(DLL_ExpList.name)
        SetPackerOption CLng(ReadVBP(PACKER_CMDLINE))
    End If
    
    'return value
    LoadSettings = True
    Exit Function
    
10:
Dim sText As String
    
    LogMsg "Error during loading settings - " & IIf(eScope = Global_, "global", "local"), Me.name, "LoadSettings", True, True
    
    LoadSettings = False
    
    'make first error string
    If bWarning = True Then
        If eScope = Global_ Then sText = "Registry does not contain " & Add34("global") & " settings." Else sText = "This project does not contain " & Add34("local") & " settings ."
    End If
    
    If bUseDefault = True Then
        
        'make second error string
        sText = sText & " Default will be used."
        If bWarning = True Then MsgBox sText, vbExclamation, MSG_TITLE
        
        'set default values
        If eScope = Global_ Then SetDefaultSettings Global_, False Else SetDefaultSettings LOCAL_, False
    End If
        
End Function

'set Default Settings - Global settings
'bWarning - if True then question msgbox will appear

Private Sub SetDefaultSettings(eScope As SET_SCOPE, Optional bWarning As Boolean = True)
Dim oControl As Object

    LogMsg "Setting default settings - " & IIf(eScope = Global_, "global", "local"), Me.name, "SetDefaultSettings", True, True

    'at first - question
    If bWarning = True Then
        If eScope = Global_ Then
            If MsgBox("Are you shure to set all " & Add34("global") & " Settings to Default?", vbQuestion + vbYesNo, MSG_TITLE) = vbNo Then Exit Sub
        Else
            If MsgBox("Are you shure to set all " & Add34("local") & " Settings to Default?", vbQuestion + vbYesNo, MSG_TITLE) = vbNo Then Exit Sub
        End If
    End If
    
    'enum checkboxes
    For Each oControl In Me.Controls
        
        'if it is checkbox
        If TypeOf oControl Is CheckBox Then
            
            If eScope = Global_ Then
                'check global and default flag
                If Left(oControl.caption, 1) <> LOCAL_VALUE Then
                    If oControl.Tag = DEFAULT_VALUE Then oControl.value = 1 Else oControl.value = 0
                End If
            Else
                'check local and default flag
                If Left(oControl.caption, 1) = LOCAL_VALUE Then
                    If oControl.Tag = DEFAULT_VALUE Then oControl.value = 1 Else oControl.value = 0
                End If
            End If
            
        'textbox
        ElseIf TypeOf oControl Is TextBox Then
            
            If eScope = Global_ Then
                'check global flag
                If Left(oControl.Tag, 1) <> LOCAL_VALUE Then oControl.Text = ""
            Else
                'check local flag
                If Left(oControl.Tag, 1) = LOCAL_VALUE Then oControl.Text = ""
            End If
            
        End If
        
    Next oControl
    
    '--- other controls set to default ---
    
    If eScope = Global_ Then
        '--- global ---
        ASM_ASMColors.SetDefaultsAsm
        C_CColors.SetDefaultsC
    Else
        '--- local ---
        DLL_ExpList.SelectedExports = ""
        DLL_BaseAddress.Text = 0
        SetPackerOption 1
    End If
    
End Sub

'save settings to variables
Private Sub SaveSettingsToVariables()

    LogMsg "Saving settings to variables", Me.name, "SaveSettingsToVariables", True, True

    'Tab General
    Let_General LoadOnStartUp, General_LoadOnStartUp.value
    Let_General PopUpExportsWindow, General_PopUpWindow.value
    Let_General SetTopMost, General_SetTopMost.value
    Let_General HookCompiler, General_HookCompiler.value
    Let_General AddTlbToReferencesIfNeeded, General_AddTLB.value
    Let_General HideErrorDialogs, General_HideErrorDialogs.value
    Let_General AddToMenu, General_AddToMenu.value
    Let_General LogThunVBTools, General_LogThunVBTools.value
    
    'Tab Paths
    Let_Paths ml, Trim(Paths_ML.Text)
    Let_Paths TextEditor, Trim(Paths_TextEditor.Text)
    Let_Paths Midl, Trim(Paths_MIDL.Text)
    Let_Paths Packer, Trim(Paths_Packer.Text)
    Let_Paths INCFiles_Directory, Trim(Paths_INCFiles.Text)
    Let_Paths LIBFiles_Directory, Trim(Paths_LIBFiles.Text)
    Let_Paths CCompiler, Trim(Paths_CCompiler.Text)

    'Tab Packer
    Let_Packer UsePacker, Packer_UsePacker.value
    Let_Packer ShowPackerOutPut, Packer_ShowPackerOutPut.value
    Let_Packer CommandLine, Trim(GetPackerCmdLine)
    Let_Packer CmdLineDescription, Trim(GetPackerCmdLineDesc)

    'Tab C
    Let_C CColors, C_CColors.ColorInfo
    Let_C CompileCCode, C_CompileCCode.value
    Let_C UseCColoring, C_UseCColors.value
    
    'Tab ASM
    Let_ASM ASMColors, ASM_ASMColors.ColorInfo
    Let_ASM CompileASMCode, ASM_CompileASMCode.value
    Let_ASM FixASMListings, ASM_FixASMListings.value
    Let_ASM UseASMColoring, ASM_UseASMColors.value
    
    'init coloring engine
    'asm
    If Get_ASM(UseASMColoring) = True Then
        initAsmColors Get_ASM(ASMColors)
        AsmColoringEn True
    Else
        AsmColoringEn False
    End If
    'c
    If Get_C(UseCColoring) = True Then
        initCcolors Get_C(CColors)
        CColoringEn True
    Else
        CColoringEn False
    End If
    
    'Tab Debug
    Let_Debug EnableOutPutToDebugLog, Debug_OutDebLog.value
    Let_Debug DeleteDebugLogBeforeCompiling, Debug_DelDebBefCom.value
    Let_Debug OutPutAssemblerMessagesToLog, Debug_OutAsmToLog.value
    Let_Debug OutPutMapFiles, Debug_OutMapFiles.value
    Let_Debug ForceLog, Debug_ForceLog.value
    Let_Debug DeleteASM, Debug_DeleteASM.value
    Let_Debug DeleteLST, Debug_DeleteLST.value

    'Tab Compile
    Let_Compile PauseBeforeAssembly, Compile_PauseAsm.value
    Let_Compile PauseBeforeLinking, Compile_PauseLink.value
    Let_Compile ModifyCmdLine, Compile_ModifyCmdLine.value

    'Tab StdCall DLL
    Let_DLL LinkAsDll, DLL_LinkAsDLL.value
    Let_DLL ExportSymbols, DLL_ExportSymbols.value
    Let_DLL BaseAddress, IIf(Len(DLL_BaseAddress.Text) = 0, 0, "&H" & DLL_BaseAddress.Text)
    Let_DLL EntryPointName, Trim(DLL_EntryPoint.Text)
    Let_DLL ExportedSymbols, DLL_ExpList.SelectedExports
    'advanced
    Let_DLL UsePreLoader, DLL_UsePreLoader.value
    Let_DLL DebugPreLoader, DLL_DebugPreLoader.value
    Let_DLL FullLoading, DLL_FullLoading.value
    
End Sub

'read setting from registry
Private Function ReadReg(sKey As String) As String
    ReadReg = GetSetting(APP_NAME, REG_SECTION, sKey)
End Function

'save setting to registry
Private Sub SaveReg(sKey As String, sValue As String)
    SaveSetting APP_NAME, REG_SECTION, sKey, sValue
End Sub

'read setting from VBP
'note - problem when loading blank string to VBP
Private Function ReadVBP(sKey As String) As String
    ReadVBP = VBI.ActiveVBProject.ReadProperty(APP_NAME, sKey)
    If ReadVBP = NO_STRING_DATA_VBP Then ReadVBP = ""
End Function

'save setting to VBP
'note - problem when loading blank string to VBP
Private Sub SaveVBP(sKey As String, ByVal sValue As String)
    If Len(sValue) = 0 Then sValue = NO_STRING_DATA_VBP
    VBI.ActiveVBProject.WriteProperty APP_NAME, sKey, sValue
End Sub

'--- helper code for Tab Packer ---

'return selected packer cmdline
Private Function GetPackerCmdLine() As String

    If optPacker_1.value = True Then
        GetPackerCmdLine = Packer_CmdLine1.Text
    ElseIf optPacker_2.value = True Then
        GetPackerCmdLine = Packer_CmdLine2.Text
    ElseIf optPacker_3.value = True Then
        GetPackerCmdLine = Packer_CmdLine3.Text
    End If

End Function

'return description of selected cmdline
Private Function GetPackerCmdLineDesc() As String

    If optPacker_1.value = True Then
        GetPackerCmdLineDesc = Packer_Desc1.Text
    ElseIf optPacker_2.value = True Then
        GetPackerCmdLineDesc = Packer_Desc2.Text
    ElseIf optPacker_3.value = True Then
        GetPackerCmdLineDesc = Packer_Desc3.Text
    End If

End Function

'return selected cmd-line
Private Function GetPackerOption() As Long

    If optPacker_1.value = True Then
        GetPackerOption = 1
    ElseIf optPacker_2.value = True Then
        GetPackerOption = 2
    ElseIf optPacker_3.value = True Then
        GetPackerOption = 3
    End If

End Function

'set selected cmdline
Private Function SetPackerOption(lOption As Long)

    Select Case lOption
        Case 1: optPacker_1.value = True
        Case 2: optPacker_2.value = True
        Case 3: optPacker_3.value = True
    End Select
    
End Function


'-------------
'--- Debug ---
'-------------

'load settings without unloading form
Private Sub cmdDebugLoad_Click()

    MsgBox "Global " & LoadSettings(Global_)
    MsgBox "Local " & LoadSettings(LOCAL_)

End Sub

'save settings without unloading form
Private Sub cmdDebugSave_Click()

    Call SaveSettingsToVariables
    MsgBox "Global " & SaveSettings(Global_)
    MsgBox "Local " & SaveSettings(LOCAL_)

End Sub

'show local settings controls
Private Sub cmdDebugEnumLocal_Click()
Dim s As String, oControl As Object

    For Each oControl In Me.Controls
    
        If TypeOf oControl Is CheckBox Then
            If Left(oControl.caption, 1) = LOCAL_VALUE Then s = s & vbCrLf & "checkbox - " & oControl.name
        ElseIf TypeOf oControl Is TextBox Then
            If oControl.Tag = LOCAL_VALUE Then s = s & vbCrLf & "textbox - " & oControl.name
        End If
        
    Next oControl
    
    MsgBox "Local settings - controls names" & vbCrLf & s, vbInformation, "Debug"

End Sub

'show defualt settings control
Private Sub cmdDebugEnumDefault_Click()
Dim s As String, oControl As Object

    For Each oControl In Me.Controls
    
        If TypeOf oControl Is CheckBox Then
            If oControl.Tag = DEFAULT_VALUE Then s = s & vbCrLf & oControl.name
        End If
        
    Next oControl
    
    MsgBox "Default settings - controls names (CheckBoxes)" & vbCrLf & s, vbInformation, "Debug"

End Sub

Private Sub cmdDebugDeleteSettings_Click()
    DeleteSetting APP_NAME
End Sub
