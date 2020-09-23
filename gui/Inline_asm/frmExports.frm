VERSION 5.00
Begin VB.Form frmExports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Exports"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin ThunderVB.ExportList ExportList1 
      Height          =   3210
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      _extentx        =   8281
      _extenty        =   5662
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
   End
End
Attribute VB_Name = "frmExports"
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

'Revision history:
'19/8/2004[dd/mm/yyyy] : Created by Raziel
'Form Created , intial version
'Code changed to use ExportList Control

'this is obsotele..

Private Sub cmdOK_Click()
    Dim temp() As String
    MsgBox "still not ready.."
    'temp = exports.SelectedExports
End Sub

