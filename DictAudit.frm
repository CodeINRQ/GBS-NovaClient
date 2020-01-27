VERSION 5.00
Begin VB.Form frmDictAudit 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Spårning användning av diktat"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9480
   HelpContextID   =   1360000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "1360100"
   Begin CareTalk.ucAuditList ucAuditList 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7435
   End
End
Attribute VB_Name = "frmDictAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DictId As Long
Public UserId As Long

Private Sub Form_Load()

   CenterAndTranslateForm Me, frmMain
   ucAuditList.DictId = DictId
   ucAuditList.UserId = UserId
   ucAuditList.RestoreSettings ""
   ucAuditList.GetDataNow
End Sub

Private Sub Form_Resize()

   ucAuditList.Move 0, 0, Me.Width, Me.Height
End Sub
