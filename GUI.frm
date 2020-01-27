VERSION 5.00
Begin VB.Form frmGUI 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin CareTalkClient.ucDSSRecGUI ucDSSRecGUI1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   873
   End
End
Attribute VB_Name = "frmGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)

   ucDSSRecGUI1.StopAndClose
End Sub

