VERSION 5.00
Begin VB.Form frmTraceOrg 
   Caption         =   "Grundig Trace"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   Icon            =   "Trace.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstTrace 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmTraceOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()

   On Error Resume Next
   lstTrace.Height = Me.Height - lstTrace.Top - 450
   lstTrace.Width = Me.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)

   WindowNotFloating Me
End Sub
