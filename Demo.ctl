VERSION 5.00
Begin VB.UserControl ucDemo 
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8445
   ScaleHeight     =   3255
   ScaleWidth      =   8445
   Begin VB.Frame fraDemo 
      Caption         =   "Demo"
      Height          =   3015
      HelpContextID   =   1070000
      Left            =   120
      TabIndex        =   0
      Tag             =   "1070101"
      Top             =   120
      Width           =   8175
      Begin VB.TextBox txtMinPerDay 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Text            =   "10"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtMaxPerDay 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "100"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtDaysInHistory 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Text            =   "365"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtNumberOfDemoDictations 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "200"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdDemo 
         Caption         =   "Fyll databas med demodiktat"
         Height          =   375
         Left            =   4680
         TabIndex        =   1
         Tag             =   "1070103"
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Min antal/dag i historik:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Tag             =   "1070111"
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "Max antal/dag i historik:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Tag             =   "1070110"
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Dagar i historik:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Tag             =   "1070109"
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Antal demodiktat:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Tag             =   "1070102"
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "ucDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event UIStatusSet(StatusText As String, Busy As Boolean)
Public Event UIStatusSetSub(SubText As String)
Public Event UIStatusProgress(Total As Long, Left As Long)
Public Event UIStatusClear()

Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub
Private Sub cmdDemo_Click()

   Dim NumDict As Long
   Dim DaysHist As Integer
   Dim MaxPerDay As Integer
   Dim MinPerDay As Integer
   
   On Error Resume Next
   NumDict = CLng(txtNumberOfDemoDictations.Text)
   'If NumDict > MaxNumberOfDictation Then
   '   NumDict = MaxNumberOfDictation
   'End If
   DaysHist = CInt(txtDaysInHistory.Text)
   MaxPerDay = CInt(txtMaxPerDay.Text)
   MinPerDay = CInt(txtMinPerDay.Text)
   On Error GoTo 0
      

   If MsgBox(Client.Texts.Txt(1070104, "Om du fortsätter kommer alla befinliga diktat och all historik att raderas!"), vbOKCancel) = vbOK Then
      RaiseEvent UIStatusSet(Client.Texts.Txt(1070105, "Generering av demo"), True)
      
         RaiseEvent UIStatusSet(Client.Texts.Txt(1070106, "Tidigare diktat raderas"), True)
         Client.Server.DeleteAllDictations
         RaiseEvent UIStatusClear
        
         RaiseEvent UIStatusSet(Client.Texts.Txt(1070107, "Tidigare historik raderas"), True)
         Client.Server.DeleteHistory 0
         RaiseEvent UIStatusClear
        
         RaiseEvent UIStatusSet(Client.Texts.Txt(1070108, "Nya diktat och ny historik genereras"), True)
         FillDemoDictation NumDict
         FillHistoryDemo DaysHist, MaxPerDay, MinPerDay
         RaiseEvent UIStatusClear
      
         Client.LoggMgr.Insert 1320109, LoggLevel_SysAdmin, 0, txtNumberOfDemoDictations.Text
         
      RaiseEvent UIStatusClear
   End If
End Sub

Private Sub txtDaysInHistory_Change()

   SetEnabled
End Sub

Private Sub txtMaxPerDay_Change()

   SetEnabled
End Sub

Private Sub txtMinPerDay_Change()

   SetEnabled
End Sub

Private Sub txtNumberOfDemoDictations_Change()

   SetEnabled
End Sub

Private Sub txtNumberOfDemoDictations_KeyPress(KeyAscii As Integer)

   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii < 32 Or KeyAscii = 45) Then
      KeyAscii = 0
   End If
End Sub
Private Sub SetEnabled()

   Dim B As Boolean
   
   B = Len(txtNumberOfDemoDictations.Text) > 0
   B = B And Len(txtDaysInHistory.Text) > 0
   B = B And Len(txtMaxPerDay.Text) > 0
   B = B And Len(txtMinPerDay.Text) > 0
   cmdDemo.Enabled = B And Client.SysSettings.DemoAllowGenerate
End Sub
