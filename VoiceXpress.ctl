VERSION 5.00
Begin VB.UserControl ucVoiceXpress 
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   ScaleHeight     =   3300
   ScaleWidth      =   9570
   Begin VB.Frame fraVx 
      Caption         =   "voiceXpress"
      Height          =   3135
      HelpContextID   =   1170000
      Left            =   0
      TabIndex        =   0
      Tag             =   "1170101"
      Top             =   0
      Width           =   8175
      Begin VB.Label lblVxListening 
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Mikrofon:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Tag             =   "1170106"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblVxVisible 
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Synligt:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Tag             =   "1170104"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblVxAppState 
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   1440
         Width           =   5895
      End
      Begin VB.Label Label4 
         Caption         =   "Status:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Tag             =   "1170105"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblVxRunning 
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Startat:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Tag             =   "1170103"
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblVxInstalled 
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Installerat:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Tag             =   "1170102"
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Timer tmrVx 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8880
      Top             =   120
   End
End
Attribute VB_Name = "ucVoiceXpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mVx As clsVoiceXpress
Attribute mVx.VB_VarHelpID = -1
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub

Public Sub Init(Vx As clsVoiceXpress)

   Set mVx = Vx
   ShowAll
   tmrVx.Enabled = True
End Sub
Private Sub ShowAll()

   lblVxInstalled.Caption = ShowValueBoolean(mVx.VxInstalled)
   lblVxRunning.Caption = ShowValueBoolean(mVx.VxRunning)
   lblVxAppState.Caption = ShowValueAppState(mVx.VxAppState)
   lblVxVisible.Caption = ShowValueBoolean(mVx.VxVisible)
   lblVxListening.Caption = ShowValueListening(mVx.vxListening)
End Sub
Private Function ShowValueBoolean(Value As Boolean) As String

   If Value Then
      ShowValueBoolean = Client.Texts.Txt(1170107, "Ja")
   Else
      ShowValueBoolean = Client.Texts.Txt(1170108, "Nej")
   End If
End Function
Private Function ShowValueAppState(Value As vxAppStateEnum) As String

   ShowValueAppState = mVx.StatusText
End Function
Private Function ShowValueListening(Value As vxListeningEnum) As String

   Select Case Value
      Case vxListeningOff:       ShowValueListening = Client.Texts.Txt(1170109, "Av")
      Case vxListeningWakeUp:    ShowValueListening = Client.Texts.Txt(1170110, "Av, vaknar på kommando")
      Case vxListeningOn:        ShowValueListening = Client.Texts.Txt(1170111, "På")
      Case Else:                 ShowValueListening = ""
   End Select
End Function

Private Sub mVx_ChangeAppState(NewValue As vxAppStateEnum)

   lblVxAppState = ShowValueAppState(NewValue)
End Sub

Private Sub mVx_ChangeListening(NewValue As vxListeningEnum)

   ShowAll
End Sub

Private Sub tmrVx_Timer()

   ShowAll
End Sub

Private Sub UserControl_Terminate()

   tmrVx.Enabled = False
   Set mVx = Nothing
End Sub
