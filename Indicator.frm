VERSION 5.00
Begin VB.Form frmIndicator 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleWidth      =   1425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIndicator 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   3975
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.Timer tmrLight 
         Interval        =   100
         Left            =   840
         Top             =   0
      End
      Begin VB.Shape shpIndicator 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   150
         Left            =   100
         Shape           =   3  'Circle
         Top             =   30
         Width           =   150
      End
      Begin VB.Label lblCurrent 
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   10
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLightTime As Long
Private mTimeCounter As Long
Private mDarkTime As Long
Private mLightNow As Boolean
Private mBlinkNumber As Long
Private mBlinkMaxNumber As Long
Private mId As String
Private mTxt As String

Public Sub SetIndicatorText(Txt As String, Id As String)

   If mTxt <> Txt Then
      mTxt = Txt
      lblCurrent.Caption = mTxt
      If Len(Txt) > 0 Then
         Me.Visible = True
         WindowFloating Me
      Else
         'WindowNotFloating Me
         Me.Visible = False
      End If
   End If
   If Id <> mId Then
      mId = Id
      If Len(mTxt) > 0 Then
            StartBlink
      End If
   End If
End Sub

Private Sub Form_Load()

   Me.BackColor = Client.SysSettings.IndicatorBgColor
   Me.picIndicator.BackColor = Me.BackColor
   Me.shpIndicator.FillColor = Client.SysSettings.IndicatorLightColor
   mLightTime = Client.SysSettings.IndicatorLightTime
   mDarkTime = Client.SysSettings.IndicatorDarkTime
   mBlinkMaxNumber = Client.SysSettings.IndicatorMaxBlink
End Sub

Private Sub lblCurrent_Click()

   If frmMain.Visible Then
      SetWindowTopMostAndForeground frmMain
      WindowFloating Me
   End If
End Sub

Private Sub tmrLight_Timer()

   mTimeCounter = mTimeCounter - 1
   If mTimeCounter < 0 Then
      If mLightNow Then
         StartDark
      Else
         StartLight
      End If
   End If
End Sub

Private Sub StartBlink()

   mBlinkNumber = 0
   StartLight
End Sub
Private Sub StartLight()

   mTimeCounter = mLightTime
   mLightNow = True
   mBlinkNumber = mBlinkNumber + 1
   If (mBlinkNumber <= mBlinkMaxNumber Or mBlinkMaxNumber < 0) And mLightTime > 0 Then
      shpIndicator.FillStyle = 0
   End If
End Sub
Private Sub StartDark()

   mTimeCounter = mDarkTime
   mLightNow = False
   shpIndicator.FillStyle = 1
End Sub


