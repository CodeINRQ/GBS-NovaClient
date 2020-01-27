VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucDSSRecGUI 
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8370
   ScaleHeight     =   465
   ScaleWidth      =   8370
   Begin VB.PictureBox picFrame 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   495
      HelpContextID   =   1090000
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.PictureBox picPlayer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   3135
         TabIndex        =   9
         Top             =   10
         Width           =   3135
         Begin VB.OptionButton optPlayer 
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton optPlayer 
            Height          =   375
            Index           =   1
            Left            =   375
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton optPlayer 
            Height          =   375
            Index           =   2
            Left            =   750
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton optPlayer 
            Height          =   375
            Index           =   3
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton optPlayer 
            Height          =   375
            Index           =   4
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton optPlayer 
            Height          =   375
            Index           =   5
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton optPlayer 
            Height          =   375
            Index           =   6
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton optPlayer 
            Height          =   375
            Index           =   7
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.OptionButton optInsert 
         Caption         =   "&Infoga"
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   8
         Tag             =   "1090101"
         ToolTipText     =   "Infoga diktat vid inspelning"
         Top             =   10
         Width           =   975
      End
      Begin VB.OptionButton optInsert 
         Caption         =   "&Ersätt"
         Height          =   195
         Index           =   1
         Left            =   3720
         TabIndex        =   7
         Tag             =   "1090102"
         ToolTipText     =   "Ersätt diktat vid inspelning"
         Top             =   250
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.PictureBox picEdit 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   2
         Top             =   10
         Width           =   495
         Begin VB.CommandButton cmdEdit 
            Enabled         =   0   'False
            Height          =   200
            Index           =   3
            Left            =   0
            Picture         =   "DSSRecGUI.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            Tag             =   "1090107"
            ToolTipText     =   "Avmarkera"
            Top             =   195
            Width           =   200
         End
         Begin VB.CommandButton cmdEdit 
            Height          =   200
            Index           =   2
            Left            =   195
            Picture         =   "DSSRecGUI.ctx":058A
            Style           =   1  'Graphical
            TabIndex        =   5
            Tag             =   "1090106"
            ToolTipText     =   "Markera slut"
            Top             =   0
            Width           =   200
         End
         Begin VB.CommandButton cmdEdit 
            Enabled         =   0   'False
            Height          =   200
            Index           =   1
            Left            =   195
            Picture         =   "DSSRecGUI.ctx":0B14
            Style           =   1  'Graphical
            TabIndex        =   4
            Tag             =   "1090108"
            ToolTipText     =   "Radera markerad del"
            Top             =   195
            Width           =   200
         End
         Begin VB.CommandButton cmdEdit 
            Height          =   200
            Index           =   0
            Left            =   0
            Picture         =   "DSSRecGUI.ctx":109E
            Style           =   1  'Graphical
            TabIndex        =   3
            Tag             =   "1090105"
            ToolTipText     =   "Markera start"
            Top             =   0
            Width           =   200
         End
      End
      Begin CareTalk.ucVUmeter ucVUmeter 
         Height          =   80
         Left            =   0
         TabIndex        =   1
         Top             =   400
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   132
      End
      Begin MSComctlLib.Slider sldPos 
         Height          =   315
         Left            =   4665
         TabIndex        =   18
         Top             =   190
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Max             =   100
         SelectRange     =   -1  'True
         TickStyle       =   1
         TickFrequency   =   20
      End
      Begin MSComctlLib.Slider sldVol 
         Height          =   255
         Left            =   5880
         TabIndex        =   19
         Top             =   190
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   20
         SmallChange     =   5
         Max             =   100
         SelStart        =   50
         TickStyle       =   1
         TickFrequency   =   20
         Value           =   50
      End
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   255
         Left            =   7125
         TabIndex        =   20
         Top             =   190
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   17
         SmallChange     =   5
         Min             =   15
         Max             =   100
         SelStart        =   50
         TickStyle       =   1
         TickFrequency   =   20
         Value           =   50
      End
      Begin VB.Label lblLength 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   4545
         TabIndex        =   23
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label lblVolume 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Volym"
         Height          =   255
         Left            =   5880
         TabIndex        =   22
         Tag             =   "1090103"
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Hastighet"
         Height          =   255
         Left            =   7125
         TabIndex        =   21
         Tag             =   "1090104"
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1920
      Top             =   600
   End
   Begin MSComctlLib.ImageList ilButtons16 
      Left            =   120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":1628
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":19C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":1F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":22F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":2690
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":2A2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":2DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":335E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":38F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilButtons32 
      Left            =   720
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":3E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":476C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":5046
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":5920
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":61FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":6AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":73AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":7C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DSSRecGUI.ctx":8562
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgCurrentIcon 
      Height          =   375
      Left            =   1440
      Top             =   600
      Width           =   375
      Visible         =   0   'False
   End
End
Attribute VB_Name = "ucDSSRecGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents DSSRec As CareTalkDSSRec3.DSSRecorder
Attribute DSSRec.VB_VarHelpID = -1
Private NowPlayingFilename As String
Private InPositioning As Boolean
Private LastState As Gru_State
Private InGruEventHandler As Boolean

Private mReadOnly As Boolean

Public Event PosChange(PosInMilliSec As Long, LengthInMilliSec As Long, Formated As String)
Public Event ChangeIcon(NewIcon As Image)

Public Enum PlayerButEnum
   butPlay = 0
   butPause = 1
   butStop = 2
   butStart = 3
   butRewind = 4
   butForward = 5
   butEnd = 6
   butRec = 7
   butNone = 200
End Enum

Private Const icoRewind = 1
Private Const icoPause = 2
Private Const icoPlay = 3
Private Const icoStart = 4
Private Const icoEnd = 5
Private Const icoForward = 6
Private Const icoStop = 7
Private Const icoRecDark = 8
Private Const icoRecLight = 9

Private Const editStart = 0
Private Const editDelete = 1
Private Const editEnd = 2
Private Const editClear = 3

Public Sub ExternalButton(B As PlayerButEnum)

   If optPlayer(CInt(B)).Enabled And optPlayer(CInt(B)).Visible Then
      optPlayer(CInt(B)).Value = True
   End If
End Sub
Public Sub ExternalVolumeChange(Increase As Boolean)

   Dim NewValue As Integer
   
   If Increase Then
      NewValue = sldVol.Value + sldVol.LargeChange
   Else
      NewValue = sldVol.Value - sldVol.LargeChange
   End If
   If NewValue < sldVol.Min Then
      NewValue = sldVol.Min
   ElseIf NewValue > sldVol.Max Then
      NewValue = sldVol.Max
   End If
   sldVol.Value = NewValue
   sldVol_Scroll
End Sub
Public Sub ExternalSpeedChange(Increase As Boolean)

   Dim NewValue As Integer
   
   If Increase Then
      NewValue = sldSpeed.Value + sldSpeed.LargeChange
   Else
      NewValue = sldSpeed.Value - sldSpeed.LargeChange
   End If
   If NewValue < sldSpeed.Min Then
      NewValue = sldSpeed.Min
   ElseIf NewValue > sldSpeed.Max Then
      NewValue = sldSpeed.Max
   End If
   sldSpeed.Value = NewValue
   sldSpeed_Scroll
End Sub
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub

Private Sub InitPlayerButtons()

   picPlayer.BackColor = BackColor
   With ilButtons16
      optPlayer(butPlay).Picture = .ListImages(icoPlay).Picture
      optPlayer(butPause).Picture = .ListImages(icoPause).Picture
      optPlayer(butStop).Picture = .ListImages(icoStop).Picture
      optPlayer(butStart).Picture = .ListImages(icoStart).Picture
      optPlayer(butRewind).Picture = .ListImages(icoRewind).Picture
      optPlayer(butForward).Picture = .ListImages(icoForward).Picture
      optPlayer(butEnd).Picture = .ListImages(icoEnd).Picture
      optPlayer(butRec).Picture = .ListImages(icoRecDark).Picture
   End With
End Sub
Private Sub InitEditButtons()

   picEdit.BackColor = BackColor
End Sub
Private Sub SetNewIcon(NewIcon As Picture)

   imgCurrentIcon.Picture = NewIcon
   RaiseEvent ChangeIcon(imgCurrentIcon)
End Sub

Private Sub cmdEdit_Click(Index As Integer)

   Dim Pos As Long
   Dim Length As Long
   Static Selstart As Long    'in ms
   Static Selend As Long      'in ms
   
   DSSRec.GetPos Pos
   DSSRec.GetLength Length
   
   Select Case Index
      Case editStart
         Selstart = Pos
         If Pos >= Selend Then
            Selend = Length
         End If
      Case editEnd
         Selend = Pos
         If Pos <= Selstart Then
            Selstart = 0
         End If
      Case editClear
         Selstart = 0
         Selend = 0
      Case editDelete
         If Selend > Selstart Then
            DSSRec.Delete Selstart, Selend
            Selstart = 0
            Selend = 0
         End If
         UpdatePos -1
   End Select
   If Selstart <= Selend Then
      sldPos.Selstart = ConvertTo100Scale(Selstart, Length)
      sldPos.SelLength = ConvertTo100Scale(Selend - Selstart, Length)
   End If
   If picEdit.Enabled Then
      cmdEdit(editClear).Enabled = Selstart <> 0 Or Selend <> 0
      cmdEdit(editDelete).Enabled = Selstart < Selend
   End If
End Sub
Private Function ConvertTo100Scale(PosInMilliSec As Long, LengthInMilliSec) As Integer

   ConvertTo100Scale = CInt(PosInMilliSec / LengthInMilliSec * CLng(100))
End Function


Private Sub lblSpeed_Click()

   sldSpeed.Value = 50
   sldSpeed_Scroll
End Sub

Private Sub lblVolume_Click()

   sldVol.Value = 50
   sldVol_Scroll
End Sub

Private Sub optInsert_Click(Index As Integer)

   Dim M As Gru_RecMode

   If optInsert(0).Value Then
      M = GRU_INSERT
   Else
      M = GRU_OVERWRITE
   End If
   DSSRec.SetRecordMode M
End Sub

Private Sub optPlayer_Click(Index As Integer)

   Dim L As Long

   If Not InGruEventHandler Then
      Select Case Index
         Case butPlay
            DSSRec.Play
         Case butPause
            DSSRec.PlayPause
         Case butStop
            DSSRec.PlayStop
         Case butStart
            DSSRec.MoveTo 0
            optPlayer(butPause).Value = True
            DSSRec.PlayPause
         Case butRewind
            DSSRec.Rewind
         Case butForward
            DSSRec.FastForward
         Case butEnd
            DSSRec.GetLength L
            DSSRec.MoveTo L
            optPlayer(butPause).Value = True
            DSSRec.PlayPause
         Case butRec
            DSSRec.Rec
      End Select
   End If
   If LastState = GRU_RECPAUSED Then
      SetNewIcon optPlayer(butRec).Picture
   Else
      SetNewIcon optPlayer(Index).Picture
   End If
End Sub
Private Sub DSSRec_GruEvent(EventType As CareTalkDSSRec3.Gru_Event, Data As Long)

   Dim S As String
   
   InGruEventHandler = True
   Select Case EventType
      Case GRU_POSCHANGE
         If InPositioning Then
            InGruEventHandler = False
            Exit Sub
         End If
         UpdatePos Data
      Case GRU_STATECHANGED
         LastState = Data
         Select Case LastState
            Case GRU_STOPPED
               optPlayer(butPause).Value = True
            Case GRU_PLAY
               optPlayer(butPlay).Value = True
            Case GRU_RECPAUSED
               optPlayer(butPause).Value = True
            Case GRU_REC
               optPlayer(butRec).Value = True
            Case GRU_REWIND
               optPlayer(butRewind).Value = True
            Case GRU_FORWARD
               optPlayer(butForward).Value = True
         End Select
         tmrBlink.Enabled = LastState = GRU_RECPAUSED
         If LastState = GRU_REC Or LastState = GRU_RECPAUSED Then
            optPlayer(butRec).Picture = ilButtons16.ListImages(icoRecLight).Picture
            SetNewIcon optPlayer(butRec).Picture
         Else
            If Client.SysSettings.PlayerAutoOverwrite Then
               optInsert(1).Value = True
            End If
            optPlayer(butRec).Picture = ilButtons16.ListImages(icoRecDark).Picture
         End If
      Case GRU_BUTTONPRESS
         If Data = GRU_BUT_INDEX Then
            If Client.SysSettings.PlayerIndexButtonAsInsertRec Then
               optInsert(0).Value = True
               DSSRec.Rec
            End If
         ElseIf Data = GRU_BUT_INSERT Then
            optInsert(0).Value = True
            DSSRec.Rec
         End If
      Case GRU_INPUTCHANGE
         ucVUmeter.Value = Data
   End Select
   InGruEventHandler = False
End Sub
Private Sub UpdatePos(ByVal Pos As Long)

   Dim L As Long
   
   If Pos < 0 Then
      DSSRec.GetPos Pos
   End If
   DSSRec.GetLength L
   
   If L > 0 Then
      sldPos.Value = (Pos / L) * 100
   Else
      sldPos.Value = 0
   End If
   ShowPos Pos, L
End Sub
Private Sub ShowPos(PosInMilliSec As Long, LenInMilliSec As Long)

   Dim S As String
   
   S = FormatPos(PosInMilliSec, LenInMilliSec)
   lblLength.Caption = S
   RaiseEvent PosChange(PosInMilliSec, LenInMilliSec, S)
End Sub
Private Function FormatPos(PosInMilliSec As Long, LenInMilliSec As Long) As String

   FormatPos = FormatLength(PosInMilliSec / 1000) & " / " & FormatLength(LenInMilliSec / 1000)
End Function
Private Sub sldPos_Scroll()

   Dim L As Long
   
   If InPositioning Then Exit Sub
   InPositioning = True
   DSSRec.GetLength L
   DSSRec.MoveTo CLng(sldPos.Value * L / CLng(100))
   InPositioning = False
End Sub

Private Sub sldSpeed_Scroll()

   DSSRec.SetPlaySpeed (sldSpeed.Value + 50) * 10
End Sub

Private Sub sldVol_Scroll()

   DSSRec.SetPlayBackVolume CInt(sldVol.Value * 0.8) + 20
End Sub

Public Sub OpenAndPlay(Filename As String)

   InitPlayerBeforeUse
   OpenFile Filename
   'DSSRec.Play
End Sub

Private Sub OpenFile(Filename As String)

   Dim L As Long
   
   If Len(Filename) > 0 And UCase$(Filename) <> UCase$(NowPlayingFilename) Then
      DSSRec.CloseFile
      If DSSRec.LoadFile(Filename, CInt(mReadOnly), CInt(False)) = 0 Then
         NowPlayingFilename = Filename
         EnableControls True
      End If
   End If
   DSSRec.GetLength L
   ShowPos 0, L
End Sub
Public Sub CreateNewFile(Filename As String)

   Trc "ucDSS CreateNewFile", ""
   InitPlayerBeforeUse
   If Len(Filename) > 0 And UCase$(Filename) <> UCase$(NowPlayingFilename) Then
      DSSRec.CloseFile
      mReadOnly = False
      If DSSRec.LoadFile(Filename, CInt(mReadOnly), CInt(True)) = 0 Then
         NowPlayingFilename = Filename
         EnableControls True
      End If
      DSSRec.Rec
   End If
End Sub
Private Sub InitPlayerBeforeUse()

   Dim I As Integer
   Dim Speed As Integer
   
   picEdit.Visible = Client.SysSettings.PlayerShowEditButtons
   
   DSSRec.SetRecordMode GRU_OVERWRITE
   optInsert(1).Value = True
   
   DSSRec.SetWindingSpeed 8000
   DSSRec.SetBackspace 2000
   
   DSSRec.GetPlaySpeed Speed
   sldSpeed.Value = Speed / 10 - 50
   
   DSSRec.GetPlayBackVolume I
   If I < 20 Then
      I = 20
   End If
   sldVol.Value = (I - 20) / 0.8
End Sub

Public Sub StopAndClose()
  
   DSSRec.PlayStop
   DSSRec.CloseFile
   EnableControls False
End Sub
Private Sub EnableControls(Value As Boolean)

   Dim I As Integer

   For I = butPlay To butEnd
      optPlayer(I).Enabled = Value
   Next I
   optPlayer(butRec).Enabled = Not mReadOnly And Client.Hw = GRU_HW_RECORD
   optInsert(0).Enabled = Not mReadOnly And Client.Hw = GRU_HW_RECORD
   optInsert(1).Enabled = Not mReadOnly And Client.Hw = GRU_HW_RECORD
   picEdit.Enabled = Not mReadOnly And Client.Hw = GRU_HW_RECORD
   sldPos.Enabled = Value
   If Not Value Then
      sldPos.Value = 0
   End If
End Sub

Public Property Set DSSRecorder(Rec As CareTalkDSSRec3.DSSRecorder)

   Trc "ucDSS set DssRecorder", Format$(Rec Is Nothing)
   Set DSSRec = Rec
End Property
Public Property Let ReadOnly(Value As Boolean)

   mReadOnly = Value
End Property
Public Property Get Dirty() As Boolean

   Dirty = DSSRec.Dirty
End Property
Public Property Get SoundLengthInSec() As Long

   Dim SLen As Long
   
   DSSRec.GetLength SLen
   SoundLengthInSec = CLng(SLen / 1000)
End Property
Private Sub tmrBlink_Timer()

   Static Dark As Boolean
   
   If Dark Then
      optPlayer(butRec).Picture = ilButtons16.ListImages(icoRecLight).Picture
      Dark = False
   Else
      optPlayer(butRec).Picture = ilButtons16.ListImages(icoRecDark).Picture
      Dark = True
   End If
   SetNewIcon optPlayer(butRec).Picture
End Sub

Private Sub UserControl_Initialize()

   picFrame.BackColor = BackColor
   Trc "ucDSS Initialize", ""
   ucVUmeter.Value = 0
   InitPlayerButtons
   InitEditButtons
End Sub

Private Sub UserControl_Terminate()

   Trc "ucDSS Terminate", ""
End Sub
