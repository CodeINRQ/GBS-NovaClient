VERSION 5.00
Begin VB.Form frmDict 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diktat"
   ClientHeight    =   4575
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8760
   HelpContextID   =   1030000
   Icon            =   "Dict.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8760
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1030100"
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkChangeDict 
      Height          =   270
      Left            =   5160
      Picture         =   "Dict.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   840
      Width           =   270
   End
   Begin VB.TextBox txtTxt 
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   11
      Top             =   4200
      Width           =   6255
   End
   Begin VB.ComboBox cboPriority 
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.ComboBox cboDictType 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CheckBox chkNoPatient 
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtPatName 
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txtPatId 
      Height          =   285
      Left            =   2400
      MaxLength       =   14
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin CareTalk.ucCloseChoice ucCloseChoice 
      Height          =   1695
      HelpContextID   =   1030000
      Left            =   6120
      TabIndex        =   12
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
   End
   Begin CareTalk.ucDSSRecGUI ucDSSRecGUI 
      Height          =   495
      HelpContextID   =   1030000
      Left            =   120
      TabIndex        =   15
      Top             =   60
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   873
   End
   Begin CareTalk.ucOrgTree ucOrgTree 
      Height          =   3855
      HelpContextID   =   1030000
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5953
   End
   Begin VB.Label lblChangeDict 
      BackStyle       =   0  'Transparent
      Caption         =   "�ndra:"
      Height          =   255
      Left            =   5160
      TabIndex        =   33
      Tag             =   "1030117"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblTxtTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Nyckelord:"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Tag             =   "1030116"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblOrgMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2280
      TabIndex        =   31
      Top             =   600
      Width           =   135
   End
   Begin VB.Label lblPatNameMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   30
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lblPatIdMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   29
      Top             =   840
      Width           =   135
   End
   Begin VB.Image imgLess 
      Height          =   480
      Left            =   8565
      Picture         =   "Dict.frx":06C0
      Top             =   405
      Width           =   480
   End
   Begin VB.Image imgMore 
      Height          =   480
      Left            =   8565
      Picture         =   "Dict.frx":0F8A
      Top             =   405
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgPinin 
      Height          =   210
      Left            =   8520
      Picture         =   "Dict.frx":1854
      Tag             =   "1030115"
      ToolTipText     =   "Alltid �verst"
      Top             =   120
      Width           =   210
      Visible         =   0   'False
   End
   Begin VB.Image imgPin 
      Height          =   210
      Left            =   8520
      Picture         =   "Dict.frx":192F
      Tag             =   "1030114"
      ToolTipText     =   "Normalt f�nster"
      Top             =   120
      Width           =   210
   End
   Begin VB.Label lblTranscribedDateTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Utskrivet:"
      Height          =   255
      Left            =   6120
      TabIndex        =   28
      Tag             =   "1030113"
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblTranscribedDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblStatusTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Tag             =   "1030106"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblExpiryDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   24
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblExpiryDateTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Utskrift senast:"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Tag             =   "1030108"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblTranscriber 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblTranscriberTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Utskrivare:"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Tag             =   "1030112"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblAuthorTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Intalare:"
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Tag             =   "1030110"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblPriorityTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Prioritet:"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Tag             =   "1030107"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Diktattyp:"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Tag             =   "1030105"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblChanged 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblChangedTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "�ndrat:"
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Tag             =   "1030111"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblCreated 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblCreatedTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Inl�st:"
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Tag             =   "1030109"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblNoPatientTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingen patient:"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Tag             =   "1030103"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblPatnameTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Patientens namn:"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Tag             =   "1030104"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblPatIdTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Patientens personnr:"
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Tag             =   "1030102"
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long

Public Event CloseChoiceSelected(Index As Integer)

Public mForceUnload As Boolean

Private WithEvents DSSRecorder As CareTalkDSSRec3.DSSRecorder
Attribute DSSRecorder.VB_VarHelpID = -1

Private FormFullHeight As Integer
Private FormLowHeihgt As Integer

Private mDict As clsDict
Private mNewDict As Boolean
Private mSoundReadOnly As Boolean
Private mTextReadOnly As Boolean
Private mFloating As Boolean
Private mAutoRewind As Integer
Private mCloseText(2) As String
Private mCloseTip(2) As String
Private mIsInChangeMode As Boolean
Public Sub ForceUnload()

   mForceUnload = True
   Unload Me
End Sub
Private Sub chkChangeDict_Click()

   mIsInChangeMode = chkChangeDict.Value = Checked
   mTextReadOnly = Not mIsInChangeMode
   SetEnabled
   SetChoiseFromChangeMode
End Sub

Private Sub SetChoiseFromChangeMode()

   If mIsInChangeMode Then
      ucCloseChoice.ChoiceText(0) = Client.Texts.Txt(1030118, "St�ng utan att spara")
      ucCloseChoice.ChoiceTip(0) = Client.Texts.ToolTip(1030118, "L�mna diktatet utan �ndring")
      ucCloseChoice.ChoiceText(1) = Client.Texts.Txt(1030119, "Spara f�r utskrift")
      ucCloseChoice.ChoiceTip(1) = Client.Texts.ToolTip(1030119, "Status inspelat")
      ucCloseChoice.ChoiceText(2) = Client.Texts.Txt(1030120, "Spara som utskrivet")
      ucCloseChoice.ChoiceTip(2) = Client.Texts.ToolTip(1030120, "Status utskrivet")
   Else
      ucCloseChoice.ChoiceText(0) = mCloseText(0)
      ucCloseChoice.ChoiceTip(0) = mCloseTip(0)
      ucCloseChoice.ChoiceText(1) = mCloseText(1)
      ucCloseChoice.ChoiceTip(1) = mCloseTip(1)
      ucCloseChoice.ChoiceText(2) = mCloseText(2)
      ucCloseChoice.ChoiceTip(2) = mCloseTip(2)
   End If
End Sub
Private Sub DSSRecorder_GruEvent(EventType As CareTalkDSSRec3.Gru_Event, Data As Long)

   Dim I As Integer
   
   If EventType = GRU_BUTTONPRESS Then
      If Data = GRU_BUT_INDEX Then
         If Client.SysSettings.PlayerIndexButtonAsCloseDict Then
            If CheckMandatoryData() Then
               For I = 2 To 0 Step -1
                  If Len(ucCloseChoice.ChoiceText(I)) > 0 Then
                     ucCloseChoice.ChoiceValue = I
                     RaiseEvent CloseChoiceSelected(I)
                     Unload Me
                     Exit Sub
                  End If
               Next I
            End If
         End If
      End If
   End If
End Sub

Private Sub Form_Activate()

   'Has to be here...
   ShowWindow Me.hwnd, SW_Hide
   Me.Caption = Me.Caption
   ShowWindow Me.hwnd, SW_ShowNormal
      
   CenterAndTranslateForm Me, frmMain
   
   SetWindowTopMostAndForeground Me
   If mFloating Then
      WindowFloating Me
   End If
End Sub

Public Property Let CloseText(Index As Integer, Text As String)

   mCloseText(Index) = Text
   ucCloseChoice.ChoiceText(Index) = Text
End Property
Public Property Let CloseTip(Index As Integer, Text As String)

   mCloseTip(Index) = Text
   ucCloseChoice.ChoiceTip(Index) = Text
End Property
Public Property Get CloseChoice() As Integer

   Dim Value As Integer

   Value = ucCloseChoice.ChoiceValue

   If mIsInChangeMode Then
      CloseChoice = Value + 10
   Else
      CloseChoice = Value
   End If
End Property
Public Property Let CloseChoice(Index As Integer)

   ucCloseChoice.ChoiceValue = Index
End Property

Private Sub cboDictType_Click()

   If Screen.ActiveControl Is cboDictType Then
      mDict.DictTypeId = cboDictType.ItemData(cboDictType.ListIndex)
      mDict.DictTypeText = cboDictType.List(cboDictType.ListIndex)
      mDict.InfoDirty = True
   End If
End Sub

Private Sub cboPriority_Click()

   If Screen.ActiveControl Is cboPriority Then
      mDict.PriorityId = Client.PriorityMgr.IdFromIndex(cboPriority.ListIndex)
      mDict.PriorityText = Client.PriorityMgr.TextFromIndex(cboPriority.ListIndex)
      mDict.ExpiryDate = DateAdd("d", Client.PriorityMgr.DaysFromIndex(cboPriority.ListIndex), mDict.Created)
      lblExpiryDate.Caption = Format$(mDict.ExpiryDate, "ddddd")
      mDict.InfoDirty = True
   End If
End Sub

Private Sub chkNoPatient_Click()

   If Screen.ActiveControl Is chkNoPatient Then
      mDict.NoPatient = (chkNoPatient.Value = vbChecked)
      mDict.InfoDirty = True
   End If
   SetEnabled
End Sub

Private Sub Form_Initialize()

   Trc "frmDict Initialize", ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   Dim K As Integer
   Dim Sh As Integer
   Dim B As PlayerButEnum
   Dim VolChange As Integer
   Dim SpeedChange As Integer
   Dim CloseIndex As Integer
   Dim CloseX As Boolean
   
   B = butNone
   CloseIndex = -1
   CloseX = False
   Sh = Shift And 7
   K = Sh * 256 + (KeyCode And 255)
   Debug.Print KeyCode, Shift, Sh, K
   Select Case K
      Case Client.SysSettings.PlayerKeyPlay
         B = butPlay
      Case Client.SysSettings.PlayerKeyPause
         B = butPause
      Case Client.SysSettings.PlayerKeyStop
         B = butStop
      Case Client.SysSettings.PlayerKeyStart
         B = butStart
      Case Client.SysSettings.PlayerKeyRewind
         B = butRewind
      Case Client.SysSettings.PlayerKeyForward
         B = butForward
      Case Client.SysSettings.PlayerKeyEnd
         B = butEnd
      Case Client.SysSettings.PlayerKeyRec
         B = butRec
      Case Client.SysSettings.PlayerKeyVolumeUp
         VolChange = 1
      Case Client.SysSettings.PlayerKeyVolumeDown
         VolChange = 2
      Case Client.SysSettings.PlayerKeySpeedUp
         SpeedChange = 1
      Case Client.SysSettings.PlayerKeySpeedDown
         SpeedChange = 2
      Case Client.SysSettings.PlayerKeyClose1, Client.SysSettings.PlayerKeyClose1Alt, _
           Client.SysSettings.PlayerKeyEscape, Client.SysSettings.PlayerKeyEscapeAlt
         CloseIndex = 0
      Case Client.SysSettings.PlayerKeyClose2, Client.SysSettings.PlayerKeyClose2Alt
         CloseIndex = 1
      Case Client.SysSettings.PlayerKeyClose3, Client.SysSettings.PlayerKeyClose3Alt
         CloseIndex = 2
      Case Client.SysSettings.PlayerKeyCloseX, Client.SysSettings.PlayerKeyCloseXAlt
         CloseX = True
   End Select
   If B <> butNone Then
      KeyCode = 0
      ucDSSRecGUI.ExternalButton B
   End If
   If VolChange <> 0 Then
      KeyCode = 0
      ucDSSRecGUI.ExternalVolumeChange VolChange = 1
   End If
   If SpeedChange <> 0 Then
      KeyCode = 0
      ucDSSRecGUI.ExternalSpeedChange SpeedChange = 1
   End If
   If CloseIndex >= 0 Then
      KeyCode = 0
      If Len(ucCloseChoice.ChoiceText(CloseIndex)) > 0 Then
         ucCloseChoice.ChoiceValue = CloseIndex
         If CheckMandatoryData() Or CloseIndex = 0 Then
            RaiseEvent CloseChoiceSelected(CloseIndex)
            Unload Me
         End If
      End If
   End If
   If CloseX Then
      KeyCode = 0
      If ucCloseChoice.ChoiceValue > -1 Then
         If CheckMandatoryData() Or CloseIndex = 0 Then
            RaiseEvent CloseChoiceSelected(ucCloseChoice.ChoiceValue)
            Unload Me
         End If
      End If
   End If
End Sub

Private Sub Form_Load()

   Dim Org As New clsOrg
   Dim I As Integer
   
   Trc "frmDict load", ""
   
   mForceUnload = False
   
   imgPin.Visible = Client.SysSettings.PlayerShowOnTop
   imgPinin.Visible = Client.SysSettings.PlayerShowOnTop
   imgLess.Visible = Client.SysSettings.PlayerShowSmallerWindow
   imgMore.Visible = False
   lblNoPatientTitle.Visible = Client.SysSettings.DictInfoUseNoPat
   chkNoPatient.Visible = Client.SysSettings.DictInfoUseNoPat
   If Client.SysSettings.DictInfoUseKeyWords Then
      lblTxtTitle.Visible = True
      txtTxt.Visible = True
      FormFullHeight = 5055
   Else
      lblTxtTitle.Visible = False
      txtTxt.Visible = False
      FormFullHeight = 4515
   End If
   FormLowHeihgt = 1100
   Me.Height = FormFullHeight
   ucOrgTree.Height = FormFullHeight - 1200
   
   Set DSSRecorder = Client.DSSRec
   Set ucDSSRecGUI.DSSRecorder = Client.DSSRec

   
   Client.OrgMgr.Init False
   
   For I = 0 To Client.OrgMgr.Count - 1
      Client.OrgMgr.GetSortedOrg Org, I
      If Org.ShowInTree Then
         If (Org.Roles.Author Or Org.Roles.TextEditor) And Org.DictContainer Then
            If Org.OrgId = Client.User.HomeOrgId Then
               ucOrgTree.AddNode Org.OrgId, Org.ShowParent, Org.OrgText, 7, True
            Else
               ucOrgTree.AddNode Org.OrgId, Org.ShowParent, Org.OrgText, 1, True
            End If
         Else
            ucOrgTree.AddNode Org.OrgId, Org.ShowParent, Org.OrgText, 5, False
         End If
      End If
   Next I
   
End Sub
Public Sub RestoreSettings(Settings As clsStringStore)

   If Client.SysSettings.PlayerShowOnTop Then
      SetFoatingWindows Settings.GetBool("Window", "Floating", False)
   End If
   mAutoRewind = Settings.GetLong("Player", "AutoRewind", 1500)
End Sub
Public Sub SaveSettings(Settings As clsStringStore)

   Settings.AddBool "Window", "Floating", mFloating
   Settings.AddLong "Player", "AutoRewind", CLng(mAutoRewind)
End Sub
Public Sub EditDictation(ByRef Dictation As clsDict, ByVal NewDict As Boolean)

   Set mDict = Dictation
   mNewDict = NewDict
   If mNewDict Then
      mTextReadOnly = False
      chkChangeDict.Visible = False
      lblChangeDict.Visible = False
   Else
      mTextReadOnly = Not Client.User.UserId = Dictation.AuthorId
      chkChangeDict.Visible = Not Dictation.TextReadOnly And Dictation.SoundReadOnly
      chkChangeDict.Value = Unchecked
      lblChangeDict.Visible = Not Dictation.TextReadOnly And Dictation.SoundReadOnly
   End If
   mSoundReadOnly = Dictation.SoundReadOnly
   
   SetEnabled
   
   Me.Caption = mDict.DictTypeText & ": " & mDict.Pat.PatIdFormatted
   If Len(mDict.LocalFilename) > 0 Then
      ucDSSRecGUI.Visible = True
      ucDSSRecGUI.AutoRewind = mAutoRewind
      If mNewDict Then
         ucDSSRecGUI.ReadOnly = mSoundReadOnly
         ucDSSRecGUI.CreateNewFile mDict.LocalFilename
      Else
         ucDSSRecGUI.ReadOnly = mSoundReadOnly
         ucDSSRecGUI.OpenAndPlay mDict.LocalFilename
      End If
   Else
      ucDSSRecGUI.Visible = False
   End If
   ShowDictation
   mDict.InfoDirty = False
End Sub

Private Sub ShowDictation()

   If mDict.OrgId > 0 Then
      ucOrgTree.PickOrgId mDict.OrgId
   End If
   If mDict.NoPatient Then
      chkNoPatient.Value = vbChecked
   Else
      chkNoPatient.Value = vbUnchecked
   End If
   txtPatId.Text = mDict.Pat.PatIdFormatted
   txtPatName.Text = mDict.Pat.PatName
   
   Client.DictTypeMgr.FillCombo cboDictType, mDict.OrgId, mDict.DictTypeId, True
   mDict.DictTypeId = cboDictType.ItemData(cboDictType.ListIndex)
   mDict.DictTypeText = cboDictType.List(cboDictType.ListIndex)
   'cboDictType.ListIndex = Client.DictTypeMgr.IndexFromId(mDict.DictTypeId)  '!!!
   
   Client.PriorityMgr.FillCombo cboPriority
   cboPriority.ListIndex = Client.PriorityMgr.IndexFromId(mDict.PriorityId)
   
   lblStatus.Caption = mDict.StatusText
   lblExpiryDate.Caption = Format$(mDict.ExpiryDate, "ddddd")
   lblAuthor.Caption = mDict.AuthorLongName
   lblTranscriber.Caption = mDict.TranscriberLongName
   lblCreated.Caption = Format$(mDict.Created, "ddddd ttttt")
   If mDict.Changed <> 0 Then
      lblChanged.Caption = Format$(mDict.Changed, "ddddd ttttt")
   Else
      lblChanged.Caption = ""
   End If
   If mDict.TranscribedDate <> 0 Then
      lblTranscribedDate.Caption = Format$(mDict.TranscribedDate, "ddddd ttttt")
   Else
      lblTranscribedDate.Caption = ""
   End If
   
   txtTxt.Text = mDict.Txt
End Sub

Private Sub Form_Terminate()

   Trc "frmDict Terminate", ""
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Dim Ok As Boolean
   Dim SoundL As Long

   Ok = True
   If Not mForceUnload Then
      If ucCloseChoice.ChoiceValue < 0 Then
         Ok = False
      ElseIf ucCloseChoice.ChoiceValue > 0 Then
         If Not CheckMandatoryData() Then
            Ok = False
         End If
      End If
   Else
      If ucCloseChoice.ChoiceValue < 0 Then
         ucCloseChoice.ChoiceValue = 0
      End If
   End If
   
   If Ok Then
      mDict.SoundDirty = mDict.SoundDirty Or ucDSSRecGUI.Dirty
      SoundL = ucDSSRecGUI.SoundLengthInSec
      If SoundL > 0 Then   'If we lost device, length can wrong 0. Don't save it
         mDict.SoundLength = ucDSSRecGUI.SoundLengthInSec
      End If
      mAutoRewind = ucDSSRecGUI.AutoRewind
      ucDSSRecGUI.StopAndClose
      Set ucDSSRecGUI.DSSRecorder = Nothing
      Set mDict = Nothing
   Else
      MsgBox Client.Texts.Txt(1030101, "Uppgifterna �r inte kompletta!"), vbCritical
      Cancel = True
   End If
End Sub

Private Sub imgLess_Click()

   imgLess.Visible = False
   imgMore.Visible = True
   Me.Height = FormLowHeihgt
End Sub

Private Sub imgMore_Click()

   imgLess.Visible = True
   imgMore.Visible = False
   Me.Height = FormFullHeight
End Sub

Private Sub imgPin_Click()

   SetFoatingWindows True
End Sub
Private Sub SetFoatingWindows(Value As Boolean)

   mFloating = Value
   If mFloating Then
      imgPinin.Visible = True
      imgPin.Visible = False
      WindowFloating Me
   Else
      imgPinin.Visible = False
      imgPin.Visible = True
      WindowNotFloating Me
   End If
End Sub

Private Sub imgPinin_Click()

   SetFoatingWindows False
End Sub

Private Sub lblPatIdTitle_Click()

   Clipboard.Clear
   Clipboard.SetText txtPatId.Text
End Sub

Private Sub lblPatnameTitle_Click()

   Clipboard.Clear
   Clipboard.SetText txtPatName.Text
End Sub

Private Sub lblTxtTitle_Click()

   Clipboard.Clear
   Clipboard.SetText txtTxt.Text
End Sub

Private Sub txtPatId_Change()

   If Screen.ActiveControl Is txtPatId Then
      mDict.Pat.PatId = txtPatId.Text
      mDict.InfoDirty = True
   End If
   SetEnabled
End Sub

Private Sub txtPatId_KeyPress(KeyAscii As Integer)

   If Not Client.SysSettings.DictInfoAlfaInPatid Then
      If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii < 32 Or KeyAscii = 45) Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtPatName_Change()

   If Screen.ActiveControl Is txtPatName Then
      mDict.Pat.PatName = txtPatName.Text
      mDict.InfoDirty = True
   End If
   SetEnabled
End Sub

Private Sub txtTxt_Change()

   If Screen.ActiveControl Is txtTxt Then
      mDict.Txt = txtTxt.Text
      mDict.InfoDirty = True
   End If
   SetEnabled
End Sub

Private Sub ucCloseChoice_CloseClicked()

   Unload Me
End Sub

Private Sub ucCloseChoice_NewSelect(Index As Integer)

   Dim Res As Integer
   
   If mIsInChangeMode Then
      Res = Index + 10
   Else
      Res = Index
   End If
   Trc "frmDict Event CloseChoiceSelected", Format$(Res)
   RaiseEvent CloseChoiceSelected(Res)
End Sub

Private Sub ucDSSRecGUI_ChangeIcon(NewIcon As Image)

   Me.Icon = NewIcon.Picture
End Sub

Private Sub ucDSSRecGUI_PosChange(PosInMilliSec As Long, LengthInMilliSec As Long, Formated As String)

   Me.Caption = Formated
End Sub

Private Sub ucOrgTree_NewSelect(OrgId As Long, Txt As String)

   If Screen.ActiveControl Is ucOrgTree Then
      mDict.OrgId = OrgId
      mDict.OrgText = Txt
      mDict.InfoDirty = True
      Client.DictTypeMgr.FillCombo cboDictType, mDict.OrgId, mDict.DictTypeId, True
      mDict.DictTypeId = cboDictType.ItemData(cboDictType.ListIndex)
      mDict.DictTypeText = cboDictType.List(cboDictType.ListIndex)
   End If
   SetEnabled
End Sub
Private Sub SetEnabled()

   Dim Enbld As Boolean
   
   Enbld = Not mTextReadOnly
   ucOrgTree.Enabled = Enbld
   txtPatId.Enabled = mDict.AuthorId = Client.User.UserId Or mDict.AuthorId = 0
   txtPatName.Enabled = mDict.AuthorId = Client.User.UserId Or mDict.AuthorId = 0
   chkNoPatient.Enabled = Enbld
   cboDictType.Enabled = Enbld
   cboPriority.Enabled = Enbld
   txtTxt.Enabled = Enbld
   
   CheckMandatoryData
   
   Client.DictMgr.SaveTempDictationInfo mDict, tdiUpdateInfo
End Sub
Private Function CheckMandatoryData() As Boolean

   Dim Ok As Boolean

   Ok = True
   If chkNoPatient.Value <> vbChecked Then
      If Not CheckPatId(txtPatId) Then
         lblPatIdMissing.Visible = True
         Ok = False
      Else
         lblPatIdMissing.Visible = False
      End If
      If Not CheckPatname(txtPatName) Then
         lblPatNameMissing.Visible = True
         Ok = False
      Else
         lblPatNameMissing.Visible = False
      End If
   Else
      lblPatIdMissing.Visible = False
      lblPatNameMissing.Visible = False
   End If
   If mDict.OrgId = 0 Then
      lblOrgMissing.Visible = True
      Ok = False
   Else
      lblOrgMissing.Visible = False
   End If
   CheckMandatoryData = Ok
End Function