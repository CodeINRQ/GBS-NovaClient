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
   Begin VB.TextBox txtTxt 
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   30
      Top             =   4200
      Width           =   6255
   End
   Begin VB.ComboBox cboPriority 
      Height          =   315
      Left            =   2400
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.ComboBox cboDictType 
      Height          =   315
      Left            =   2400
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CheckBox chkNoPatient 
      Height          =   255
      Left            =   4320
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
      MaxLength       =   13
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin CareTalk.ucCloseChoice ucCloseChoice 
      Height          =   1695
      HelpContextID   =   1030000
      Left            =   6120
      TabIndex        =   6
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
   End
   Begin CareTalk.ucDSSRecGUI ucDSSRecGUI 
      Height          =   495
      HelpContextID   =   1030000
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   873
   End
   Begin CareTalk.ucOrgTree ucOrgTree 
      Height          =   3855
      HelpContextID   =   1030000
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5953
   End
   Begin VB.Label lblTxtTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Nyckelord:"
      Height          =   255
      Left            =   2400
      TabIndex        =   31
      Tag             =   "1030116"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblOrgMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2280
      TabIndex        =   29
      Top             =   600
      Width           =   135
   End
   Begin VB.Label lblPatNameMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   28
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lblPatIdMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   27
      Top             =   840
      Width           =   135
   End
   Begin VB.Image imgLess 
      Height          =   480
      Left            =   8565
      Picture         =   "Dict.frx":058A
      Top             =   405
      Width           =   480
   End
   Begin VB.Image imgMore 
      Height          =   480
      Left            =   8565
      Picture         =   "Dict.frx":0E54
      Top             =   405
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgPinin 
      Height          =   210
      Left            =   8520
      Picture         =   "Dict.frx":171E
      Tag             =   "1030115"
      ToolTipText     =   "Alltid överst"
      Top             =   120
      Width           =   210
      Visible         =   0   'False
   End
   Begin VB.Image imgPin 
      Height          =   210
      Left            =   8520
      Picture         =   "Dict.frx":17F9
      Tag             =   "1030114"
      ToolTipText     =   "Normalt fönster"
      Top             =   120
      Width           =   210
   End
   Begin VB.Label lblTranscribedDateTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Utskrivet:"
      Height          =   255
      Left            =   6120
      TabIndex        =   26
      Tag             =   "1030113"
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblTranscribedDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   25
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   24
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblStatusTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Tag             =   "1030106"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblExpiryDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblExpiryDateTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Utskrift senast:"
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Tag             =   "1030108"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblTranscriber 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblTranscriberTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Utskrivare:"
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Tag             =   "1030112"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblAuthorTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Intalare:"
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Tag             =   "1030110"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblPriorityTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Prioritet:"
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Tag             =   "1030107"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Diktattyp:"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Tag             =   "1030105"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblChanged 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblChangedTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Ändrat:"
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Tag             =   "1030111"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblCreated 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblCreatedTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Inläst:"
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Tag             =   "1030109"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblNoPatientTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingen patient:"
      Height          =   255
      Left            =   4320
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

Private WithEvents DSSRecorder As CareTalkDSSRec3.DSSRecorder
Attribute DSSRecorder.VB_VarHelpID = -1

Private FormFullHeight As Integer
Private FormLowHeihgt As Integer

Private mDict As clsDict
Private mNewDict As Boolean
Private mReadOnly As Boolean
Private mFloating As Boolean

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
End Sub

Public Property Let CloseText(Index As Integer, Text As String)

   ucCloseChoice.ChoiceText(Index) = Text
End Property
Public Property Let CloseTip(Index As Integer, Text As String)

   ucCloseChoice.ChoiceTip(Index) = Text
End Property
Public Property Get CloseChoice() As Integer

   CloseChoice = ucCloseChoice.ChoiceValue
End Property
Public Property Let CloseChoice(Index As Integer)

   ucCloseChoice.ChoiceValue = Index
End Property

Private Sub cboDictType_Click()

   If Screen.ActiveControl Is cboDictType Then
      mDict.DictTypeId = Client.DictTypeMgr.IdFromIndex(cboDictType.ListIndex)
      mDict.DictTypeText = Client.DictTypeMgr.TextFromIndex(cboDictType.ListIndex)
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
   
   B = butNone
   CloseIndex = -1
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
      Case Client.SysSettings.PlayerKeyClose1
         CloseIndex = 0
      Case Client.SysSettings.PlayerKeyClose2
         CloseIndex = 1
      Case Client.SysSettings.PlayerKeyClose3
         CloseIndex = 2
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
End Sub

Private Sub Form_Load()

   Dim Org As New clsOrg
   Dim I As Integer
   
   Trc "frmDict load", ""
   
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
         If Org.Roles.Author And Org.DictContainer Then
            ucOrgTree.AddNode Org.OrgId, Org.ShowParent, Org.OrgText, 1, True
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
End Sub
Public Sub SaveSettings(Settings As clsStringStore)

   Settings.AddBool "Window", "Floating", mFloating
End Sub
Public Sub EditDictation(ByRef Dictation As clsDict, ByVal NewDict As Boolean)

   Set mDict = Dictation
   mNewDict = NewDict
   mReadOnly = Dictation.ReadOnly
   
   SetEnabled
   
   Me.Caption = mDict.DictTypeText & ": " & mDict.Pat.PatIdFormatted
   If Len(mDict.LocalFilename) > 0 Then
      ucDSSRecGUI.Visible = True
      If mNewDict Then
         ucDSSRecGUI.ReadOnly = mReadOnly
         ucDSSRecGUI.CreateNewFile mDict.LocalFilename
      Else
         ucDSSRecGUI.ReadOnly = mReadOnly
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
   
   Client.DictTypeMgr.FillCombo cboDictType
   cboDictType.ListIndex = Client.DictTypeMgr.IndexFromId(mDict.DictTypeId)
   
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

   Ok = True
   If ucCloseChoice.ChoiceValue < 0 Then
      Ok = False
   ElseIf ucCloseChoice.ChoiceValue > 0 Then
      If Not CheckMandatoryData() Then
         Ok = False
      End If
   End If
   
   If Ok Then
      mDict.SoundDirty = mDict.SoundDirty Or ucDSSRecGUI.Dirty
      mDict.SoundLength = ucDSSRecGUI.SoundLengthInSec
      ucDSSRecGUI.StopAndClose
      Set ucDSSRecGUI.DSSRecorder = Nothing
      Set mDict = Nothing
   Else
      MsgBox Client.Texts.Txt(1030101, "Uppgifterna är inte kompletta!"), vbCritical
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

   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii < 32 Or KeyAscii = 45) Then
      KeyAscii = 0
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

   Trc "frmDict Event CloseChoiceSelected", Format$(Index)
   RaiseEvent CloseChoiceSelected(Index)
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
   End If
   SetEnabled
End Sub
Private Sub SetEnabled()

   ucOrgTree.Enabled = Not mReadOnly
   txtPatId.Enabled = Not mReadOnly
   txtPatName.Enabled = Not mReadOnly
   chkNoPatient.Enabled = Not mReadOnly
   cboDictType.Enabled = Not mReadOnly
   cboPriority.Enabled = Not mReadOnly
   txtTxt.Enabled = Not mReadOnly
   
   CheckMandatoryData
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
