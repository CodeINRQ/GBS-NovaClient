VERSION 5.00
Begin VB.Form frmDict 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diktat"
   ClientHeight    =   5160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8760
   HelpContextID   =   1030000
   Icon            =   "Dict.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8760
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1030100"
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picWarning 
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   8115
      TabIndex        =   38
      Top             =   2400
      Width           =   8175
      Visible         =   0   'False
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Låg insignal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   0
         TabIndex        =   39
         Tag             =   "1030124"
         Top             =   0
         Width           =   8175
      End
   End
   Begin VB.TextBox txtNote 
      Height          =   285
      Left            =   2400
      MaxLength       =   100
      TabIndex        =   11
      Top             =   4200
      Width           =   6255
   End
   Begin VB.CheckBox chkChangeDict 
      Height          =   270
      Left            =   5160
      Picture         =   "Dict.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   840
      Width           =   270
   End
   Begin VB.TextBox txtTxt 
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   13
      Top             =   4800
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
      Left            =   6120
      TabIndex        =   14
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
   End
   Begin CareTalk.ucDSSRecGUI ucDSSRecGUI 
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   60
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   873
   End
   Begin CareTalk.ucOrgTree ucOrgTree 
      Height          =   4455
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5953
   End
   Begin VB.Label lblNoteTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "&Notering:"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Tag             =   "1030121"
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lblPriorityMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   37
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label lblDictTypeMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   36
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label lblChangeDict 
      BackStyle       =   0  'Transparent
      Caption         =   "Ändra:"
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Tag             =   "1030117"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblTxtTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Nyckelord:"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Tag             =   "1030116"
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label lblOrgMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2280
      TabIndex        =   33
      Top             =   600
      Width           =   135
   End
   Begin VB.Label lblPatNameMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   32
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lblPatIdMissing 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   31
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
      ToolTipText     =   "Alltid överst"
      Top             =   120
      Width           =   210
      Visible         =   0   'False
   End
   Begin VB.Image imgPin 
      Height          =   210
      Left            =   8520
      Picture         =   "Dict.frx":192F
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
      TabIndex        =   30
      Tag             =   "1030113"
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblTranscribedDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   29
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   28
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblStatusTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Tag             =   "1030106"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblExpiryDate 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblExpiryDateTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Utskrift senast:"
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Tag             =   "1030108"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblTranscriber 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblTranscriberTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Utskrivare:"
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Tag             =   "1030112"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblAuthorTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Intalare:"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
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
   Begin VB.Label lblDictTypeTitle 
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
      TabIndex        =   20
      Top             =   3240
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblChangedTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Ändrat:"
      Height          =   255
      Left            =   6120
      TabIndex        =   19
      Tag             =   "1030111"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblCreated 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblCreatedTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Inläst:"
      Height          =   255
      Left            =   6120
      TabIndex        =   16
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
    (ByVal hWnd As Long, _
     ByVal nCmdShow As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, _
   ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, _
   ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Const SC_CLOSE = &HF060
Private Const MF_BYCOMMAND = &H0&
Private Const WM_NCACTIVATE = &H86
               
Public Event CloseChoiceSelected(Index As Integer)

Public mForceUnload As Boolean

Private WithEvents DSSRecorder As clsDSSRecorder
Attribute DSSRecorder.VB_VarHelpID = -1
Private FormFullHeight As Integer
Private FormLowHeight As Integer

Private mDict As clsDict
Private mNewDict As Boolean
Private mSoundReadOnly As Boolean
Private mTextReadOnly As Boolean
Private mFloating As Boolean
Private mAutoRewind As Integer
Private mCloseText(2) As String
Private mCloseTip(2) As String
Private mIsInChangeMode As Boolean
Private mPos As Long
Private mUserIsSysAdmin As Boolean

Private UseAutomaticTranscribersStatusChange As Boolean
Private WeHaveBeenInMidlePartOfDictation As Boolean
Private WeHaveBeenInLastPartOfDictation As Boolean
Private InitiallyLengthMilliSec As Long

Public Property Let AutomaticTranscribersStatusChange(Activate As Boolean)

   UseAutomaticTranscribersStatusChange = Activate
End Property
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

   Static ChoiceValueBeforeCahngeMode As Integer

   If mIsInChangeMode Then
      ChoiceValueBeforeCahngeMode = ucCloseChoice.ChoiceValue
      ucCloseChoice.ChoiceText(0) = Client.Texts.Txt(1030118, "Stäng utan att spara")
      ucCloseChoice.ChoiceTip(0) = Client.Texts.ToolTip(1030118, "Lämna diktatet utan ändring")
      ucCloseChoice.ChoiceText(1) = Client.Texts.Txt(1030119, "Spara för utskrift")
      ucCloseChoice.ChoiceTip(1) = Client.Texts.ToolTip(1030119, "Status inspelat")
      ucCloseChoice.ChoiceText(2) = Client.Texts.Txt(1030120, "Spara som utskrivet")
      ucCloseChoice.ChoiceTip(2) = Client.Texts.ToolTip(1030120, "Status utskrivet")
      If mDict.StatusId < Transcribed Then
         ucCloseChoice.ChoiceValue = 1
      Else
         ucCloseChoice.ChoiceValue = 2
      End If
   Else
      ucCloseChoice.ChoiceText(0) = mCloseText(0)
      ucCloseChoice.ChoiceTip(0) = mCloseTip(0)
      ucCloseChoice.ChoiceText(1) = mCloseText(1)
      ucCloseChoice.ChoiceTip(1) = mCloseTip(1)
      ucCloseChoice.ChoiceText(2) = mCloseText(2)
      ucCloseChoice.ChoiceTip(2) = mCloseTip(2)
      ucCloseChoice.ChoiceValue = ChoiceValueBeforeCahngeMode
   End If
End Sub
Private Sub DSSRecorder_GruEvent(EventType As Gru_Event, Data As Long)

   Dim I As Integer
   Dim Pos As Long
   
   If EventType = GRU_BUTTONPRESS Then
      If Data = GRU_BUT_INDEX Then
         'Debug.Print "GruEvent But_Index+"
         If Client.SysSettings.PlayerIndexButtonAsCloseDict Then
            'If CheckMandatoryData() Then
               For I = 2 To 0 Step -1
                  If Len(ucCloseChoice.ChoiceText(I)) > 0 Then
                     ucCloseChoice.ChoiceValue = I
                     RaiseEvent CloseChoiceSelected(I)
                     Debug.Print "GruEvent Unload Me"
                     Unload Me
                     Exit Sub
                  End If
               Next I
            'End If
         End If
      End If
   Else
      If EventType = GRU_POSCHANGE Then
         If UseAutomaticTranscribersStatusChange Then
            If Data > 3000 Then
               If Not WeHaveBeenInMidlePartOfDictation Then
                  WeHaveBeenInMidlePartOfDictation = True
                  ucCloseChoice.ChoiceValue = 1
               End If
            End If
            If Data > InitiallyLengthMilliSec - 5000 Then
               If Not WeHaveBeenInLastPartOfDictation Then
                  WeHaveBeenInMidlePartOfDictation = True
                  WeHaveBeenInLastPartOfDictation = True
                  ucCloseChoice.ChoiceValue = 2
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub Form_Activate()

   If LastfrmDictLeft <> 0 Or LastfrmDictTop <> 0 Then
      Me.Move LastfrmDictLeft, LastfrmDictTop
      TranslateForm Me
   Else
      CenterAndTranslateForm Me, frmMain
   End If
   
   'Has to be here...
   ShowWindow Me.hWnd, SW_Hide
   Me.Caption = Me.Caption
   ShowWindow Me.hWnd, SW_ShowNormal
         
   SetWindowTopMostAndForeground Me
   If mFloating Then
      WindowFloating Me
   End If
   ShowFormCaption
        
End Sub

Public Property Let CloseText(Index As Integer, Text As String)

   mCloseText(Index) = Text
   ucCloseChoice.ChoiceText(Index) = Text
   If UseAutomaticTranscribersStatusChange Then
      ucCloseChoice.ChoiceValue = 0
   End If
End Property
Public Property Let CloseTip(Index As Integer, Text As String)

   mCloseTip(Index) = Text
   ucCloseChoice.ChoiceTip(Index) = Text
   If UseAutomaticTranscribersStatusChange Then
      ucCloseChoice.ChoiceValue = 0
   End If
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

   Dim DT As clsDictType

   If Screen.ActiveControl Is cboDictType Then
      Client.DictTypeMgr.GetFromId DT, cboDictType.ItemData(cboDictType.ListIndex)
     
      mDict.DictTypeId = DT.DictTypeId
      mDict.DictTypeText = DT.DictTypeText
      mDict.InfoDirty = True
   End If
   SetEnabled
End Sub

Private Sub cboPriority_Click()

   Dim Prio As clsPriority

   If Screen.ActiveControl Is cboPriority Then
      Client.PriorityMgr.GetFromId Prio, cboPriority.ItemData(cboPriority.ListIndex)
      
      mDict.PriorityId = Prio.PriorityId
      mDict.PriorityText = Prio.PriortyText
      mDict.ExpiryDate = DateAdd("d", Prio.Days, mDict.Created)
      lblExpiryDate.Caption = Format$(mDict.ExpiryDate, "ddddd")
      mDict.InfoDirty = True
   End If
   SetEnabled
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
   Dim WindowSize As Boolean
   
   
   B = butNone
   CloseIndex = -1
   CloseX = False
   Sh = Shift And 7
   K = Sh * 256 + (KeyCode And 255)
   'Debug.Print KeyCode, Shift, Sh, K,
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
      Case Client.SysSettings.PlayerKeyWindowSize
         WindowSize = True
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
            Unload Me
         End If
      End If
   End If
   If CloseX Then
      KeyCode = 0
      If ucCloseChoice.ChoiceValue > -1 Then
         If CheckMandatoryData() Or CloseIndex = 0 Then
            Unload Me
         End If
      End If
   End If
   If WindowSize Then
      If imgLess.Visible Then
         imgLess_Click
      ElseIf imgMore.Visible Then
         imgMore_Click
      End If
   End If
End Sub

Private Sub Form_Load()

   Dim Org As New clsOrg
   Dim I As Integer
   Dim hMenu As Long, Success As Long

   Trc "frmDict load", ""

   'Disable Close button (X)
   hMenu = GetSystemMenu(Me.hWnd, 0)
   Success = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
   SendMessage Me.hWnd, WM_NCACTIVATE, 0&, 0&
   SendMessage Me.hWnd, WM_NCACTIVATE, 1&, 0&
      
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
      FormFullHeight = 5655
   Else
      lblTxtTitle.Visible = False
      txtTxt.Visible = False
      FormFullHeight = 5115
   End If
   FormLowHeight = 1100
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
Public Sub EditDictation(ByRef Dictation As clsDict, ByVal NewDict As Boolean, DictButton As Long)

   mUserIsSysAdmin = Client.OrgMgr.CheckUserRole(Dictation.OrgId, RTSysAdmin)
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
   
   If mDict.LocalDictFile.IsSoundToPlay Then
      ucDSSRecGUI.Visible = True
      ucDSSRecGUI.AutoRewind = mAutoRewind
      If mNewDict Then
         ucDSSRecGUI.ReadOnly = mSoundReadOnly
         ucDSSRecGUI.CreateNewFile mDict.LocalDictFile.LocalFilenamePlay, DictButton = GRU_BUT_BUTREC
      Else
         ucDSSRecGUI.ReadOnly = mSoundReadOnly
         ucDSSRecGUI.OpenAndPlay mDict.LocalDictFile.LocalFilenamePlay
      End If
   Else
      ucDSSRecGUI.Visible = False
   End If
   ucDSSRecGUI.Position = mDict.CurrentPos
   InitiallyLengthMilliSec = ucDSSRecGUI.SoundLengthInSec * 1000
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
   
   Client.DictTypeMgr.FillCombo cboDictType, mDict.OrgId, mDict.DictTypeId, mDict.DictTypeIdNoDefault, True
   If cboDictType.ListIndex >= 0 Then
      Dim DictType As clsDictType
      Client.DictTypeMgr.GetFromId DictType, cboDictType.ItemData(cboDictType.ListIndex)
      mDict.DictTypeId = DictType.DictTypeId
      mDict.DictTypeText = DictType.DictTypeText
      Set DictType = Nothing
   End If
   
   Client.PriorityMgr.FillCombo cboPriority, mDict.OrgId, mDict.PriorityId, True
   If cboPriority.ListIndex >= 0 Then
      Dim Priority As clsPriority
      Client.PriorityMgr.GetFromId Priority, cboPriority.ItemData(cboPriority.ListIndex)
      mDict.PriorityId = Priority.PriorityId
      mDict.PriorityText = Priority.PriortyText
      mDict.ExpiryDate = DateAdd("d", Priority.Days, mDict.Created)
      Set Priority = Nothing
   Else
      mDict.ExpiryDate = mDict.Created
   End If
   
   lblStatus.Caption = mDict.StatusText
   lblExpiryDate.Caption = Format$(mDict.ExpiryDate, "ddddd")
   lblAuthor.Caption = mDict.AuthorLongName
   lblTranscriber.Caption = mDict.TranscriberLongName
   lblCreated.Caption = Format$(mDict.Created, "ddddd ttttt")
   If mDict.Changed <> 0 Then
      lblChanged.Caption = Format$(mDict.Changed, "ddddd ttttt")
      lblChanged.ToolTipText = mDict.ChangedByUserLongName
   Else
      lblChanged.Caption = ""
      lblChanged.ToolTipText = ""
   End If
   If mDict.TranscribedDate <> 0 Then
      lblTranscribedDate.Caption = Format$(mDict.TranscribedDate, "ddddd ttttt")
   Else
      lblTranscribedDate.Caption = ""
   End If
   
   txtTxt.Text = mDict.Txt
   txtNote.Text = mDict.Note
   
   ShowFormCaption
   SetEnabled
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Dim Ok As Boolean
   Dim SoundL As Long

   Debug.Print "frmDict_Unload+"
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
      If mDict.StatusId < Recorded Then
         If Len(ucCloseChoice.ChoiceText(1)) > 0 Then
            ucCloseChoice.ChoiceValue = 1
         End If
      End If
   End If
   
   If Ok Then
      mDict.SoundDirty = mDict.SoundDirty Or ucDSSRecGUI.Dirty
      mDict.CurrentPos = mPos
      SoundL = ucDSSRecGUI.SoundLengthInSec
      If SoundL > 0 Then   'If we lost device, length can wrong 0. Don't save it
         mDict.SoundLength = ucDSSRecGUI.SoundLengthInSec
      End If
      
      mAutoRewind = ucDSSRecGUI.AutoRewind
      ucDSSRecGUI.StopAndClose
      Set ucDSSRecGUI.DSSRecorder = Nothing
      Set mDict = Nothing
      If Me.WindowState = 0 Then
         LastfrmDictLeft = Me.Left
         LastfrmDictTop = Me.Top
      End If
   Else
      MsgBox Client.Texts.Txt(1030101, "Uppgifterna är inte kompletta!"), vbCritical
      Cancel = True
   End If
   Debug.Print "frmDict_Unload-"

End Sub

Private Sub imgLess_Click()

   imgLess.Visible = False
   imgMore.Visible = True
   Me.Height = FormLowHeight
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


Private Sub lblNoteTitle_Click()

   Clipboard.Clear
   Clipboard.SetText txtNote.Text
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

Private Sub txtNote_Change()

   If Screen.ActiveControl Is txtNote Then
      mDict.Note = txtNote.Text
      mDict.InfoDirty = True
   End If
   SetEnabled
End Sub

Private Sub txtNote_GotFocus()

   SelectAllText ActiveControl
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

Private Sub txtTxt_GotFocus()

   SelectAllText ActiveControl
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

   ShowFormCaption Formated
   mPos = PosInMilliSec
End Sub

Private Sub ucDSSRecGUI_WarningLowInputWhenRecording(TimeWithLowInput As Long, MaxInput As Long)

   'Debug.Print "Warning: " & TimeWithLowInput & ":" & MaxInput
   If TimeWithLowInput > 0 Then
      picWarning.Visible = True
      SetWindowTopMostAndForeground Me
   Else
      picWarning.Visible = False
   End If
End Sub

Private Sub ucOrgTree_NewSelect(OrgId As Long, Txt As String)

   If Screen.ActiveControl Is ucOrgTree Then
      mDict.OrgId = OrgId
      mDict.OrgText = Txt
      mDict.InfoDirty = True
      
      Client.DictTypeMgr.FillCombo cboDictType, mDict.OrgId, mDict.DictTypeId, mDict.DictTypeIdNoDefault, True
      If cboDictType.ListIndex >= 0 Then
         Dim DictType As clsDictType
         Client.DictTypeMgr.GetFromId DictType, cboDictType.ItemData(cboDictType.ListIndex)
         mDict.DictTypeId = DictType.DictTypeId
         mDict.DictTypeText = DictType.DictTypeText
         Set DictType = Nothing
      Else
         mDict.DictTypeId = -1
      End If
      
      Client.PriorityMgr.FillCombo cboPriority, mDict.OrgId, mDict.PriorityId, True
      If cboPriority.ListIndex >= 0 Then
         Dim Priority As clsPriority
         Client.PriorityMgr.GetFromId Priority, cboPriority.ItemData(cboPriority.ListIndex)
         mDict.PriorityId = Priority.PriorityId
         mDict.PriorityText = Priority.PriortyText
         mDict.ExpiryDate = DateAdd("d", Priority.Days, mDict.Created)
         Set Priority = Nothing
         lblExpiryDate.Caption = Format$(mDict.ExpiryDate, "ddddd")
      Else
         mDict.PriorityId = -1
         lblExpiryDate.Caption = ""
      End If
      
   End If
   SetEnabled
End Sub
Private Sub SetEnabled()

   Dim Enbld As Boolean
   
   Enbld = Not mTextReadOnly
   ucOrgTree.Enabled = Enbld
   txtPatId.Enabled = mDict.AuthorId = Client.User.UserId Or mDict.AuthorId = 0 Or (mUserIsSysAdmin And mIsInChangeMode)
   txtPatName.Enabled = mDict.AuthorId = Client.User.UserId Or mDict.AuthorId = 0 Or (mUserIsSysAdmin And mIsInChangeMode)
   chkNoPatient.Enabled = Enbld
   cboDictType.Enabled = Enbld
   cboPriority.Enabled = Enbld
   txtTxt.Enabled = Enbld
   txtNote.Enabled = Enbld
   
   CheckMandatoryData
   
   Client.DictMgr.SaveTempDictationInfo mDict, tdiUpdateInfo
   ShowFormCaption
End Sub
Private Function CheckMandatoryData() As Boolean

   Dim Ok As Boolean

   If mDict Is Nothing Then
      Debug.Print "CheckMandatoryData Nothing"
      CheckMandatoryData = False
      Exit Function
   End If
   Debug.Print "CheckMandatoryData"
   
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
   If mDict.DictTypeId < 0 Then
      lblDictTypeMissing.Visible = True
      Ok = False
   Else
      lblDictTypeMissing.Visible = False
   End If
   If mDict.PriorityId < 0 Then
      lblPriorityMissing.Visible = True
      Ok = False
   Else
      lblPriorityMissing.Visible = False
   End If
   
   CheckMandatoryData = Ok
End Function
Private Sub ShowFormCaption(Optional FormattedPos As String)

   Dim s As String
   
   s = Client.SysSettings.PlayerCaption
   If Len(s) = 0 Then
      Me.Caption = FormattedPos
   Else
      s = ChangeParam(s, "PatId", mDict.Pat.PatIdFormatted)
      s = ChangeParam(s, "PatName", mDict.Pat.PatName)
      s = ChangeParam(s, "Pos", FormattedPos)
      s = ChangeParam(s, "DictType", mDict.DictTypeText)
      s = ChangeParam(s, "Priority", mDict.PriorityText)
      s = ChangeParam(s, "Org", mDict.OrgText)
      
      Me.Caption = s
   End If
End Sub
Private Function ChangeParam(ByVal s As String, ByVal Param As String, ByVal Value As String) As String

   ChangeParam = Replace(s, "%" & Param & "%", Value, 1, -1, vbTextCompare)
End Function
