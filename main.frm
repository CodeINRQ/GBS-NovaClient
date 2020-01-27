VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{A455B2A1-A33C-11D1-A8BD-002078104456}#1.0#0"; "CP5OCX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "CareTalk"
   ClientHeight    =   8520
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13410
   HelpContextID   =   1000000
   Icon            =   "main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   13410
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9240
      Picture         =   "main.frx":030A
      ScaleHeight     =   285
      ScaleWidth      =   1935
      TabIndex        =   14
      Top             =   108
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   600
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   8280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   8265
      Width           =   13410
      _ExtentX        =   23654
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Timer tmrCheckButtons 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10920
      Top             =   120
   End
   Begin VB.Timer tmrUpdateList 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10560
      Top             =   120
   End
   Begin CareTalk.ucOrgTree ucOrgTree 
      Height          =   7335
      HelpContextID   =   1000000
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   12938
   End
   Begin TabDlg.SSTab Tabs 
      Height          =   7335
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   10815
      Visible         =   0   'False
      _ExtentX        =   19076
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      Tab             =   4
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "Diktat"
      TabPicture(0)   =   "main.frx":080D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ucDictList"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Statistik"
      TabPicture(1)   =   "main.frx":0829
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ucStatList"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Historik"
      TabPicture(2)   =   "main.frx":0845
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ucHistList"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Sök"
      TabPicture(3)   =   "main.frx":0861
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ucSearch"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Administration"
      TabPicture(4)   =   "main.frx":087D
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "ucEditUser"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Systeminställningar"
      TabPicture(5)   =   "main.frx":0899
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ucEditOrg"
      Tab(5).Control(1)=   "ucEditGroup"
      Tab(5).Control(2)=   "ucEditSysSettings"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "main.frx":08B5
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ucVoiceXpress"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Demo"
      TabPicture(7)   =   "main.frx":0E17
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "ucDemo1"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Logg"
      TabPicture(8)   =   "main.frx":0E33
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "ucLoggList"
      Tab(8).ControlCount=   1
      Begin CareTalk.ucDemo ucDemo1 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7011
      End
      Begin CareTalk.ucVoiceXpress ucVoiceXpress 
         Height          =   4095
         HelpContextID   =   1170000
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7223
      End
      Begin CareTalk.ucEditSysSettings ucEditSysSettings 
         Height          =   2655
         HelpContextID   =   1100000
         Left            =   -74880
         TabIndex        =   9
         Top             =   4560
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5530
      End
      Begin CareTalk.ucEditUser ucEditUser 
         Height          =   6735
         HelpContextID   =   1040000
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   10575
         _ExtentX        =   15055
         _ExtentY        =   4048
      End
      Begin CareTalk.ucEditGroup ucEditGroup 
         Height          =   2295
         HelpContextID   =   1100000
         Left            =   -74880
         TabIndex        =   7
         Top             =   2280
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4048
      End
      Begin CareTalk.ucEditOrg ucEditOrg 
         Height          =   1815
         HelpContextID   =   1100000
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2566
      End
      Begin CareTalk.ucStatList ucStatList 
         Height          =   6735
         HelpContextID   =   1160000
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   11880
      End
      Begin CareTalk.ucDictList ucDictList 
         Height          =   6735
         HelpContextID   =   1080000
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   11880
      End
      Begin CareTalk.ucSearch ucSearch 
         Height          =   4695
         HelpContextID   =   1150000
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   4935
         _ExtentX        =   9975
         _ExtentY        =   6376
      End
      Begin CareTalk.ucHistList ucHistList 
         Height          =   6735
         HelpContextID   =   1140000
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   11880
      End
      Begin CareTalk.ucLoggList ucLoggList 
         Height          =   6735
         HelpContextID   =   1330000
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   11880
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13410
      _ExtentX        =   23654
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Spela in"
            Object.Tag             =   "1000701"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Voice Xpress"
            Object.Tag             =   "1000702"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Importera diktat från portabel diktafon"
            Object.Tag             =   "1000703"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7320
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":0E4F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":1351
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":3B03
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":3C15
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":3D27
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":3E39
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "main.frx":3F4B
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin CompplusLib.MhZip MhZip 
      Left            =   8280
      Top             =   7920
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Overwrite       =   1
      Prompts         =   0   'False
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Arkiv"
      Index           =   10
      Tag             =   "1000201"
      Begin VB.Menu mnuFile 
         Caption         =   "Importera diktat..."
         Index           =   5
         Tag             =   "1000102"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Avsluta"
         Index           =   10
         Tag             =   "1000101"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Hjälp"
      Index           =   40
      Tag             =   "1000202"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Hjälp om CareTalk"
         Index           =   1
         Tag             =   "1000302"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Om CareTalk..."
         Index           =   10
         Tag             =   "1000301"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
      (ByVal hwnd As Long, ByVal lpHelpFile As String, _
      ByVal wCommand As Long, ByVal dwData As Long) As Long

'Help Constants
Private Const HELP_CONTEXT = &H1           'Display topic in ulTopic
Private Const HELP_QUIT = &H2              'Terminate help
Private Const HELP_INDEX = &H3             'Display index
Private Const HELP_CONTENTS = &H3
Private Const HELP_HELPONHELP = &H4        'Display help on using help
Private Const HELP_SETINDEX = &H5          'Set the current Index for multi index help
Private Const HELP_SETCONTENTS = &H5
Private Const HELP_CONTEXTPOPUP = &H8
Private Const HELP_FORCEFILE = &H9
Private Const HELP_KEY = &H101             'Display topic for keyword in offabData
Private Const HELP_COMMAND = &H102
Private Const HELP_PARTIALKEY = &H105      'call the search engine in winhelp

Public Event OnOpenDictation(Dict As clsDict)
Public Event OnCloseDictation(Dict As clsDict)
Public Event OnNewDictation(Dict As clsDict)
Public Event OnCreateDictation()
Public Event OnLogon()
Public Event OnLogout()
Public Event OnOrgChanged()

Private Const tabDictList = 0
Private Const tabStatList = 1
Private Const tabHistList = 2
Private Const tabSearch = 3
Private Const tabAdmin = 4
Private Const tabSysSettings = 5
Private Const tabVoiceXpress = 6
Private Const tabDemo = 7
Private Const tabLoggList = 8

Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long

Private CurrentUIStatus As New clsUIStatus
Private UIStatusStack As New clsStack

Private WithEvents mClient As clsClient
Attribute mClient.VB_VarHelpID = -1

Private WithEvents mDSSRec As CareTalkDSSRec3.DSSRecorder
Attribute mDSSRec.VB_VarHelpID = -1
Private WithEvents mVx As clsVoiceXpress
Attribute mVx.VB_VarHelpID = -1
Private mDictCloseChoice As Integer
Private WithEvents mDictForm As frmDict
Attribute mDictForm.VB_VarHelpID = -1
Private WithEvents mPopupForm As frmPopup
Attribute mPopupForm.VB_VarHelpID = -1
Public CurrentOrg As Long
Private LastOrgidForNewDictation As Long
Private EditDictDialogShown As Boolean
Private IsDictButtonPressed As Boolean
Public IsRecNewFromAPI As Boolean
Public IsPlayFromAPI As Boolean
Private DictFormSettings As New clsStringStore
Private RecordingAllowed As Boolean
Private VoiceXpressAllowed As Boolean

Private UIBusy As Boolean
Private defProgBarHwnd  As Long

Private Declare Function SetParent Lib "user32" _
  (ByVal hWndChild As Long, _
   ByVal hWndNewParent As Long) As Long
   
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   Dim K As Integer
   Dim Sh As Integer
   
   Sh = Shift And 7
   K = Sh * 256 + (KeyCode And 255)
   Select Case K
      Case Client.SysSettings.PlayerKeyRec
         If Me.Toolbar1.Buttons(1).Visible Then
            RecordNewDictation CurrentOrg = 30005
         End If
   End Select
End Sub

Private Sub Form_Load()

   Dim I As Integer
   Dim Ver As String
   Dim LoginResult As Integer
   
   On Error GoTo frmMain_Form_Load_Err
   
   StartUpFormMainIsLoaded = 1
   
   ApplicationVersion = App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000")
   GlobalCommandLine = Command$
   
   If App.PrevInstance Then
      GotoPrevInstance
   End If

   Debug.Print App.StartMode
   
   
   SetUpStatusBar
   
   Me.Show
   
   Set Client = New clsClient
   Set mClient = Client

   'We login just to get authenticationmethod and some settings
   If Not Client.Server.DictationStorageOpen("", "") Then
      ErrorHandleExplicit "1000421", "", 1000421, "CareTalk databas kan inte öppnas", True
   End If
   Client.SysSettings.Init "CT"

   Client.CultureLanguage = Client.SysSettings.CultureDefaultLanguage
   Client.Texts.NewLanguage Client.CultureLanguage

   Do While LoginResult < 100
      LoginResult = Client.UserMgr.LoginUser()
      Select Case LoginResult
         Case 0:
            Client.LoggMgr.Insert 1320102, LoggLevel_UserInfo, 0, Client.User.LoggData
            Exit Do
         Case 1:
            Client.LoggMgr.Insert 1320101, LoggLevel_UserFailure, 0, Client.User.LoggData
            MsgBox Client.Texts.Txt(1000401, "Ditt konto är låst. Vänta en stund och försök senare!"), vbExclamation
         Case 2, 3:
            Client.LoggMgr.Insert 1320103, LoggLevel_UserFailure, 0, Client.User.LoggData
            MsgBox Client.Texts.Txt(1000402, "Felaktiga inloggningsuppgifter. Försök igen!"), vbExclamation
         Case 4:
            Client.LoggMgr.Insert 1320104, LoggLevel_UserFailure, 0, Client.User.LoggData
            MsgBox Client.Texts.Txt(1000423, "Lösenordet kunde inte bytas"), vbExclamation
            Exit Do
         Case Else
            Client.LoggMgr.Insert 1320105, LoggLevel_UserFailure, 0, Client.User.LoggData
            MsgBox Client.Texts.Txt(1000422, "Inloggningen misslyckades"), vbExclamation
            Unload Me
            End
      End Select
   Loop
   
   StatusBar.Panels(4).Text = Client.User.ShortName
   
   ucDictList.RestoreSettings Client.Server.ReadUserData("CT", "DL", "", Ver)
   
   DictFormSettings.Serialized = Client.Server.ReadUserData("CT", "DF", "", Ver)

   Client.DictTypeMgr.Init
   Client.GroupMgr.Init
   Client.UserMgr.Init
   Client.PriorityMgr.Init
   Client.ExtSystemMgr.Init
   Client.EventMgr.Init
   
   Client.EventMgr.OnAppEvent "OnLogin"
   RaiseEvent OnLogon
   
   Client.DSSRec.GetHardWare Client.Hw
   Set mDSSRec = Client.DSSRec
        
   Set mPopupForm = New frmPopup
   
   ShowOrgTree False, True, False
   If Client.OrgMgr.CheckUserRole(0, "A") Then
      ucOrgTree.PickOrgId 30010
   Else
      ucOrgTree.PickOrgId 30025
   End If
   
   Set mVx = Client.VoiceXpress
   ucVoiceXpress.Init mVx
   ucSearch.Init
   
   CheckHardware
   mDSSRec.CheckLicens RecordingAllowed

   Client.DoBatchUpdates
   
   LastOrgidForNewDictation = Client.User.HomeOrgId
   tmrUpdateList.Enabled = True
   tmrCheckButtons.Enabled = True
   
   SetVisibleTabs
   frmMain.Tabs.Visible = True
   ReadyForApiCalls = True
   StartUpFormMainIsLoaded = 2
   Exit Sub
   
frmMain_Form_Load_Err:
   ErrorHandle "1000420", Err, 1000420, "CareTalk kan inte startas", False
   End
End Sub
Private Sub CheckHardware()

   Dim NewValue As Gru_Harware
   Static NotFirst As Boolean
   
   Client.DSSRec.GetHardWare NewValue
   If Client.Hw <> NewValue Or Not NotFirst Then
      Client.Hw = NewValue
      NotFirst = True
      RecordingAllowed = (Client.Hw = GRU_HW_RECORD) And Client.OrgMgr.CheckUserRole(0, "A")
      
      If Client.Hw = GRU_HW_RECORD Then
         StatusBar.Panels(5).Text = Client.Texts.Txt(1000425, "Inspelning")
      ElseIf Client.Hw = GRU_HW_TYPIST Then
         StatusBar.Panels(5).Text = Client.Texts.Txt(1000426, "Uppspelning")
      Else
         StatusBar.Panels(5).Text = ""
      End If
      
      Me.Toolbar1.Buttons(1).Visible = RecordingAllowed
      Me.Toolbar1.Buttons(5).Visible = Client.OrgMgr.CheckUserRole(0, "A") And Client.SysSettings.ImportAllowTool
      Me.mnuFile(5).Visible = Client.OrgMgr.CheckUserRole(0, "A") And Client.SysSettings.ImportAllowMenu
   
      VoiceXpressAllowed = Client.SysSettings.VoiceExpressActive And Client.Hw = GRU_HW_RECORD And mVx.VxInstalled
   
      If VoiceXpressAllowed Then
         Me.Toolbar1.Buttons(3).Visible = Client.SysSettings.VoiceExpressShowInToolbar
      Else
         Me.Toolbar1.Buttons(3).Visible = False
      End If
   End If
End Sub

Private Sub SetVisibleTabs()

   Dim Ver As String

   frmMain.Tabs.TabCaption(tabDictList) = Client.Texts.Txt(1000403, "Diktatlista")
   frmMain.Tabs.TabCaption(tabStatList) = Client.Texts.Txt(1000404, "Statistik")
   frmMain.Tabs.TabCaption(tabHistList) = Client.Texts.Txt(1000405, "Historik")
   frmMain.Tabs.TabCaption(tabSearch) = Client.Texts.Txt(1000406, "Sök")
   frmMain.Tabs.TabCaption(tabAdmin) = Client.Texts.Txt(1000407, "Administration")
   frmMain.Tabs.TabCaption(tabSysSettings) = Client.Texts.Txt(1000408, "Systeminställningar")
   frmMain.Tabs.TabCaption(tabVoiceXpress) = ""
   frmMain.Tabs.TabCaption(tabDemo) = Client.Texts.Txt(1000409, "Demo")
   frmMain.Tabs.TabCaption(tabLoggList) = Client.Texts.Txt(1000424, "Logg")
   
   
   frmMain.Tabs.TabVisible(tabDictList) = True
   frmMain.Tabs.TabVisible(tabSearch) = True
   If Client.OrgMgr.CheckUserRole(0, "S") Then
      frmMain.Tabs.TabVisible(tabStatList) = True
      ucStatList.Init
      ucStatList.RestoreSettings Client.Server.ReadUserData("CT", "SL", "", Ver)
         
      frmMain.Tabs.TabVisible(tabHistList) = True
      ucHistList.Init
      ucHistList.RestoreSettings Client.Server.ReadUserData("CT", "HL", "", Ver)
   
   
      frmMain.Tabs.TabVisible(tabAdmin) = True
      ucEditUser.Init
   Else
      frmMain.Tabs.TabVisible(tabStatList) = False
      frmMain.Tabs.TabVisible(tabHistList) = False
      frmMain.Tabs.TabVisible(tabAdmin) = False
   End If
   If Client.OrgMgr.CheckUserRole(0, "I") Then
      frmMain.Tabs.TabVisible(tabSysSettings) = True
      ucEditGroup.Init
      Set ucEditSysSettings.Settings = Client.SysSettings.Store
      
      frmMain.Tabs.TabVisible(tabLoggList) = True
      ucLoggList.Init
      ucLoggList.RestoreSettings Client.Server.ReadUserData("CT", "LL", "", Ver)
      
      frmMain.Tabs.TabVisible(tabDemo) = Client.SysSettings.DemoShowTab
      
   Else
      frmMain.Tabs.TabVisible(tabSysSettings) = False
      frmMain.Tabs.TabVisible(tabLoggList) = False
      frmMain.Tabs.TabVisible(tabDemo) = False
   End If
   frmMain.Tabs.TabVisible(tabVoiceXpress) = Client.SysSettings.VoiceExpressShowTab
   frmMain.Tabs.Tab = tabDictList
End Sub

Private Sub ShowOrgTree(ShowAll As Boolean, ShowVirtual As Boolean, JustSupervisorRights As Boolean)

   Dim I As Integer
   Dim Org As clsOrg
   Dim StartOrgId As Long
   Dim EnabledDueToSupervisory As Boolean

   Client.OrgMgr.Init ShowAll
   ucOrgTree.Clear
   For I = 0 To Client.OrgMgr.Count - 1
      Client.OrgMgr.GetSortedOrg Org, I
      If Org.ShowInTree Or ShowAll Then
         If Org.ShowBelow Or Org.DictContainer Or ShowAll Then
            EnabledDueToSupervisory = Client.OrgMgr.CheckUserRole(Org.OrgId, "S") Or Not JustSupervisorRights
            If EnabledDueToSupervisory Then
               ucOrgTree.AddNode Org.OrgId, Org.ShowParent, Org.OrgText, 1, True
            Else
               ucOrgTree.AddNode Org.OrgId, Org.ShowParent, Org.OrgText, 5, False
            End If
         Else
            ucOrgTree.AddNode Org.OrgId, Org.ShowParent, Org.OrgText, 5, False
         End If
      End If
   Next I
   
   If ShowVirtual Then
      ucOrgTree.AddNode 30000, 0, Client.Texts.Txt(1000410, "Mina diktat"), 3, True
      ucOrgTree.AddNode 30050, 0, Client.Texts.Txt(1000416, "Sökresultat"), 3, True
      ucOrgTree.AddNode 30005, 0, Client.Texts.Txt(1000419, "Aktuell patient"), 3, True
      
      If Client.OrgMgr.CheckUserRole(0, "A") Then
         ucOrgTree.AddNode 30010, 30000, Client.Texts.Txt(1000411, "Under inspelning"), 3, True
         ucOrgTree.AddNode 30020, 30000, Client.Texts.Txt(1000412, "Inspelade"), 3, True
      End If
      ucOrgTree.AddNode 30025, 30000, Client.Texts.Txt(1000413, "Under utskrift"), 3, True
      If Client.SysSettings.UseAuthorsSign Then
         ucOrgTree.AddNode 30030, 30000, Client.Texts.Txt(1000414, "För signering"), 3, True
      End If
      ucOrgTree.AddNode 30040, 30000, Client.Texts.Txt(1000415, "Utskrivna"), 3, True
   End If
End Sub

Private Sub Form_Resize()

   If Me.WindowState <> vbMinimized Then
      If Me.Width < 6200 Then
         Me.Width = 6200
      Else
         Me.Tabs.Width = Me.Width - 11 * 240
         Me.ucDictList.Width = Me.Tabs.Width - 1 * 240
         Me.ucStatList.Width = Me.Tabs.Width - 1 * 240
         Me.ucHistList.Width = Me.Tabs.Width - 1 * 240
         Me.ucLoggList.Width = Me.Tabs.Width - 1 * 240
         Me.picLogo.Left = Me.Width - Me.picLogo.Width - 300
         'Me.imgLogo.Left = Me.Width - Me.imgLogo.Width - 300
         'Me.Toolbar1.Width = Me.imgLogo.Left - 200
      End If
      If Me.Height < 5200 Then
         Me.Height = 5200
      Else
         Me.Tabs.Height = Me.Height - 5 * 240 - Me.StatusBar.Height
         Me.ucDictList.Height = Me.Tabs.Height - 3 * 240
         Me.ucStatList.Height = Me.Tabs.Height - 3 * 240
         Me.ucHistList.Height = Me.Tabs.Height - 3 * 240
         Me.ucLoggList.Height = Me.Tabs.Height - 3 * 240
         Me.ucEditUser.Height = Me.Tabs.Height - 3 * 240
         Me.ucOrgTree.Height = Me.Height - Me.ucOrgTree.Top - 4 * 240 - Me.StatusBar.Height
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Dim Res As Integer
          
   On Error Resume Next
   StartUpFormMainIsLoaded = 1

   Res = WinHelp(frmMain.hwnd, App.HelpFile, HELP_QUIT, 0&)
   
   If Client.User.UserId > 0 Then
      Client.EventMgr.OnAppEvent "OnLogout"
      RaiseEvent OnLogout
   End If
   
   If VoiceXpressAllowed Then
      mVx.Activate = False
   End If
   'On Error GoTo 0
   If Client.Server.StorageOpened And Client.User.UserId > 0 Then
      If frmMain.Tabs.TabVisible(tabHistList) Then
         Client.Server.WriteUserData "CT", "HL", ucHistList.GetSetting()
      End If
      If frmMain.Tabs.TabVisible(tabStatList) Then
         Client.Server.WriteUserData "CT", "SL", ucStatList.GetSetting()
      End If
      Client.Server.WriteUserData "CT", "DL", ucDictList.GetSetting()
      Client.Server.WriteUserData "CT", "DF", DictFormSettings.Serialized
   End If
   Client.LoggMgr.Insert 1320106, LoggLevel_UserInfo, 0, Client.User.LoggData
   Set mClient = Nothing
   Set Client = Nothing
   StartUpFormMainIsLoaded = 0
End Sub

Private Sub mClient_UIStatusClear()

   UIStatusClear
End Sub

Private Sub mClient_UIStatusProgress(Total As Long, Left As Long)

   UIStatusProgress Total, Left
End Sub

Private Sub mClient_UIStatusSet(StatusText As String, Busy As Boolean)

   UIStatusSet StatusText, Busy
End Sub

Private Sub mClient_UIStatusSetSub(SubText As String)

   UIStatusSetSub SubText
End Sub

Private Sub mDictForm_CloseChoiceSelected(Index As Integer)

   mDictCloseChoice = Index
End Sub

Private Sub mDSSRec_GruEvent(EventType As CareTalkDSSRec3.Gru_Event, Data As Long)

   'Debug.Print "GruEvent " & CInt(EventType)
   Select Case EventType
      Case GRU_BUTTONPRESS
         Select Case Data
            Case GRU_BUT_DICT, GRU_BUT_INSERT
               If Not EditDictDialogShown Then
                  If RecordingAllowed Then
                     IsDictButtonPressed = True
                  End If
               End If
            Case GRU_BUT_INDEX
               If Not EditDictDialogShown Then
                  If Client.SysSettings.VoiceExpressActivateOnIndexButton Then
                     If frmMain.Toolbar1.Buttons(3).Value = tbrPressed Then
                        frmMain.Toolbar1.Buttons(3).Value = tbrUnpressed
                     Else
                        mDSSRec.SetMicRecordMode True
                        frmMain.Toolbar1.Buttons(3).Value = tbrPressed
                     End If
                     mVx.Activate = frmMain.Toolbar1.Buttons(3).Value = tbrPressed
                  End If
               End If
         End Select
      Case GRU_HWCHANGED
         CheckHardware
   End Select
End Sub

Private Sub mnuFile_Click(Index As Integer)

   Select Case Index
      Case 5
         ImportNewDictation
      Case 10
         Unload Me
   End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)

   Dim Res As Integer

   Select Case Index
      Case 1
         On Error Resume Next
         Res = WinHelp(frmMain.hwnd, App.HelpFile, HELP_CONTENTS, 0&)
      Case 10
         frmAbout.Show vbModal
   End Select
End Sub

Private Sub mPopupForm_Choise(MenuIndex As Integer, ItemIndex As Integer, Id As Long)

   Select Case MenuIndex
      Case 0
         Select Case ItemIndex
            Case 10
               Client.DictMgr.UnlockDict Id
            Case 20
               frmDictAudit.DictId = Id
               frmDictAudit.Show vbModal
         End Select
   End Select
End Sub

Private Sub mVx_ChangeAppState(NewValue As vxAppStateEnum)

   If NewValue = vxAppStateQuiting Then
      frmMain.Toolbar1.Buttons(3).Value = tbrUnpressed
   End If
End Sub

Private Sub mVx_ChangeListening(NewValue As vxListeningEnum)

   If NewValue = vxListeningOn Then
      If EditDictDialogShown Then
         mVx.vxListening = vxListeningOff
      Else
         frmMain.Toolbar1.Buttons(3).Value = tbrPressed
         mDSSRec.SetMicRecordMode 1
      End If
   Else
      frmMain.Toolbar1.Buttons(3).Value = tbrUnpressed
      If Not EditDictDialogShown Then
         mDSSRec.SetMicRecordMode 0
      End If
   End If
End Sub

Private Sub Tabs_Click(PreviousTab As Integer)

   UpdateCurrentView
End Sub

Private Sub tmrCheckButtons_Timer()

   If IsDictButtonPressed Then
      IsDictButtonPressed = False
      RecordNewDictation CurrentOrg = 30005
   End If
   If IsRecNewFromAPI Then
      IsRecNewFromAPI = False
      If RecordingAllowed Then
         If Not EditDictDialogShown Then
            RecordNewDictation CurrentOrg = 30005
         End If
      End If
   End If
   If IsPlayFromAPI Then
      IsPlayFromAPI = False
      If Not EditDictDialogShown Then
         EditExistingDictation Client.PlayDictIdFromAPI
      End If
   End If
End Sub

Private Sub tmrUpdateList_Timer()

   Static TimeForUpdates As New clsTimeKeeping
   Dim MeanTime As Double

   If Not EditDictDialogShown Then
      TimeForUpdates.StartMeasure
      UpdateCurrentView
      TimeForUpdates.StopMeasure
      
      MeanTime = TimeForUpdates.SlidingMeanValue(5)
      If MeanTime > 0.5 Then
         Debug.Print "Interval 5000"
         tmrUpdateList.Interval = 5000
      ElseIf MeanTime > 0.3 Then
         Debug.Print "Interval 3000"
         tmrUpdateList.Interval = 3000
      Else
         Debug.Print "Interval 2000"
         tmrUpdateList.Interval = 2000
      End If
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

   Select Case Button.Index
      Case 1
         RecordNewDictation CurrentOrg = 30005
      Case 3
         mVx.Activate = frmMain.Toolbar1.Buttons(3).Value = tbrPressed
      Case 5
         ImportNewDictation
   End Select
End Sub

Private Sub ucDemo1_UIStatusClear()

   UIStatusClear
End Sub

Private Sub ucDemo1_UIStatusProgress(Total As Long, Left As Long)

   UIStatusProgress Total, Left
End Sub

Private Sub ucDemo1_UIStatusSet(StatusText As String, Busy As Boolean)

   UIStatusSet StatusText, Busy
End Sub

Private Sub ucDemo1_UIStatusSetSub(SubText As String)

   UIStatusSetSub SubText
End Sub

Private Sub ucDictList_DblClick(DictId As Long)

   EditExistingDictation DictId
End Sub

Private Sub ucDictList_RightClick(DictId As Long)

   If Client.OrgMgr.CheckUserRole(0, "I") Then
      On Error Resume Next
      mPopupForm.Id = DictId
      PopupMenu mPopupForm.mnuPopup(0)
   End If
End Sub

Private Sub ucEditOrg_OrgSaved(Org As clsOrg)

   ShowOrgTree True, False, True
   CurrentOrg = Org.OrgId
   ucOrgTree.PickOrgId CurrentOrg
End Sub

Private Sub ucEditSysSettings_SaveClicked(Settings As clsStringStore)

   Set Client.SysSettings.Store = Settings
   Client.SysSettings.Save "CT"
   Client.SysSettings.Init "CT"
End Sub

Private Sub ucOrgTree_NewSelect(OrgId As Long, Txt As String)

   CurrentOrg = OrgId
   Me.Caption = Client.Texts.Txt(1000417, "CareTalk") & " - " & Txt
   UpdateCurrentView
End Sub
Private Sub UpdateCurrentView()

   Static PreviousTab As Integer
   Static CurrOrgIdWhenShowAll As Long
   Static CurrOrgIdWhenNotShowAll As Long
   Static PreviousOrg As Long
   
   UIBusy = True
   
   If PreviousOrg <> CurrentOrg Then
      Client.EventMgr.OnAppEvent "OnOrgChanged"
      RaiseEvent OnOrgChanged
      PreviousOrg = CurrentOrg
   End If
   
      UIStatusSet Client.Texts.Txt(1000418, "Mappen uppdateras"), False
      If PreviousTab <> Tabs.Tab Then
         Select Case PreviousTab
            Case tabSysSettings, tabStatList, tabHistList, tabAdmin
               CurrOrgIdWhenNotShowAll = CurrentOrg
            Case Else
               CurrOrgIdWhenShowAll = CurrentOrg
         End Select
'         If PreviousTab = tabSysSettings Then
'            PreviousTab = Tabs.Tab                 'must be here
'            CurrOrgIdWhenShowAll = CurrentOrg
'            ShowOrgTree False, True, False
'            ucOrgTree.PickOrgId CurrOrgIdWhenNotShowAll
'         End If
         PreviousTab = Tabs.Tab                    'must be here also
         Select Case Tabs.Tab
            Case tabSysSettings
               ShowOrgTree True, False, True
               ucOrgTree.PickOrgId CurrOrgIdWhenNotShowAll
            Case tabStatList, tabHistList, tabAdmin
               ShowOrgTree True, False, True
               ucOrgTree.PickOrgId CurrOrgIdWhenNotShowAll
            Case Else
               ShowOrgTree False, True, False
               ucOrgTree.PickOrgId CurrOrgIdWhenShowAll
         End Select
'         If Tabs.Tab = tabSysSettings Then
'            CurrOrgIdWhenNotShowAll = CurrentOrg
'            ShowOrgTree True, False, True
'            ucOrgTree.PickOrgId CurrOrgIdWhenShowAll
'         End If
      End If
      Select Case Tabs.Tab
         Case 0
            If CurrentOrg > 0 Then
               Me.ucDictList.GetData CurrentOrg
            End If
         Case 1
            If CurrentOrg > 0 Then
               Me.ucStatList.GetData CurrentOrg
            End If
         Case 2
            If CurrentOrg > 0 Then
               Me.ucHistList.GetData CurrentOrg
            End If
         Case 5
            If CurrentOrg > 0 Then
               Me.ucEditOrg.OrgSelected CurrentOrg
               Me.ucEditGroup.NewOrg CurrentOrg
            End If
      End Select
      UIStatusClear
      
   UIBusy = False
End Sub
Private Sub RecordNewDictation(UseCurrPat As Boolean)

   Dim Dict As clsDict
   Static AllreadyStarted As Boolean

   On Error GoTo RecordNewDictation_Err
   WaitForUIBusy
   If AllreadyStarted Then Exit Sub
   If EditDictDialogShown Then Exit Sub
   AllreadyStarted = True
   EditDictDialogShown = True
   
   If VoiceXpressAllowed Then
      mVx.Activate = False
   End If
      
   Set Dict = New clsDict
   Client.DictMgr.CreateNew Dict
   
   Client.EventMgr.OnDictEvent "OnCreate", Dict
   RaiseEvent OnCreateDictation

   Dict.ExtSystem = Client.NewRecInfo.ExtSystem
   Dict.ExtDictId = Client.NewRecInfo.ExtDictId
   Dict.Pat.PatId = Client.NewRecInfo.PatId
   If Len(Dict.Pat.PatId) = 0 And UseCurrPat Then
      Dict.Pat.PatId = Client.CurrPatient.PatId
   End If
   Dict.Pat.PatName = Client.NewRecInfo.PatName
   If Len(Dict.Pat.PatName) = 0 And UseCurrPat Then
      Dict.Pat.PatName = Client.CurrPatient.PatName
   End If
   If Client.NewRecInfo.DictTypeId > 0 Then
      Dict.DictTypeId = Client.NewRecInfo.DictTypeId
   End If
   If Client.NewRecInfo.OrgId > 0 Then
      If Client.OrgMgr.CheckUserRole(Client.NewRecInfo.OrgId, "A") Then
         Dict.OrgId = Client.NewRecInfo.OrgId
      Else
         Dict.OrgId = LastOrgidForNewDictation
      End If
   Else
      Dict.OrgId = LastOrgidForNewDictation
   End If
   If Client.NewRecInfo.PrioId > 0 Then
      Dict.PriorityId = Client.NewRecInfo.PrioId
   End If
   Set Client.NewRecInfo = Nothing
   
   Set mDictForm = New frmDict
   Load mDictForm
   mDictForm.RestoreSettings DictFormSettings
   mDictForm.EditDictation Dict, True
   mDictForm.CloseText(0) = Client.Texts.Txt(1000501, "Radera diktatet")
   mDictForm.CloseTip(0) = Client.Texts.ToolTip(1000501, "Inspelningen kastas!")
   mDictForm.CloseText(1) = Client.Texts.Txt(1000502, "Fortsätt diktera senare")
   mDictForm.CloseTip(1) = Client.Texts.ToolTip(1000502, "Under inspelning")
   mDictForm.CloseText(2) = Client.Texts.Txt(1000503, "Klart för utskrift")
   mDictForm.CloseTip(2) = ""
   ShowWindow Me.hwnd, SW_Hide
   mDictForm.Show vbModal
   ShowWindow Me.hwnd, SW_Show
   Select Case mDictCloseChoice
      Case 0
         'no action
      Case 1
         LastOrgidForNewDictation = Dict.OrgId
         Dict.StatusId = 20
         Client.DictMgr.CheckInNew Dict
         Client.EventMgr.OnDictEvent "OnNew", Dict
         RaiseEvent OnNewDictation(Dict)
      Case 2
         LastOrgidForNewDictation = Dict.OrgId
         Dict.StatusId = 30
         Client.DictMgr.CheckInNew Dict
         Client.EventMgr.OnDictEvent "OnNew", Dict
         RaiseEvent OnNewDictation(Dict)
   End Select
   mDictForm.SaveSettings DictFormSettings
   Unload mDictForm
   Set mDictForm = Nothing
   EditDictDialogShown = False
   AllreadyStarted = False
   Exit Sub
   
RecordNewDictation_Err:
   ErrorHandle "1000504", Err, 1000504, "Ett fel har uppstått", True
   Resume Next
End Sub
Private Sub ImportNewDictation()

   Dim Dict As clsDict
   Static AllreadyStarted As Boolean
   Dim ImportFileName As String

   WaitForUIBusy
   If AllreadyStarted Then Exit Sub
   If EditDictDialogShown Then Exit Sub
   AllreadyStarted = True
   EditDictDialogShown = True
   
   If VoiceXpressAllowed Then
      mVx.Activate = False
   End If
   
   ImportFileName = GetImportFileName()
   If Len(ImportFileName) > 0 Then
      Set Dict = New clsDict
      Client.DictMgr.CreateNew Dict
      
      Client.EventMgr.OnDictEvent "OnCreate", Dict
      RaiseEvent OnCreateDictation
      
      If CopyImportFileToTempStorage(ImportFileName, Dict.LocalFilename) Then
         KillFileIgnoreError ImportFileName
         Dict.OrgId = LastOrgidForNewDictation
         Set mDictForm = New frmDict
         Load mDictForm
         mDictForm.RestoreSettings DictFormSettings
         mDictForm.EditDictation Dict, False
         mDictForm.CloseText(0) = Client.Texts.Txt(1000501, "Radera diktatet")
         mDictForm.CloseTip(0) = Client.Texts.ToolTip(1000501, "Inspelningen kastas!")
         mDictForm.CloseText(1) = Client.Texts.Txt(1000502, "Fortsätt diktera senare")
         mDictForm.CloseTip(1) = Client.Texts.ToolTip(1000502, "Under inspelning")
         mDictForm.CloseText(2) = Client.Texts.Txt(1000503, "Klart för utskrift")
         mDictForm.CloseTip(2) = Client.Texts.ToolTip(1000503, "")
         ShowWindow Me.hwnd, SW_Hide
         mDictForm.Show vbModal
         ShowWindow Me.hwnd, SW_Show
         Select Case mDictCloseChoice
            Case 0
               'no action
            Case 1
               LastOrgidForNewDictation = Dict.OrgId
               Dict.StatusId = 20
               Client.DictMgr.CheckInNew Dict
               Client.EventMgr.OnDictEvent "OnNew", Dict
               RaiseEvent OnNewDictation(Dict)
            Case 2
               LastOrgidForNewDictation = Dict.OrgId
               Dict.StatusId = 30
               Client.DictMgr.CheckInNew Dict
               Client.EventMgr.OnDictEvent "OnNew", Dict
               RaiseEvent OnNewDictation(Dict)
         End Select
         mDictForm.SaveSettings DictFormSettings
         Unload mDictForm
         Set mDictForm = Nothing
      End If
   End If
   EditDictDialogShown = False
   AllreadyStarted = False
End Sub
Public Sub EditExistingDictation(DictId As Long)

   Dim Dict As clsDict
   Dim Discard As Boolean
   Dim IsUserTranscriber As Boolean
   Dim IsUserAuthor As Boolean
   
   On Error GoTo EditExistingDictation_Err
   If EditDictDialogShown Then Exit Sub
   EditDictDialogShown = True
   
   WaitForUIBusy
   If VoiceXpressAllowed Then
      mVx.Activate = False
   End If
   
   
   If Client.DictMgr.CheckOut(Dict, DictId, True) = 0 Then
      
      If Client.OrgMgr.CheckUserAllowListening(Dict.OrgId) Then
         Client.EventMgr.OnDictEvent "OnOpen", Dict
         RaiseEvent OnOpenDictation(Dict)
      
         IsUserAuthor = Client.OrgMgr.CheckUserRole(Dict.OrgId, "A")
         IsUserTranscriber = Client.OrgMgr.CheckUserRole(Dict.OrgId, "T")
         
         Set mDictForm = New frmDict
         Load mDictForm
         mDictForm.RestoreSettings DictFormSettings
         Client.Trace.AddRow Trace_Level_Full, "10006", "10006C", "", CStr(Dict.DictId), CStr(Dict.StatusId)
         mDictForm.EditDictation Dict, False
         Client.Trace.AddRow Trace_Level_Full, "10006", "10006D", "", CStr(Dict.DictId), CStr(Dict.StatusId)
         If IsUserAuthor And Dict.AuthorId = Client.User.UserId Then
            If Dict.StatusId < Recorded Then
               Discard = ShowAndSetNewStatus(Dict, Client.Texts.Txt(1000601, "Radera hela diktatet"), _
                                                   Client.Texts.ToolTip(1000601, "Inspelningen kastas!"), _
                                                   SoundDeleted, _
                                                   Client.Texts.Txt(1000602, "Fortsätt diktera senare"), _
                                                   Client.Texts.ToolTip(1000602, "Under inspelning"), _
                                                   BeingRecorded, _
                                                   Client.Texts.Txt(1000603, "Klart för utskrift"), _
                                                   Client.Texts.ToolTip(1000603, "Diktatet klart för utskrift"), _
                                                   Recorded)
            ElseIf Dict.StatusId = Recorded Then
               Discard = ShowAndSetNewStatus(Dict, Client.Texts.Txt(1000604, "Ångra ändringar"), _
                                                   Client.Texts.ToolTip(1000604, "Lämna dikatet oförändrat"), _
                                                   0, _
                                                   Client.Texts.Txt(1000602, "Fortsätt diktera senare"), _
                                                   Client.Texts.ToolTip(1000602, "Under inspelning"), _
                                                   BeingRecorded, _
                                                   Client.Texts.Txt(1000603, "Klart för utskrift"), _
                                                   Client.Texts.ToolTip(1000603, "Diktatet klart för utskrift"), _
                                                   Recorded)
            ElseIf Dict.StatusId >= WaitForSign Then
               If Client.SysSettings.UseAuthorsSign Then
                  Discard = ShowAndSetNewStatus(Dict, Client.Texts.Txt(1000607, "Signera senare"), _
                                                      Client.Texts.ToolTip(1000607, "Lämna diktatet för signering senare"), _
                                                      0, _
                                                      "", "", WaitForSign, _
                                                      Client.Texts.Txt(1000608, "Signerat"), _
                                                      Client.Texts.ToolTip(1000608, "Signerat, diktatet kan raderas"), _
                                                      Transcribed)
               Else
                  Discard = ShowAndSetNewStatus(Dict, Client.Texts.Txt(1000609, "Stäng"), _
                                                      Client.Texts.ToolTip(1000609, "Diktatet kan inte ändras"), _
                                                      0, _
                                                      "", _
                                                      "", _
                                                      0, _
                                                      "", _
                                                      "", _
                                                      0)
               End If
            Else
               Discard = ShowAndSetNewStatus(Dict, Client.Texts.Txt(1000609, "Stäng"), _
                                                   Client.Texts.ToolTip(1000609, "Diktatet kan inte ändras"), _
                                                   0, _
                                                   "", _
                                                   "", _
                                                   0, _
                                                   "", _
                                                   "", _
                                                   0)
            End If
         ElseIf IsUserTranscriber And (Dict.TranscriberId = Client.User.UserId Or Dict.TranscriberId = 0) Then
            If Dict.StatusId >= Recorded And Dict.StatusId < WaitForSign Then
               If Client.SysSettings.UseAuthorsSign Then
                  Discard = ShowAndSetNewStatus(Dict, Client.Texts.Txt(1000610, "Ångra"), _
                                                      Client.Texts.ToolTip(1000610, "Lämna dikatet oförändrat"), _
                                                      0, _
                                                      Client.Texts.Txt(1000613, "Fortsätt utskrift senare"), _
                                                      Client.Texts.ToolTip(1000613, "Under utskrift"), _
                                                      BeingTrancribed, _
                                                      Client.Texts.Txt(1000611, "Utskriften klar"), _
                                                      Client.Texts.ToolTip(1000611, "Klart för signering"), _
                                                      WaitForSign)
               Else
                  Discard = ShowAndSetNewStatus(Dict, Client.Texts.Txt(1000610, "Ångra"), _
                                                      Client.Texts.ToolTip(1000610, "Lämna dikatet oförändrat"), _
                                                      0, _
                                                      Client.Texts.Txt(1000613, "Fortsätt utskrift senare"), _
                                                      Client.Texts.ToolTip(1000613, "Under utskrift"), _
                                                      BeingTrancribed, _
                                                      Client.Texts.Txt(1000612, "Utskriften klar"), _
                                                      Client.Texts.ToolTip(1000612, "Diktatet kan raderas"), _
                                                      Transcribed)
               End If
            Else
               Discard = ShowAndSetNewStatus(Dict, Client.Texts.Txt(1000609, "Stäng"), _
                                                   Client.Texts.ToolTip(1000609, "Diktatet kan inte ändras"), _
                                                   0, _
                                                   "", _
                                                   "", _
                                                   0, _
                                                   "", _
                                                   "", _
                                                   0)
            End If
         Else
            Discard = ShowAndSetNewStatus(Dict, Client.Texts.Txt(1000609, "Stäng"), _
                                                Client.Texts.ToolTip(1000609, "Diktatet kan inte ändras"), _
                                                0, _
                                                "", _
                                                "", _
                                                0, _
                                                "", _
                                                "", _
                                                0)
         End If
      End If
      Client.Trace.AddRow Trace_Level_Full, "10006", "10006A", "", CStr(Dict.DictId), CStr(Dict.StatusId)
      Client.DictMgr.CheckIn Dict, Discard
      Client.Trace.AddRow Trace_Level_Full, "10006", "10006B", "", CStr(Dict.DictId), CStr(Dict.StatusId)
         
      Client.EventMgr.OnDictEvent "OnClose", Dict
      RaiseEvent OnCloseDictation(Dict)
      
      mDictForm.SaveSettings DictFormSettings
      Unload mDictForm
      Set mDictForm = Nothing
      UpdateCurrentView
   End If
   EditDictDialogShown = False
   Exit Sub
   
EditExistingDictation_Err:
   ErrorHandle "1000614", Err, 1000614, "Ett fel har uppstått", True
   Resume Next
End Sub
Private Function ShowAndSetNewStatus(Dict As clsDict, _
                                                      Text1 As String, Tip1 As String, NewStatus1 As Integer, _
                                                      Text2 As String, Tip2 As String, NewStatus2 As Integer, _
                                                      Text3 As String, Tip3 As String, NewStatus3 As Integer) As Boolean

   Dim NewStatus As Integer

   mDictForm.CloseText(0) = Text1
   mDictForm.CloseTip(0) = Tip1
   mDictForm.CloseText(1) = Text2
   mDictForm.CloseTip(1) = Tip2
   mDictForm.CloseText(2) = Text3
   mDictForm.CloseTip(2) = Tip3
   ShowWindow Me.hwnd, SW_Hide
   mDictForm.Show vbModal
   ShowWindow Me.hwnd, SW_Show
   Select Case mDictCloseChoice
      Case 0
         NewStatus = NewStatus1
      Case 1
         NewStatus = NewStatus2
         Dict.StatusId = NewStatus2
         ShowAndSetNewStatus = False
      Case 2
         NewStatus = NewStatus3
   End Select
   If NewStatus <> 0 Then
      Dict.StatusId = NewStatus
      ShowAndSetNewStatus = False
   Else
      ShowAndSetNewStatus = True
   End If
End Function

Private Sub ucSearch_NewSearch(SearchFilter As clsFilter)

   UIStatusSet Client.Texts.Txt(1000427, "Sökning sker..."), True

      Set ucDictList.SearchFilter = SearchFilter
      ucOrgTree.PickOrgId 30050
      Tabs.Tab = 0
      
   UIStatusClear
End Sub
Private Sub SetUpStatusBar()

   Dim pnl As Panel
   Dim btn As Button
   Dim x As Long
   Dim pading As Long
   
  'create statusbar
   With StatusBar
      For x = 1 To 5
         Set pnl = .Panels.Add(, , "", sbrText)
         'If x = 4 Then
         '   pnl.Alignment = sbrRight
         'Else
            pnl.Alignment = sbrLeft
         'End If
         pnl.Bevel = sbrInset
         If x = 5 Then
            pnl.AutoSize = sbrSpring
         Else
            pnl.Width = 2800
         End If
      Next x
   End With
   
   With ProgressBar
      .Min = 0
      .Max = 100
      .Value = .Max
   End With

  'parent the progress bar in the status bar
   pading = 60
   AttachProgBar ProgressBar, StatusBar, 3, pading
   
  'change the bar colour
  ' Call SendMessage(ProgressBar.hwnd, _
  '                  PBM_SETBARCOLOR, _
  '                  0&, _
  '                  ByVal RGB(205, 0, 205))

   ProgressBar.Value = 0

End Sub
Private Function AttachProgBar(pb As ProgressBar, _
                               sb As StatusBar, _
                               nPanel As Long, _
                               pading As Long)
    
   If defProgBarHwnd = 0 Then
       
     'change the parent
      defProgBarHwnd = SetParent(pb.hwnd, sb.hwnd)
   
      With sb
      
        'adjust statusbar. Doing it this way
        'relieves the necessity of calculating
        'the statusbar position relative to the
        'top of the form. It happens so fast
        'the change is not seen.
         .Align = vbAlignTop
         .Visible = False
         
        'change, move, set size and re-show
        'the progress bar in the new parent
         With pb
            .Visible = False
            .Align = vbAlignNone
            .Appearance = ccFlat
            .BorderStyle = ccNone
            .Width = sb.Panels(nPanel).Width
            .Move (sb.Panels(nPanel).Left + pading), _
                 (sb.Top + pading), _
                 (sb.Panels(nPanel).Width - (pading * 2)), _
                 (sb.Height - (pading * 2))
                  
            .Visible = True
            .ZOrder 0
         End With
           
        'restore the statusbar to the
        'bottom of the form and show
         .Panels(nPanel).AutoSize = sbrNoAutoSize
         .Align = vbAlignBottom
         .Visible = True
         
       End With
      
    End If
       
End Function
Private Sub UIStatusSet(StatusText As String, Busy As Boolean)

   Dim Stat As New clsUIStatus
   
   Stat.Text1 = CurrentUIStatus.Text1
   Stat.Text2 = CurrentUIStatus.Text2
   Set Stat.ActiveControlAtBusy = CurrentUIStatus.ActiveControlAtBusy
   Stat.ActiveMousePointerAtBusy = CurrentUIStatus.ActiveMousePointerAtBusy
   Stat.Busy = CurrentUIStatus.Busy
   Stat.Progress = CurrentUIStatus.Progress
   
   UIStatusStack.Push Stat
   
   CurrentUIStatus.Text1 = StatusText
   StatusBar.Panels(1).Text = CurrentUIStatus.Text1
   CurrentUIStatus.Text2 = ""
   StatusBar.Panels(2).Text = CurrentUIStatus.Text2
   CurrentUIStatus.Progress = 0
   ProgressBar.Value = CurrentUIStatus.Progress
   If Busy And Not CurrentUIStatus.Busy Then
      On Error Resume Next
      CurrentUIStatus.Busy = True
      Set CurrentUIStatus.ActiveControlAtBusy = Screen.ActiveForm.ActiveControl
      CurrentUIStatus.ActiveMousePointerAtBusy = Screen.MousePointer
      Screen.MousePointer = Hourglass
      Screen.ActiveForm.Enabled = False
      DoEvents
   End If
End Sub
Private Sub UIStatusSetSub(SubText As String)

   CurrentUIStatus.Text2 = SubText
   StatusBar.Panels(2).Text = CurrentUIStatus.Text2
End Sub
Private Sub UIStatusProgress(Total As Long, Left As Long)

   If Total > 0 Then
      CurrentUIStatus.Progress = 100 - CInt(Left / Total * 100)
   Else
      CurrentUIStatus.Progress = 0
   End If
   ProgressBar.Value = CurrentUIStatus.Progress
End Sub
Private Sub UIStatusClear()

   Dim CurrBusy As Boolean

   CurrBusy = CurrentUIStatus.Busy

   UIStatusStack.Pop CurrentUIStatus
   StatusBar.Panels(1).Text = CurrentUIStatus.Text1
   StatusBar.Panels(2).Text = CurrentUIStatus.Text2
   ProgressBar.Value = CurrentUIStatus.Progress
   If CurrBusy And Not CurrentUIStatus.Busy Then
      On Error Resume Next
      Set Screen.ActiveForm.ActiveControl = CurrentUIStatus.ActiveControlAtBusy
      Screen.MousePointer = CurrentUIStatus.ActiveMousePointerAtBusy
      Screen.ActiveForm.Enabled = True
      DoEvents
   End If
End Sub

Private Sub WaitForUIBusy()

   Dim T As Double
   
   T = Timer + 10
   Do While UIBusy And T > Timer
      DoEvents
   Loop
End Sub
Private Function GetImportFileName() As String

   Dim Filter As String
   Dim Pos As Integer
   
   Filter = Client.Texts.Txt(1000801, "DSS diktat") & " (*.dss)|*.dss|"
   Filter = Filter & Client.Texts.Txt(1000802, "Alla filer") & " (*.*)|*.*"
   
   frmMain.CDialog.Filename = ""
   frmMain.CDialog.InitDir = GetDigtaDSSFolder()
   frmMain.CDialog.CancelError = True
   frmMain.CDialog.DefaultExt = "dss"
   frmMain.CDialog.DialogTitle = Client.Texts.Txt(1000800, "Importera diktat")
   frmMain.CDialog.Filter = Filter
   frmMain.CDialog.FilterIndex = 1
   frmMain.CDialog.Flags = cdlOFNExplorer Or cdlOFNFileMustExist
   frmMain.CDialog.HelpFile = ""
   frmMain.CDialog.HelpCommand = 0
   frmMain.CDialog.HelpContext = 0
   On Error Resume Next
   frmMain.CDialog.Action = 1
   If Err <> 0 Then
      Exit Function
   End If
   On Error GoTo 0

   GetImportFileName = frmMain.CDialog.Filename
End Function
Private Function CopyImportFileToTempStorage(Source As String, Dest As String) As Boolean

   On Error GoTo CopyImportFileToTempStorage_Err
   FileCopy Source, Dest
   CopyImportFileToTempStorage = True
   Exit Function
   
CopyImportFileToTempStorage_Err:
   CopyImportFileToTempStorage = False
   Exit Function
End Function
