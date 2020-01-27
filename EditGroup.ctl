VERSION 5.00
Begin VB.UserControl ucEditGroup 
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   ScaleHeight     =   4065
   ScaleWidth      =   8265
   Begin VB.Frame fraGroups 
      Caption         =   "Grupper"
      Height          =   3975
      HelpContextID   =   1100000
      Left            =   0
      TabIndex        =   25
      Tag             =   "1100101"
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdSaveRoles 
         Caption         =   "Spara &rättigheter"
         Height          =   300
         Left            =   6000
         TabIndex        =   24
         Tag             =   "1100109"
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Ändra..."
         Height          =   300
         Left            =   6000
         TabIndex        =   23
         Tag             =   "1100103"
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Lägg till..."
         Height          =   300
         Left            =   6000
         TabIndex        =   22
         Tag             =   "1100102"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkUnlockingInherit 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Tag             =   "1100107"
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkStatisticsInherit 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Tag             =   "1100107"
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox chkUserAdminInherit 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Tag             =   "1100107"
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkTranscriberInherit 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Tag             =   "1100107"
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox chkAuthorInherit 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Tag             =   "1100107"
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkListenerInherit 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Tag             =   "1100107"
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkAuditingInherit 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   26
         Tag             =   "1100107"
         Top             =   3600
         Width           =   255
      End
      Begin VB.CheckBox chkHistoryInherit 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Tag             =   "1100107"
         Top             =   3120
         Width           =   255
      End
      Begin VB.CheckBox chkTextEditorInherit 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Tag             =   "1100107"
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkAuditing 
         Caption         =   "&Diktatspårning"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Tag             =   "1100116"
         Top             =   3600
         Width           =   2775
      End
      Begin VB.CheckBox chkUnlocking 
         Caption         =   "U&pplåsning av diktat"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Tag             =   "1100115"
         Top             =   3360
         Width           =   2775
      End
      Begin VB.CheckBox chkHistory 
         Caption         =   "&Historik"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Tag             =   "1100114"
         Top             =   3120
         Width           =   2775
      End
      Begin VB.CheckBox chkStatistics 
         Caption         =   "&Statistik"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Tag             =   "1100113"
         Top             =   2880
         Width           =   2775
      End
      Begin VB.CheckBox chkTextEditor 
         Caption         =   "&Text redigering"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Tag             =   "1100112"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtDelayedHours 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   13
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox chkDelayed 
         Caption         =   "För&dröjd"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Tag             =   "1100110"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtOrgText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox chkUserAdmin 
         Caption         =   "&Användaradministration"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Tag             =   "1100108"
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CheckBox chkListener 
         Caption         =   "L&yssning"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Tag             =   "1100107"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox chkTranscriber 
         Caption         =   "Utskrift"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Tag             =   "1100106"
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CheckBox chkAuthor 
         Caption         =   "&Intalning nya diktat"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Tag             =   "1100105"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ListBox lstGroup 
         Height          =   2985
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Tillåt"
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Tag             =   "1100118"
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Arv"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Tag             =   "1100117"
         Top             =   960
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   3120
         X2              =   6360
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lblDelayedHoursTitle 
         Caption         =   "t&immar"
         Height          =   255
         Left            =   6480
         TabIndex        =   12
         Tag             =   "1100111"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblTitla 
         Caption         =   "Rättigheter i &organisationsenheten:"
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Tag             =   "1100104"
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "ucEditGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event GroupsChanged()

Private WithEvents frmEdit As frmEditGroup
Attribute frmEdit.VB_VarHelpID = -1
Private CurrGroup As clsGroup
Private CurrOrg As clsOrg
Private CurrOrgId As Long
Private CurrGroupId As Long
Private RolesChanged As Boolean

Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub

Public Sub Init()

   Dim I As Integer
   
   RolesChanged = True
   CurrGroupId = 0
   lstGroup.Clear
   For I = 0 To Client.GroupMgr.Count - 1
      lstGroup.AddItem Client.GroupMgr.TextFromIndex(I)
   Next I
   SetEnabled
End Sub
Private Sub SetEnabled()

   Dim ThereIsAValidCurrentOrg As Boolean
   Dim DelayedComplete As Boolean
      
   ThereIsAValidCurrentOrg = CurrOrgId > 0 And CurrOrgId < 30000
   If chkTranscriberInherit.Value = vbUnchecked Then
      If chkTranscriber.Value = vbChecked Then
         chkDelayed.Enabled = chkTranscriber.Enabled
      Else
         chkDelayed.Enabled = False
      End If
      If chkDelayed.Value = vbChecked Then
         txtDelayedHours.Enabled = chkDelayed.Enabled
         DelayedComplete = Len(txtDelayedHours) > 0
      Else
         txtDelayedHours.Enabled = False
         DelayedComplete = True
      End If
   Else
      chkDelayed.Enabled = False
      txtDelayedHours.Enabled = False
      DelayedComplete = True
   End If
   
   chkListener.Enabled = chkListenerInherit.Value = vbUnchecked
   chkAuthor.Enabled = chkAuthorInherit.Value = vbUnchecked
   chkTextEditor.Enabled = chkTextEditorInherit.Value = vbUnchecked
   chkTranscriber.Enabled = chkTranscriberInherit.Value = vbUnchecked
   chkUserAdmin.Enabled = chkUserAdminInherit.Value = vbUnchecked
   chkStatistics.Enabled = chkStatisticsInherit.Value = vbUnchecked
   chkHistory.Enabled = chkHistoryInherit.Value = vbUnchecked
   chkUnlocking.Enabled = chkUnlockingInherit.Value = vbUnchecked
   chkAuditing.Enabled = chkAuditingInherit.Value = vbUnchecked

   cmdNew.Enabled = ThereIsAValidCurrentOrg
   cmdChange.Enabled = CurrGroupId > 0 And ThereIsAValidCurrentOrg
   cmdSaveRoles.Enabled = CurrGroupId > 0 And ThereIsAValidCurrentOrg And RolesChanged And DelayedComplete
End Sub


Private Sub chkAuditingInherit_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkAuthorInherit_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkHistoryInherit_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkListener_Click()

   RolesChanged = True
   SetEnabled
End Sub
Private Sub chkAuthor_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkListenerInherit_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkStatisticsInherit_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkTextEditor_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkTextEditorInherit_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkTranscriber_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkDelayed_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkStatistics_Click()

   RolesChanged = True
   SetEnabled
End Sub
Private Sub chkHistory_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkTranscriberInherit_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkUnlocking_Click()

   RolesChanged = True
   SetEnabled
End Sub
Private Sub chkAuditing_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkUnlockingInherit_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkUserAdmin_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkUserAdminInherit_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub cmdChange_Click()

   EditCurrGroup
End Sub

Private Sub cmdNew_Click()

  Set CurrGroup = New clsGroup
  EditCurrGroup
End Sub

Private Sub EditCurrGroup()

  Set frmEdit = New frmEditGroup
  Set frmEdit.GroupToEdit = CurrGroup
  Set frmEdit.AdmOrg = CurrOrg
  frmEdit.Show vbModal
End Sub

Private Sub cmdSaveRoles_Click()

   Dim r As New clsRoles
   
   cmdSaveRoles.Enabled = False
   If chkListenerInherit.Value = vbChecked Then
      r.ListenInherit = True
   Else
      r.Listen = chkListener.Value = vbChecked
   End If
   
   If chkAuthorInherit.Value = vbChecked Then
      r.AuthorInherit = True
   Else
      r.Author = chkAuthor.Value = vbChecked
   End If
   
   If chkTextEditorInherit.Value = vbChecked Then
      r.TextEditorInherit = True
   Else
      r.TextEditor = chkTextEditor.Value = vbChecked
   End If
   
   If chkTranscriberInherit.Value = vbChecked Then
      r.TranscriberInherit = True
   Else
      r.Transcriber = chkTranscriber.Value = vbChecked
      If r.Transcriber Then
         r.Delayed = chkDelayed.Value = vbChecked
         On Error Resume Next
         r.DelayedHours = CInt(txtDelayedHours.Text)
         On Error GoTo 0
      Else
         r.Delayed = False
         r.DelayedHours = 0
      End If
   End If
   
   If chkUserAdminInherit.Value = vbChecked Then
      r.UserAdminInherit = True
   Else
      r.UserAdmin = chkUserAdmin.Value = vbChecked
   End If
   
   If chkStatisticsInherit.Value = vbChecked Then
      r.StatisticsInherit = True
   Else
      r.Statistics = chkStatistics.Value = vbChecked
   End If
   
   If chkHistoryInherit.Value = vbChecked Then
      r.HistoryInherit = True
   Else
      r.History = chkHistory.Value = vbChecked
   End If
   
   If chkUnlockingInherit.Value = vbChecked Then
      r.UnlockingInherit = True
   Else
      r.Unlocking = chkUnlocking.Value = vbChecked
   End If
   
   If chkAuditingInherit.Value = vbChecked Then
      r.AuditingInherit = True
   Else
      r.Auditing = chkAuditing.Value = vbChecked
   End If
   
   r.GroupId = CurrGroupId
   r.OrgId = CurrOrgId
   
   If CurrOrg.OrgParent = 0 Then
      r.CleanRolesBeforeSaveRoot
   End If
   
   Client.RolesMgr.SaveRoles r
   RolesChanged = False
   SetEnabled
End Sub

Private Sub frmEdit_SaveClicked()

   Init
   RaiseEvent GroupsChanged
   SetEnabled
End Sub

Private Sub lstGroup_Click()

   CurrGroupId = Client.GroupMgr.IdFromIndex(lstGroup.ListIndex)
   Client.GroupMgr.GetGroupFromId CurrGroup, CurrGroupId
   ShowRolesForOrgIdAndGroupId
   SetEnabled
End Sub
Public Sub NewOrg(OrgId As Long)

   CurrOrgId = OrgId
   Client.OrgMgr.GetOrgFromId CurrOrg, CurrOrgId
   ShowRolesForOrgIdAndGroupId
   SetEnabled
End Sub
Private Sub ShowRolesForOrgIdAndGroupId()

   Static LastOrgId As Long
   Static LastGroupId As Long
   Dim OId As Long
   Dim IsCurrentGroupSysAdmin As Boolean
   Dim ListenFixed As Boolean
   Dim AuthorFixed As Boolean
   Dim TextEditorFixed As Boolean
   Dim TranscriberFixed As Boolean
   Dim UserAdminFixed As Boolean
   Dim StatisticsFixed As Boolean
   Dim HistoryFixed As Boolean
   Dim UnlockingFixed As Boolean
   Dim AuditingFixed As Boolean
   Dim ThisOrg As Boolean
   Dim ThisIsRoot As Boolean
   
   If CurrOrgId < 30000 Then
      If CurrOrgId <> LastOrgId Or CurrGroupId <> LastGroupId Then
         LastOrgId = CurrOrgId
         LastGroupId = CurrGroupId
         
         Dim Org As clsOrg
         Dim Roles As clsRoles
         
         SetAllCheckboxesToInitValue False, True
         
         txtOrgText.Text = ""
         
         If CurrGroupId > 0 Then
            Set Org = Nothing
            Client.OrgMgr.GetOrgFromId Org, CurrOrgId
            If Not Org Is Nothing Then
               txtOrgText.Text = Org.OrgText
               
               IsCurrentGroupSysAdmin = LCase(CurrGroup.GroupText) = "sysadmin"
                              
               ThisOrg = Org.OrgId = CurrOrgId
               ThisIsRoot = Org.OrgParent = 0
               SetAllCheckboxesToInitValue Not IsCurrentGroupSysAdmin, ThisIsRoot
                              
               Do

                  If Client.RolesMgr.GetRoles(Roles, CurrGroupId, Org.OrgId) Then
                     
                     ThisOrg = Org.OrgId = CurrOrgId
                     ThisIsRoot = Org.OrgParent = 0
                     SetCheckboxesForOneRight Roles.Listen, Roles.ListenInherit, chkListener, chkListenerInherit, ThisOrg, ThisIsRoot, ListenFixed
                     SetCheckboxesForOneRight Roles.Author, Roles.AuthorInherit, chkAuthor, chkAuthorInherit, ThisOrg, ThisIsRoot, AuthorFixed
                     SetCheckboxesForOneRight Roles.TextEditor, Roles.TextEditorInherit, chkTextEditor, chkTextEditorInherit, ThisOrg, ThisIsRoot, TextEditorFixed
                     
                     If Roles.Transcriber And Not TranscriberFixed Then
                        If Roles.Delayed Then
                           chkDelayed.Value = vbChecked
                           txtDelayedHours.Text = CStr(Roles.DelayedHours)
                        Else
                           txtDelayedHours.Enabled = False
                        End If
                     End If
                     
                     SetCheckboxesForOneRight Roles.Transcriber, Roles.TranscriberInherit, chkTranscriber, chkTranscriberInherit, ThisOrg, ThisIsRoot, TranscriberFixed
                     
                     SetCheckboxesForOneRight Roles.UserAdmin, Roles.UserAdminInherit, chkUserAdmin, chkUserAdminInherit, ThisOrg, ThisIsRoot, UserAdminFixed
                     SetCheckboxesForOneRight Roles.Statistics, Roles.StatisticsInherit, chkStatistics, chkStatisticsInherit, ThisOrg, ThisIsRoot, StatisticsFixed
                     SetCheckboxesForOneRight Roles.History, Roles.HistoryInherit, chkHistory, chkHistoryInherit, ThisOrg, ThisIsRoot, HistoryFixed
                     SetCheckboxesForOneRight Roles.Unlocking, Roles.UnlockingInherit, chkUnlocking, chkUnlockingInherit, ThisOrg, ThisIsRoot, UnlockingFixed
                     SetCheckboxesForOneRight Roles.Auditing, Roles.AuditingInherit, chkAuditing, chkAuditingInherit, ThisOrg, ThisIsRoot, AuditingFixed
                     
                     SetEnabled
                  End If
                  If Org.OrgParent > 0 Then
                     OId = Org.OrgParent
                     Set Org = Nothing
                     Client.OrgMgr.GetOrgFromId Org, OId
                     If Org Is Nothing Then Exit Do
                  Else
                     Exit Do
                  End If
               Loop
               RolesChanged = False
            End If
         End If
      End If
   End If
   SetEnabled
End Sub
Private Sub SetCheckboxesForOneRight(OneRole As Boolean, OneRoleInherit As Boolean, chkRole As CheckBox, chkRoleInherit As CheckBox, ThisOrg As Boolean, ThisIsRoot As Boolean, ByRef ThisRoleFixed As Boolean)

   If ThisIsRoot And ThisOrg Then
      chkRoleInherit.Value = vbUnchecked
      chkRoleInherit.Enabled = False
      chkRole.Enabled = True
      If OneRole Then
         chkRole.Value = vbChecked
      Else
         chkRole.Value = vbUnchecked
      End If
      ThisRoleFixed = True
   Else
      If Not ThisRoleFixed Then
         If Not OneRoleInherit Then
            If ThisOrg Then
               chkRoleInherit.Value = vbUnchecked
            End If
            If OneRole Then
               chkRole.Value = vbChecked
            Else
               chkRole.Value = vbUnchecked
            End If
            ThisRoleFixed = True
         Else
            chkRoleInherit.Value = vbChecked
         End If
      End If
   End If
End Sub

Private Sub SetAllCheckboxesToInitValue(Enbld As Boolean, ThisIsRoot As Boolean)

   Dim InheritValue As Integer
   Dim InheritEnabled As Boolean
   
   If ThisIsRoot Then
      InheritValue = vbUnchecked
      InheritEnabled = False
   Else
      InheritValue = vbChecked
      InheritEnabled = Enbld
   End If

   chkListener.Value = vbUnchecked: chkListener.Enabled = Enbld
   chkListenerInherit.Value = InheritValue: chkListenerInherit.Enabled = InheritEnabled
   chkAuthor.Value = vbUnchecked: chkAuthor.Enabled = Enbld
   chkAuthorInherit.Value = InheritValue: chkAuthorInherit.Enabled = InheritEnabled
   chkTextEditor.Value = vbUnchecked: chkTextEditor.Enabled = Enbld
   chkTextEditorInherit.Value = InheritValue: chkTextEditorInherit.Enabled = InheritEnabled
   chkTranscriber.Value = vbUnchecked: chkTranscriber.Enabled = Enbld
   chkTranscriberInherit.Value = InheritValue: chkTranscriberInherit.Enabled = InheritEnabled
   chkDelayed.Value = vbUnchecked: chkDelayed.Enabled = Enbld
   txtDelayedHours.Text = "": txtDelayedHours.Enabled = Enbld
   chkUserAdmin.Value = vbUnchecked: chkUserAdmin.Enabled = Enbld
   chkUserAdminInherit.Value = InheritValue: chkUserAdminInherit.Enabled = InheritEnabled
   chkStatistics.Value = vbUnchecked: chkStatistics.Enabled = Enbld
   chkStatisticsInherit.Value = InheritValue: chkStatisticsInherit.Enabled = InheritEnabled
   chkHistory.Value = vbUnchecked: chkHistory.Enabled = Enbld
   chkHistoryInherit.Value = InheritValue: chkHistoryInherit.Enabled = InheritEnabled
   chkUnlocking.Value = vbUnchecked: chkUnlocking.Enabled = Enbld
   chkUnlockingInherit.Value = InheritValue: chkUnlockingInherit.Enabled = InheritEnabled
   chkAuditing.Value = vbUnchecked: chkAuditing.Enabled = Enbld
   chkAuditingInherit.Value = InheritValue: chkAuditingInherit.Enabled = InheritEnabled
End Sub

Private Sub txtDelayedHours_Change()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub txtDelayedHours_GotFocus()

   SelectAllText ActiveControl
End Sub
