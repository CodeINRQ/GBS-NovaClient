VERSION 5.00
Begin VB.UserControl ucEditGroup 
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   ScaleHeight     =   2490
   ScaleWidth      =   8265
   Begin VB.Frame fraGroups 
      Caption         =   "Grupper"
      Height          =   2415
      HelpContextID   =   1100000
      Left            =   0
      TabIndex        =   0
      Tag             =   "1100101"
      Top             =   0
      Width           =   8175
      Begin VB.CheckBox chkTextEditor 
         Caption         =   "&Text redigerare"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Tag             =   "1100112"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtDelayedHours 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   11
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox chkDelayed 
         Caption         =   "För&dröjd"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Tag             =   "1100110"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtOrgText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton cmdSaveRoles 
         Caption         =   "Spara &roller"
         Height          =   300
         Left            =   6000
         TabIndex        =   14
         Tag             =   "1100109"
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox chkSupervisor 
         Caption         =   "&Administratör"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Tag             =   "1100108"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox chkListener 
         Caption         =   "L&yssnare"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Tag             =   "1100107"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CheckBox chkTranscriber 
         Caption         =   "Utskrivare"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Tag             =   "1100106"
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CheckBox chkAuthor 
         Caption         =   "&Intalare"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         TabIndex        =   4
         Tag             =   "1100105"
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Ändra..."
         Height          =   300
         Left            =   6000
         TabIndex        =   13
         Tag             =   "1100103"
         Top             =   600
         Width           =   2055
      End
      Begin VB.ListBox lstGroup 
         Height          =   2010
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Lägg till..."
         Height          =   300
         Left            =   6000
         TabIndex        =   12
         Tag             =   "1100102"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblDelayedHoursTitle 
         Caption         =   "t&immar"
         Height          =   255
         Left            =   5880
         TabIndex        =   10
         Tag             =   "1100111"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblTitla 
         Caption         =   "Roller i &organisationsenheten:"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
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
   cmdNew.Enabled = ThereIsAValidCurrentOrg
   cmdChange.Enabled = CurrGroupId > 0 And ThereIsAValidCurrentOrg
   cmdSaveRoles.Enabled = CurrGroupId > 0 And ThereIsAValidCurrentOrg And RolesChanged And DelayedComplete
End Sub

Private Sub chkAuthor_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkTextEditor_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkDelayed_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkListener_Click()

   RolesChanged = True
   SetEnabled
End Sub

Private Sub chkSupervisor_Click()

   RolesChanged = True
   SetEnabled
End Sub


Private Sub chkTranscriber_Click()

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

   Dim R As New clsRoles
   
   cmdSaveRoles.Enabled = False
   If chkAuthor.Enabled Then
      R.Author = chkAuthor.Value = vbChecked
   End If
   If chkTextEditor.Enabled Then
      R.TextEditor = chkTextEditor.Value = vbChecked
   End If
   If chkTranscriber.Enabled Then
      R.Transcriber = chkTranscriber.Value = vbChecked
   End If
   If chkDelayed.Enabled Then
      R.Delayed = chkDelayed.Value = vbChecked
      On Error Resume Next
      R.DelayedHours = CInt(txtDelayedHours.Text)
   End If
   If chkListener.Enabled Then
      R.Listen = chkListener.Value = vbChecked
   End If
   If chkSupervisor.Enabled Then
      R.Supervise = chkSupervisor.Value = vbChecked
   End If
   R.GroupId = CurrGroupId
   R.OrgId = CurrOrgId
   Client.RolesMgr.SaveRoles R
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
   
   If CurrOrgId < 30000 Then
      If CurrOrgId <> LastOrgId Or CurrGroupId <> LastGroupId Then
         LastOrgId = CurrOrgId
         LastGroupId = CurrGroupId
         
         Dim Org As clsOrg
         Dim Roles As clsRoles
         
         chkAuthor.Value = vbUnchecked: chkAuthor.Enabled = False
         chkTextEditor.Value = vbUnchecked: chkTextEditor.Enabled = False
         chkTranscriber.Value = vbUnchecked: chkTranscriber.Enabled = False
         chkListener.Value = vbUnchecked: chkListener.Enabled = False
         chkSupervisor.Value = vbUnchecked: chkSupervisor.Enabled = False
         txtOrgText.Text = ""
         
         If CurrGroupId > 0 Then
            Set Org = Nothing
            Client.OrgMgr.GetOrgFromId Org, CurrOrgId
            If Not Org Is Nothing Then
               txtOrgText.Text = Org.OrgText
               
               IsCurrentGroupSysAdmin = LCase(CurrGroup.GroupText) = "sysadmin"
               
               chkAuthor.Value = vbUnchecked: chkAuthor.Enabled = Not IsCurrentGroupSysAdmin
               chkTextEditor.Value = vbUnchecked: chkTextEditor.Enabled = Not IsCurrentGroupSysAdmin
               chkTranscriber.Value = vbUnchecked: chkTranscriber.Enabled = Not IsCurrentGroupSysAdmin
               chkDelayed.Value = vbUnchecked: chkDelayed.Enabled = Not IsCurrentGroupSysAdmin
               txtDelayedHours.Text = "": txtDelayedHours.Enabled = Not IsCurrentGroupSysAdmin
               chkListener.Value = vbUnchecked: chkListener.Enabled = Not IsCurrentGroupSysAdmin
               chkSupervisor.Value = vbUnchecked: chkSupervisor.Enabled = Not IsCurrentGroupSysAdmin
               
               Do
                  If Client.RolesMgr.GetRoles(Roles, CurrGroupId, Org.OrgId) Then
                     If Roles.Author Then
                        chkAuthor.Value = vbChecked
                        If Org.OrgId <> CurrOrgId Then
                           chkAuthor.Enabled = False
                        End If
                     End If
                     If Roles.TextEditor Then
                        chkTextEditor.Value = vbChecked
                        If Org.OrgId <> CurrOrgId Then
                           chkTextEditor.Enabled = False
                        End If
                     End If
                     If Roles.Transcriber Then
                        chkTranscriber.Value = vbChecked
                        If Roles.Delayed Then
                           chkDelayed.Value = vbChecked
                           txtDelayedHours.Text = CStr(Roles.DelayedHours)
                        Else
                           txtDelayedHours.Enabled = False
                        End If
                        If Org.OrgId <> CurrOrgId Then
                           chkTranscriber.Enabled = False
                           chkDelayed.Enabled = False
                           txtDelayedHours.Enabled = False
                        End If
                     End If
                     If Roles.Listen Then
                        chkListener.Value = vbChecked
                        If Org.OrgId <> CurrOrgId Then
                           chkListener.Enabled = False
                        End If
                     End If
                     If Roles.Supervise Then
                        chkSupervisor.Value = vbChecked
                        If Org.OrgId <> CurrOrgId Then
                           chkSupervisor.Enabled = False
                        End If
                     End If
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

Private Sub txtDelayedHours_Change()

   RolesChanged = True
   SetEnabled
End Sub
