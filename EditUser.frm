VERSION 5.00
Begin VB.Form frmEditUser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Användare"
   ClientHeight    =   5565
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   7350
   HelpContextID   =   1050000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1050100"
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Inaktivera"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Tag             =   "1050110"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ListBox lstUserGroup 
      Height          =   2085
      Left            =   3360
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   2760
      Width           =   3855
   End
   Begin CareTalk.ucOrgTree ucOrgTree 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   7435
   End
   Begin VB.TextBox txtLongName 
      Height          =   285
      Left            =   3360
      MaxLength       =   255
      TabIndex        =   11
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4920
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtShortName 
      Height          =   285
      Left            =   3360
      MaxLength       =   255
      TabIndex        =   9
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtLoginName 
      Height          =   285
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Tag             =   "1050109"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Spara"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Tag             =   "1050108"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgDice 
      Height          =   300
      Left            =   6480
      Picture         =   "EditUser.frx":0000
      Top             =   960
      Width           =   390
   End
   Begin VB.Label Label7 
      Caption         =   "&Tillhör grupper:"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Tag             =   "1050107"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "&Standard organisation:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "1050101"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Långt n&amn:"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Tag             =   "1050106"
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "&Bekräfta lösenord:"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Tag             =   "1050104"
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "L&ösenord:"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Tag             =   "1050103"
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "&Kort namn:"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Tag             =   "1050105"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "&Login namn:"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Tag             =   "1050102"
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public UserToEdit As clsUser
Public Event SaveClicked()
Public Event DeleteClicked()
Public CurrHomeOrgId As Long

Private Dirty As Boolean
Private FirstPossibleHomeOrg As Long

Private Sub cmdCancel_Click()

   Unload Me
End Sub

Private Sub cmdDelete_Click()

   If MsgBox(Client.Texts.Txt(1050111, "Användare kommer att tas bort! Är du säker?"), vbYesNo) = vbYes Then
      Client.UserMgr.DeleteUser UserToEdit

      RaiseEvent DeleteClicked
      Unload Me
   End If
End Sub

Private Sub cmdSave_Click()

   Dim I As Integer
   Dim Grp As clsGroup
   Dim OkToSave As Boolean
   Dim UsrTmp As clsUser

   UserToEdit.LoginName = txtLoginName.Text
   UserToEdit.Password = txtPassword.Text
   UserToEdit.Password = txtConfirmPassword.Text
   UserToEdit.ShortName = txtShortName.Text
   UserToEdit.LongName = txtLongName.Text
   UserToEdit.HomeOrgId = CurrHomeOrgId

   OkToSave = True 'default
   If UserToEdit.UserId = 0 Then   'new user
      Client.UserMgr.GetUserFromLoginName UsrTmp, UserToEdit.LoginName
      If Not UsrTmp Is Nothing Then
         MsgBox Client.Texts.Txt(1050112, "Användaren finns redan!"), vbOKOnly
         OkToSave = False
      End If
   End If
   
   If OkToSave Then
      Client.UserMgr.SaveUser UserToEdit
   
      If UserToEdit.UserId > 0 Then
         For I = 0 To lstUserGroup.ListCount - 1
            If lstUserGroup.Selected(I) Then
               Client.GroupMgr.SaveUserGroup UserToEdit.UserId, lstUserGroup.ItemData(I)
            Else
               Client.GroupMgr.DeleteOneUserGroup UserToEdit.UserId, lstUserGroup.ItemData(I)
            End If
         Next I
      End If
      
      RaiseEvent SaveClicked
      Unload Me
   End If
End Sub

Private Sub Form_Activate()

   Dim Org As clsOrg

   txtLoginName.Text = UserToEdit.LoginName
   txtPassword.Text = UserToEdit.Password
   txtConfirmPassword.Text = UserToEdit.Password
   txtShortName.Text = UserToEdit.ShortName
   txtLongName.Text = UserToEdit.LongName
   Client.OrgMgr.GetOrgFromId Org, UserToEdit.HomeOrgId
   If Not Org Is Nothing Then
      If Org.DictContainer Then
         CurrHomeOrgId = Org.OrgId
      Else
         CurrHomeOrgId = FirstPossibleHomeOrg
      End If
   Else
      CurrHomeOrgId = 0
   End If
   If CurrHomeOrgId > 0 Then
      ucOrgTree.PickOrgId CurrHomeOrgId
   Else
      ucOrgTree.PickOrgId FirstPossibleHomeOrg
   End If
   
   Dirty = False
   SetEnabled
End Sub

Private Sub SetEnabled()

   Dim Enbl As Boolean
   
   Enbl = Dirty And Len(txtLoginName.Text) > 0 And Len(txtPassword.Text) >= Client.SysSettings.LoginPasswordMinLength
   Enbl = Enbl And txtPassword.Text = txtConfirmPassword.Text And Len(txtShortName.Text) > 0 And Len(txtLongName) > 0
   Enbl = Enbl And CurrHomeOrgId > 0
   Enbl = Enbl And LCase$(UserToEdit.ShortName) <> "sa"
   'Enbl = Enbl And lstUserGroup.SelCount > 0
   cmdSave.Enabled = Enbl
   cmdDelete.Enabled = UserToEdit.UserId > 0 And UserToEdit.InactivatedTime = 0
End Sub

Private Sub Form_Load()

   Dim Org As New clsOrg
   Dim I As Integer
   Dim GrpLstIdx As Integer
   Dim RightToSetAsHomeOrg As Boolean
   Dim Grp As clsGroup
   
   CenterAndTranslateForm Me, frmMain
   
   Client.OrgMgr.Init False
   
   For I = 0 To Client.OrgMgr.Count - 1
      Client.OrgMgr.GetSortedOrg Org, I
      If Org.ShowInTree Then
         RightToSetAsHomeOrg = Client.OrgMgr.CheckUserRole(Org.OrgId, RTUserAdmin) And Org.DictContainer
         If RightToSetAsHomeOrg Then
            If FirstPossibleHomeOrg = 0 Or Org.OrgId = Client.User.HomeOrgId Then
               FirstPossibleHomeOrg = Org.OrgId
            End If
            ucOrgTree.AddNode Org.OrgId, Org.ShowParent, Org.OrgText, 1, True
         Else
            ucOrgTree.AddNode Org.OrgId, Org.ShowParent, Org.OrgText, 5, False
         End If
      End If
   Next I
   
   Client.GroupMgr.Init
   lstUserGroup.Clear
   GrpLstIdx = 0
   For I = 0 To Client.GroupMgr.Count - 1
      Client.GroupMgr.GetGroupFromIndex Grp, I
      If Client.OrgMgr.CheckUserRole(Grp.AdmOrgId, RTUserAdmin) Then
         lstUserGroup.AddItem Grp.GroupText, GrpLstIdx
         lstUserGroup.ItemData(GrpLstIdx) = Grp.GroupId
         GrpLstIdx = GrpLstIdx + 1
      End If
   Next I
   MarkGroups
End Sub
Private Sub MarkGroups()

   Dim GroupId As Long
   Dim Idx As Integer
   
   Client.Server.CreateGroupListForUser UserToEdit.UserId
   Do While Client.Server.GroupListForUserGetNext(GroupId)
      If GroupId > 0 Then
         For Idx = 0 To lstUserGroup.ListCount - 1
            If GroupId = lstUserGroup.ItemData(Idx) Then
               lstUserGroup.Selected(Idx) = True
            End If
         Next Idx
      End If
   Loop
End Sub

Private Sub imgDice_DblClick()

   GeneratePassword
End Sub

Private Sub GeneratePassword()

   Dim NewPassword As String
   Dim PasswordLen As Integer
   Dim C As String
   
   Randomize
   NewPassword = ""
   PasswordLen = RndNumber(10, 15)
   Do While Len(NewPassword) < PasswordLen
      C = Chr$(RndNumber(48, 122))
      If C <> "'" And C <> """" Then
         NewPassword = NewPassword & C
      End If
   Loop
   txtPassword.Text = NewPassword
   txtConfirmPassword.Text = NewPassword
End Sub

Private Function RndNumber(Min As Integer, Max As Integer) As Integer

   RndNumber = Int(Rnd * (Max - Min + 1)) + Min
End Function

Private Sub Label1_Click()

   Clipboard.Clear
   Clipboard.SetText txtLoginName.Text
End Sub

Private Sub Label2_Click()

   Clipboard.Clear
   Clipboard.SetText txtShortName.Text
End Sub

Private Sub Label5_Click()

   Clipboard.Clear
   Clipboard.SetText txtLongName.Text
End Sub

Private Sub lstUserGroup_Click()

   Dim Idx As Integer
   
   
   Dirty = True
   SetEnabled
End Sub

Private Sub txtConfirmPassword_Change()

   Dirty = True
   SetEnabled
End Sub

Private Sub txtConfirmPassword_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtLoginName_Change()

   Dirty = True
   SetEnabled
End Sub

Private Sub txtLoginName_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtLongName_Change()

   Dirty = True
   SetEnabled
End Sub

Private Sub txtLongName_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtPassword_Change()

   Dirty = True
   SetEnabled
End Sub

Private Sub txtPassword_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtShortName_Change()

   Dirty = True
   SetEnabled
End Sub

Private Sub txtShortName_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub ucOrgTree_NewSelect(OrgId As Long, Txt As String)

   CurrHomeOrgId = OrgId
   Dirty = True
   SetEnabled
End Sub
