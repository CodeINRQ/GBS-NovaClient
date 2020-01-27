VERSION 5.00
Begin VB.UserControl ucEditOrg 
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   ScaleHeight     =   1770
   ScaleWidth      =   8220
   Begin VB.Frame fraEditOrg 
      Caption         =   "&Organisation"
      Height          =   1695
      HelpContextID   =   1110000
      Left            =   0
      TabIndex        =   8
      Tag             =   "1110101"
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtParentText 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   9
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Ny under"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5520
         TabIndex        =   7
         Tag             =   "1110109"
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Återställ"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5520
         TabIndex        =   6
         Tag             =   "1110108"
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Spara"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5520
         TabIndex        =   5
         Tag             =   "1110107"
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chkShowBelow 
         Caption         =   "Visa diktat för &underliggande"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Tag             =   "1110106"
         Top             =   1320
         Width           =   5175
      End
      Begin VB.CheckBox chkShowInTree 
         Caption         =   "&Visa i trädstruktur"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Tag             =   "1110105"
         Top             =   1080
         Width           =   5175
      End
      Begin VB.CheckBox chkDictContainer 
         Caption         =   "&Kan lagra diktat"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Tag             =   "1110104"
         Top             =   840
         Width           =   5175
      End
      Begin VB.TextBox txtOrgText 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Plats:"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Tag             =   "1110103"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "&Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Tag             =   "1110102"
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "ucEditOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event OrgSaved(Org As clsOrg)

Dim CurrOrgId As Long
Dim mOrg As clsOrg
Dim mOrgParent As clsOrg
Dim Dirty As Boolean
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub

Public Sub OrgSelected(OrgId As Long)

   If CurrOrgId <> OrgId Then
      CurrOrgId = OrgId
      Set mOrg = Nothing
      Set mOrgParent = Nothing
      
      Client.OrgMgr.GetOrgFromId mOrg, OrgId
      If Not mOrg Is Nothing Then
         If mOrg.OrgParent <> 0 Then
            Client.OrgMgr.GetOrgFromId mOrgParent, mOrg.OrgParent
         Else
            Set mOrgParent = New clsOrg
         End If
      Else
         Set mOrg = New clsOrg
         Set mOrgParent = New clsOrg
      End If
      ShowOrg
   End If
End Sub
Private Sub ShowOrg()

   txtOrgText.Text = mOrg.OrgText
   txtParentText.Text = mOrgParent.OrgText
   If mOrg.DictContainer Then
      chkDictContainer.Value = vbChecked
   Else
      chkDictContainer.Value = vbUnchecked
   End If
   If mOrg.ShowInTree Then
      chkShowInTree.Value = vbChecked
   Else
      chkShowInTree.Value = vbUnchecked
   End If
   If mOrg.ShowBelow Then
      chkShowBelow.Value = vbChecked
   Else
      chkShowBelow.Value = vbUnchecked
   End If
   Dirty = False
   SetEnabled
End Sub
Private Sub SetEnabled()

   cmdSave.Enabled = Dirty And mOrg.OrgParent <> 0
   cmdNew.Enabled = Not Dirty And mOrg.OrgId <> 0
   cmdReset.Enabled = Dirty
End Sub

Private Sub chkDictContainer_Click()

   Dirty = True
   SetEnabled
End Sub

Private Sub chkShowBelow_Click()

   Dirty = True
   SetEnabled
End Sub

Private Sub chkShowInTree_Click()

   Dirty = True
   SetEnabled
End Sub

Private Sub cmdNew_Click()

   Set mOrgParent = mOrg
   Set mOrg = New clsOrg
   mOrg.OrgParent = mOrgParent.OrgId
   ShowOrg
End Sub

Private Sub cmdReset_Click()

   ShowOrg
End Sub

Private Sub cmdSave_Click()

   mOrg.OrgText = txtOrgText.Text
   mOrg.DictContainer = chkDictContainer.Value = vbChecked
   mOrg.ShowInTree = chkShowInTree.Value = vbChecked
   mOrg.ShowBelow = chkShowBelow.Value = vbChecked

   Client.OrgMgr.SaveOrg mOrg
   Dirty = False
   SetEnabled
   RaiseEvent OrgSaved(mOrg)
End Sub


Private Sub txtOrgText_Change()

   Dirty = True
   SetEnabled
End Sub
