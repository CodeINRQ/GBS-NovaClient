VERSION 5.00
Begin VB.Form frmEditGroup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Grupp"
   ClientHeight    =   1935
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   6030
   HelpContextID   =   1040000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1040100"
   Begin VB.TextBox txtAdmOrgText 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   6
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtGroupDesc 
      Height          =   285
      Left            =   120
      MaxLength       =   255
      TabIndex        =   3
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox txtGroupText 
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Avbryt"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Tag             =   "1040104"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Spara"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Tag             =   "1040103"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Administreras av:"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Tag             =   "1040105"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "&Beskrivning:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Tag             =   "1040102"
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "&Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "1040101"
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmEditGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public GroupToEdit As clsGroup
Public AdmOrg As clsOrg
Public Event SaveClicked()

Private Dirty As Boolean

Private Sub cmdCancel_Click()

   Unload Me
End Sub

Private Sub cmdSave_Click()

   GroupToEdit.GroupText = txtGroupText.Text
   GroupToEdit.GroupDesc = txtGroupDesc.Text
   GroupToEdit.AdmOrgId = AdmOrg.OrgId
   Client.GroupMgr.SaveGroup GroupToEdit
   
   RaiseEvent SaveClicked
   Unload Me
End Sub

Private Sub Form_Activate()

   txtGroupText.Text = GroupToEdit.GroupText
   txtGroupDesc.Text = GroupToEdit.GroupDesc
   txtAdmOrgText.Text = AdmOrg.OrgText
End Sub

Private Sub SetEnabled()

   cmdSave.Enabled = Dirty And Len(txtGroupText.Text) > 0
End Sub

Private Sub Form_Load()

   CenterAndTranslateForm Me, frmMain

   SetEnabled
End Sub

Private Sub txtGroupDesc_Change()

   Dirty = True
   SetEnabled
End Sub

Private Sub txtGroupText_Change()

   Dirty = True
   SetEnabled
End Sub
