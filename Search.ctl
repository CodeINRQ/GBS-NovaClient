VERSION 5.00
Begin VB.UserControl ucSearch 
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ScaleHeight     =   4350
   ScaleWidth      =   4815
   Begin VB.TextBox txtTxt 
      Height          =   285
      HelpContextID   =   1150000
      Left            =   120
      MaxLength       =   50
      TabIndex        =   14
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox txtTranscriber 
      Height          =   285
      HelpContextID   =   1150000
      Left            =   120
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      HelpContextID   =   1150000
      Left            =   120
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Återställ"
      Height          =   310
      HelpContextID   =   1150000
      Left            =   3360
      TabIndex        =   13
      Tag             =   "1150108"
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Sök"
      Height          =   310
      HelpContextID   =   1150000
      Left            =   3360
      TabIndex        =   12
      Tag             =   "1150107"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtPatId 
      Height          =   285
      HelpContextID   =   1150000
      Left            =   120
      MaxLength       =   13
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtPatName 
      Height          =   285
      HelpContextID   =   1150000
      Left            =   120
      MaxLength       =   50
      TabIndex        =   3
      Top             =   960
      Width           =   3015
   End
   Begin VB.ComboBox cboDictType 
      Height          =   315
      HelpContextID   =   1150000
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox cboPriority 
      Height          =   315
      HelpContextID   =   1150000
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblTxtTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Nyckelord:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Tag             =   "1150109"
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblPatIdTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "&Patientens personnr:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "1150101"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Patientens &namn:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Tag             =   "1150102"
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Diktattyp:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Tag             =   "1150103"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblPriorityTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "P&rioritet:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Tag             =   "1150104"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblAuthorTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "&Inläsare:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Tag             =   "1150105"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblTranscriberTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "&Utskrivare:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Tag             =   "1150106"
      Top             =   3120
      Width           =   1815
   End
End
Attribute VB_Name = "ucSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event NewSearch(ByRef SearchFilter As clsFilter)

Private mFilter As clsFilter
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub

Public Sub Init()

   txtPatId.Text = ""
   txtPatName.Text = ""
   Client.DictTypeMgr.FillCombo cboDictType
   cboDictType.ListIndex = -1
   Client.PriorityMgr.FillCombo cboPriority
   cboPriority.ListIndex = -1
   txtAuthor.Text = ""
   txtTranscriber.Text = ""
   txtTxt.Text = ""
   If Client.SysSettings.DictInfoUseKeyWords Then
      txtTxt.Visible = True
      lblTxtTitle.Visible = True
   Else
      txtTxt.Visible = False
      lblTxtTitle.Visible = False
   End If
   
   SetEnabled
End Sub


Private Sub cboDictType_Click()

   SetEnabled
End Sub

Private Sub cboPriority_Click()

   SetEnabled
End Sub

Private Sub cmdReset_Click()

   Init
End Sub

Private Sub cmdSearch_Click()

   Set mFilter = New clsFilter
   mFilter.Pat.PatId = StringReplace(txtPatId.Text, "-", "")
   mFilter.Pat.PatName = txtPatName.Text
   If cboDictType.ListIndex < 0 Then
      mFilter.DictTypeId = -1
   Else
      mFilter.DictTypeId = Client.DictTypeMgr.IdFromIndex(cboDictType.ListIndex)
   End If
   If cboPriority.ListIndex < 0 Then
      mFilter.PriorityId = -1
   Else
      mFilter.PriorityId = Client.PriorityMgr.IdFromIndex(cboPriority.ListIndex)
   End If
   mFilter.AuthorName = txtAuthor.Text
   mFilter.TranscriberName = txtTranscriber.Text
   mFilter.Txt = txtTxt.Text
   RaiseEvent NewSearch(mFilter)
End Sub

Private Sub SetEnabled()

   Dim B As Boolean
   
   B = Len(txtPatId.Text) > 0
   B = B Or Len(txtPatName.Text) > 0
   B = B Or cboDictType.ListIndex >= 0
   B = B Or cboPriority.ListIndex >= 0
   B = B Or Len(txtAuthor.Text) > 0
   B = B Or Len(txtTranscriber.Text) > 0
   B = B Or Len(txtTxt.Text) > 0
   
   cmdSearch.Enabled = B
End Sub

Private Sub txtAuthor_Change()

   SetEnabled
End Sub

Private Sub txtPatId_Change()

   SetEnabled
End Sub

Private Sub txtPatId_KeyPress(KeyAscii As Integer)

   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii < 32 Or KeyAscii = 45) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtPatName_Change()

   SetEnabled
End Sub

Private Sub txtTranscriber_Change()

   SetEnabled
End Sub

Private Sub txtTxt_Change()

   SetEnabled
End Sub
