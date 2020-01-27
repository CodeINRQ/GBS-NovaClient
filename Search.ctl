VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ucSearch 
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   DefaultCancel   =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   4815
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   120
      TabIndex        =   25
      Text            =   "Combo1"
      Top             =   6600
      Width           =   3015
   End
   Begin VB.CheckBox chkSearchOrg 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6000
      Width           =   255
   End
   Begin VB.CheckBox chkTranscribedDate 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkRecDate 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtTxt 
      Height          =   285
      HelpContextID   =   1150000
      Left            =   120
      MaxLength       =   50
      TabIndex        =   21
      Top             =   5400
      Width           =   3015
   End
   Begin VB.TextBox txtTranscriber 
      Height          =   285
      HelpContextID   =   1150000
      Left            =   120
      MaxLength       =   50
      TabIndex        =   19
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      HelpContextID   =   1150000
      Left            =   120
      MaxLength       =   50
      TabIndex        =   17
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton cmdReset 
      Cancel          =   -1  'True
      Caption         =   "Åt&erställ"
      Height          =   310
      HelpContextID   =   1150000
      Left            =   3360
      TabIndex        =   27
      Tag             =   "1150108"
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Sök"
      Default         =   -1  'True
      Height          =   310
      HelpContextID   =   1150000
      Left            =   3360
      TabIndex        =   26
      Tag             =   "1150107"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtPatId 
      Height          =   285
      HelpContextID   =   1150000
      Left            =   120
      MaxLength       =   14
      TabIndex        =   1
      Top             =   360
      Width           =   1455
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
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.ComboBox cboPriority 
      Height          =   315
      HelpContextID   =   1150000
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtpRecStartDate 
      Height          =   375
      HelpContextID   =   1330000
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16842753
      CurrentDate     =   38595
      MaxDate         =   401768
      MinDate         =   38353
   End
   Begin MSComCtl2.DTPicker dtpRecEndDate 
      Height          =   375
      HelpContextID   =   1330000
      Left            =   1920
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16842753
      CurrentDate     =   38595
      MaxDate         =   401768
      MinDate         =   38353
   End
   Begin MSComCtl2.DTPicker dtpTranscribedStartDate 
      Height          =   375
      HelpContextID   =   1330000
      Left            =   480
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16842753
      CurrentDate     =   38595
      MaxDate         =   401768
      MinDate         =   38353
   End
   Begin MSComCtl2.DTPicker dtpTranscribedEndDate 
      Height          =   375
      HelpContextID   =   1330000
      Left            =   1920
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16842753
      CurrentDate     =   38595
      MaxDate         =   401768
      MinDate         =   38353
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Tag             =   "1150113"
      Top             =   6360
      Width           =   3015
   End
   Begin VB.Label lblSearchOrg 
      BackStyle       =   0  'Transparent
      Caption         =   "Sö&k bara i"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Tag             =   "1150112"
      Top             =   5760
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "U&tskrivet:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Tag             =   "1150111"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Inta&lat:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Tag             =   "1150110"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblTxtTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "N&yckelord:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Tag             =   "1150109"
      Top             =   5160
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
      TabIndex        =   12
      Tag             =   "1150103"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblPriorityTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "P&rioritet:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Tag             =   "1150104"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblAuthorTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "&Inläsare:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Tag             =   "1150105"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblTranscriberTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "&Utskrivare:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Tag             =   "1150106"
      Top             =   4560
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
Public mOrgId As Long

Private mFilter As clsFilter
Private RecStartDate As Date
Private RecEndDate As Date
Private TranscribedStartDate As Date
Private TranscribedEndDate As Date

Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub

Public Sub Init()

   Dim Idx As Integer

   txtPatId.Text = ""
   txtPatName.Text = ""
   
   chkRecDate.Value = Unchecked
   RecStartDate = DateAdd("m", -1, Int(Now))
   dtpRecStartDate = Format$(RecStartDate, "ddddd")
   RecEndDate = DateAdd("d", 1, Int(Now))
   dtpRecEndDate = Format$(Now, "ddddd")

   chkTranscribedDate.Value = Unchecked
   TranscribedStartDate = DateAdd("m", -1, Int(Now))
   dtpTranscribedStartDate = Format$(RecStartDate, "ddddd")
   TranscribedEndDate = DateAdd("d", 1, Int(Now))
   dtpTranscribedEndDate = Format$(Now, "ddddd")

   Client.DictTypeMgr.FillCombo cboDictType, -1, -1, False
   Client.PriorityMgr.FillCombo cboPriority, -1, -1, False
   
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
   
   chkSearchOrg.Value = vbUnchecked
   
   cboStatus.Clear
   Idx = 0
   cboStatus.AddItem "", Idx
   cboStatus.ItemData(Idx) = 0
   Idx = Idx + 1
   cboStatus.AddItem Client.Texts.Txt(1250100 + 20, ""), Idx
   cboStatus.ItemData(Idx) = 20
   Idx = Idx + 1
   cboStatus.AddItem Client.Texts.Txt(1250100 + 30, ""), Idx
   cboStatus.ItemData(Idx) = 30
   Idx = Idx + 1
   cboStatus.AddItem Client.Texts.Txt(1250100 + 50, ""), Idx
   cboStatus.ItemData(Idx) = 50
   If Client.SysSettings.UseAuthorsSign Then
      Idx = Idx + 1
      cboStatus.AddItem Client.Texts.Txt(1250100 + 60, ""), Idx
      cboStatus.ItemData(Idx) = 60
   End If
   Idx = Idx + 1
   cboStatus.AddItem Client.Texts.Txt(1250100 + 70, ""), Idx
   cboStatus.ItemData(Idx) = 70
      
   SetEnabled
End Sub


Private Sub cboDictType_Click()

   SetEnabled
End Sub

Private Sub cboPriority_Click()

   SetEnabled
End Sub

Private Sub chkRecDate_Click()

   dtpRecStartDate.Enabled = chkRecDate.Value = Checked
   dtpRecEndDate.Enabled = chkRecDate.Value = Checked
   SetEnabled
End Sub

Private Sub chkTranscribedDate_Click()

   dtpTranscribedStartDate.Enabled = chkTranscribedDate.Value = Checked
   dtpTranscribedEndDate.Enabled = chkTranscribedDate.Value = Checked
   SetEnabled
End Sub

Private Sub cmdReset_Click()

   Init
End Sub

Private Sub cmdSearch_Click()

   Dim Status As Integer
  
   Set mFilter = New clsFilter
   
   mFilter.FilterFromSearch = True
   
   mFilter.Pat.PatId = StringReplace(txtPatId.Text, "-", "")
   mFilter.Pat.PatName = txtPatName.Text
   
   mFilter.RecDateUsed = chkRecDate.Value = Checked
   If mFilter.RecDateUsed Then
      mFilter.RecStartDate = RecStartDate
      mFilter.RecEndDate = RecEndDate
   End If
   
   mFilter.TranscribedDateUsed = chkTranscribedDate.Value = Checked
   If mFilter.TranscribedDateUsed Then
      mFilter.TranscribedStartDate = TranscribedStartDate
      mFilter.TranscribedEndDate = TranscribedEndDate
   End If
   
   If cboDictType.ListIndex < 0 Then
      mFilter.DictTypeId = -1
   Else
      mFilter.DictTypeId = cboDictType.ItemData(cboDictType.ListIndex)
   End If
   If cboPriority.ListIndex < 0 Then
      mFilter.PriorityId = -1
   Else
      mFilter.PriorityId = cboPriority.ItemData(cboPriority.ListIndex)
   End If
   mFilter.AuthorName = txtAuthor.Text
   mFilter.TranscriberName = txtTranscriber.Text
   mFilter.Txt = txtTxt.Text
   If mOrgId < 30000 Then
      If chkSearchOrg.Value = vbChecked Then
         mFilter.OrgId = mOrgId
      End If
   End If
   If cboStatus.ListIndex > 0 Then
      Status = cboStatus.ItemData(cboStatus.ListIndex)
      mFilter.StatusStart = Status
      If Status >= 70 Then
         mFilter.StatusEnd = 100
      Else
         mFilter.StatusEnd = mFilter.StatusStart
      End If
   End If
   RaiseEvent NewSearch(mFilter)
End Sub
Public Sub SetNewCurrentOrg(OrgId As Long, OrgText As String)

   Dim S As String
   Dim LastDictTypeId As Integer
   Dim LastPriorityId As Integer
   
   mOrgId = OrgId
   
   If cboDictType.ListIndex >= 0 Then
      LastDictTypeId = cboDictType.ItemData(cboDictType.ListIndex)
   Else
      LastDictTypeId = -1
   End If
   If cboPriority.ListIndex >= 0 Then
      LastPriorityId = cboPriority.ItemData(cboPriority.ListIndex)
   Else
      LastPriorityId = -1
   End If
   S = Client.Texts.Txt(1150112, "")
   If mOrgId < 30000 Then
      Client.DictTypeMgr.FillCombo cboDictType, mOrgId, LastDictTypeId, False
      Client.PriorityMgr.FillCombo cboPriority, mOrgId, LastPriorityId, False
      S = S & " " & OrgText
   Else
      Client.DictTypeMgr.FillCombo cboDictType, Client.User.HomeOrgId, LastDictTypeId, False
      Client.PriorityMgr.FillCombo cboPriority, Client.User.HomeOrgId, LastPriorityId, False
   End If
   lblSearchOrg.Caption = S
   
   SetEnabled
End Sub

Private Sub SetEnabled()

   Dim B As Boolean
   
   B = Len(txtPatId.Text) > 0
   B = B Or Len(txtPatName.Text) > 0
   B = B Or chkRecDate.Value = Checked
   B = B Or chkTranscribedDate.Value = Checked
   B = B Or cboDictType.ListIndex >= 0
   B = B Or cboPriority.ListIndex >= 0
   B = B Or Len(txtAuthor.Text) > 0
   B = B Or Len(txtTranscriber.Text) > 0
   B = B Or Len(txtTxt.Text) > 0
   
   chkSearchOrg.Enabled = mOrgId < 30000
   
   cmdSearch.Enabled = B
End Sub

Private Sub dtpRecEndDate_Change()

   RecEndDate = DateAdd("d", 1, DateSerial(dtpRecEndDate.Year, dtpRecEndDate.Month, dtpRecEndDate.Day))
End Sub

Private Sub dtpRecStartDate_Change()

   RecStartDate = DateSerial(dtpRecStartDate.Year, dtpRecStartDate.Month, dtpRecStartDate.Day)
End Sub

Private Sub dtpTranscribedEndDate_Change()

   TranscribedEndDate = DateAdd("d", 1, DateSerial(dtpTranscribedEndDate.Year, dtpTranscribedEndDate.Month, dtpTranscribedEndDate.Day))
End Sub


Private Sub dtpTranscribedStartDate_Change()

   TranscribedStartDate = DateSerial(dtpTranscribedStartDate.Year, dtpTranscribedStartDate.Month, dtpTranscribedStartDate.Day)
End Sub

Private Sub txtAuthor_Change()

   SetEnabled
End Sub

Private Sub txtAuthor_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtPatId_Change()

   SetEnabled
End Sub

Private Sub txtPatId_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtPatId_KeyPress(KeyAscii As Integer)

   If Not Client.SysSettings.DictInfoAlfaInPatid Then
      If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii < 32 Or KeyAscii = 45) Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtPatName_Change()

   SetEnabled
End Sub

Private Sub txtPatName_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtTranscriber_Change()

   SetEnabled
End Sub

Private Sub txtTranscriber_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub txtTxt_Change()

   SetEnabled
End Sub

Private Sub txtTxt_GotFocus()

   SelectAllText ActiveControl
End Sub

