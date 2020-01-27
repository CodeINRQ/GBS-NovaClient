VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.UserControl ucEditUser 
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10155
   ScaleHeight     =   7185
   ScaleWidth      =   10155
   Begin VB.Frame fraUsers 
      Caption         =   "Användare"
      Height          =   7095
      HelpContextID   =   1130000
      Left            =   0
      TabIndex        =   7
      Tag             =   "1130101"
      Top             =   0
      Width           =   10095
      Begin VB.CheckBox chkSearch 
         Caption         =   "V&isa sök"
         Height          =   255
         Left            =   8520
         TabIndex        =   5
         Tag             =   "1130108"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   8520
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox chkShowAll 
         Caption         =   "&Visa alla"
         Height          =   375
         Left            =   8520
         TabIndex        =   6
         Tag             =   "1130107"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Lägg till..."
         Height          =   300
         Left            =   8520
         TabIndex        =   1
         Tag             =   "1130102"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Ändra..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   8520
         TabIndex        =   2
         Tag             =   "1130103"
         Top             =   600
         Width           =   1455
      End
      Begin FPSpreadADO.fpSpread lstUsers 
         Height          =   6255
         HelpContextID   =   1080000
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8295
         _Version        =   458752
         _ExtentX        =   14631
         _ExtentY        =   11033
         _StockProps     =   64
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   1
         MaxRows         =   0
         SpreadDesigner  =   "EditUser.ctx":0000
      End
      Begin VB.Label lblSearch 
         Caption         =   "&Sökning:"
         Height          =   255
         Left            =   8520
         TabIndex        =   3
         Tag             =   "1130109"
         Top             =   1800
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ucEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event UsersChanged()
Public Event RightClick(UserId As Long)

Private WithEvents frmEdit As frmEditUser
Attribute frmEdit.VB_VarHelpID = -1
Private CurrUser As clsUser
Private CurrentOrgId As Long


Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub
Public Sub ExportExcelFile(Fn As String)

   lstUsers.ExportExcelBookEx Fn, "", ExcelSaveFlagNone
End Sub
Public Sub ExportToHtml(Fn As String)

   lstUsers.ExportToHtml Fn, False, ""
End Sub
Public Sub ExportToXml(Fn As String)

   lstUsers.ExportToXml Fn, "", "", ExportToXMLFormattedData, ""
End Sub
Public Sub ExportTextFile(Fn As String)

   lstUsers.SaveTabFile Fn
End Sub
Public Sub ExportListToFile(DefFileName As String)

   Dim Fn As String
   Dim Ext As String
   
   If Len(DefFileName) = 0 Then
      DefFileName = Client.Texts.Txt(1130101, "Användare")
   End If
   
   Fn = GetExportFileName(DefFileName)
   If Len(Fn) = 0 Then Exit Sub
      
   Ext = LCase$(Right$(Fn, 3))
   Select Case Ext
      Case "xml"
         ExportToXml Fn
      Case "htm"
         ExportToHtml Fn
      Case "xls"
        ExportExcelFile Fn
      Case Else
        ExportTextFile Fn
   End Select
End Sub

Public Sub GetData(OrgId As Long)

   Dim I As Integer
   Dim LstIdx As Integer
   Dim Usr As clsUser
   Dim Row As Integer
   Static SettingDone As Boolean
   Dim ShowThisUser As Boolean
   
   If CurrentOrgId = OrgId Then
      Exit Sub
   End If
   CurrentOrgId = OrgId
   
   lstUsers.MaxRows = 0
   lstUsers.ClearRange -1, -1, -1, -1, True
   If Not SettingDone Then
      RestoreSettings
      SettingDone = True
   End If
   Row = 1
   For I = 0 To Client.UserMgr.Count - 1
      Client.UserMgr.GetUserFromIndex Usr, I
      
      ShowThisUser = False
      If Usr.HomeOrgId > 0 Then
         If Client.OrgMgr.CheckUserRole(Usr.HomeOrgId, RTUserAdmin) Then
            If chkShowAll.Value = vbChecked Then
               ShowThisUser = True
            Else
               If Client.OrgMgr.IsOrgBelowTop(OrgId, Usr.HomeOrgId) Then
                  ShowThisUser = True
               End If
            End If
         End If
      Else
         If chkShowAll.Value = vbChecked Then
            ShowThisUser = True
         End If
      End If
      If ShowThisUser And chkSearch.Value = vbChecked Then
         ShowThisUser = CheckSearchFilter(Usr, txtSearch.Text)
      End If
      If ShowThisUser Then
         lstUsers.MaxRows = Row
         UpdateRowInList Row, Usr
         Row = Row + 1
      End If
   Next I
   lstUsers.UserColAction = UserColActionSort
   SetEnabled
End Sub
Private Function CheckSearchFilter(U As clsUser, ByVal SearchCrit As String) As Boolean

   SearchCrit = UCase$(SearchCrit)
   
   CheckSearchFilter = InStr(UCase$(U.LoginName), SearchCrit) > 0 Or _
                       InStr(UCase$(U.ShortName), SearchCrit) > 0 Or _
                       InStr(UCase$(U.LongName), SearchCrit) > 0
End Function

Private Sub RestoreSettings()

   With lstUsers
      .ReDraw = False
      .Reset
       
      .OperationMode = OperationModeSingle
      .UserColAction = UserColActionSort
      .ColHeadersShow = True
      .ColHeaderDisplay = DispBlank
      .RowHeadersAutoText = DispBlank
      .AllowCellOverflow = False
      .AllowColMove = True
      
      .Col = -1
      .Row = -1
      .SelBackColor = &HC0C0C0
      .FontBold = False
      .TypeEditLen = 200
   
      .ReDraw = True
         
      .ShowScrollTips = ShowScrollTipsOff
      .TextTip = TextTipFloating
      
      .Col = 0:        .ColID = CStr(.Col):   .ColWidth(.Col) = 5
      .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 18
      .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 25
      .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 18
      .MaxCols = .Col
      
      SetCellValue 0, 0, "Id"
      SetCellValue 0, 1, Client.Texts.Txt(1130104, "Login")
      SetCellValue 0, 2, Client.Texts.Txt(1130105, "Namn")
      SetCellValue 0, 3, Client.Texts.Txt(1130106, "Org enhet")
   
      .RowHeadersShow = False
   End With
End Sub

Private Sub UpdateRowInList(Row As Integer, Usr As clsUser)

   Dim C As Integer
   Dim Mark As Integer
   Dim Ddiff As Integer

   With lstUsers
      .SetRowItemData Row, CStr(Usr.UserId)
      .Row = Row
      .Col = -1
      .ForeColor = 0
      .FontItalic = False
      
      C = 0:     SetCellValue Row, C, Usr.UserId
      C = C + 1: SetCellValue Row, C, Usr.LoginName
      C = C + 1: SetCellValue Row, C, Usr.LongName
      C = C + 1: SetCellValue Row, C, Client.OrgMgr.TextFromId(Usr.HomeOrgId)
   End With
End Sub
Sub SetCellValue(Row As Integer, Col As Integer, Txt As String)

   With lstUsers
      .Row = Row
      .Col = .GetColFromID(CStr(Col))
      .Value = Txt
   End With
End Sub

Private Sub SetEnabled()

   cmdChange.Enabled = True
End Sub

Private Sub chkSearch_Click()

   UpdateList
End Sub

Private Sub chkShowAll_Click()

   UpdateList
End Sub

Private Sub cmdChange_Click()

   ChangeOneUser
End Sub

Private Sub cmdNew_Click()

  Set CurrUser = New clsUser
  CurrUser.HomeOrgId = CurrentOrgId
  EditCurrUser
End Sub

Private Sub EditCurrUser()

  Set frmEdit = New frmEditUser
  Set frmEdit.UserToEdit = CurrUser
  frmEdit.Show vbModal
End Sub

Private Sub frmEdit_DeleteClicked()

   UpdateList
   RaiseEvent UsersChanged
   SetEnabled
End Sub

Private Sub frmEdit_SaveClicked()

   UpdateList
   RaiseEvent UsersChanged
   SetEnabled
End Sub

Private Sub lstUsers_AfterUserSort(ByVal Col As Long)

   Dim Sortkeys As Variant
   Dim SortKeyOrder As Variant
   Static Desc As Boolean
   
   Desc = lstUsers.ColUserSortIndicator(Col) = ColUserSortIndicatorDescending
   lstUsers.Col = Col
   Sortkeys = Array(Col, 0)
   If Desc Then
      SortKeyOrder = Array(2, 2)
   Else
      SortKeyOrder = Array(1, 1)
   End If
   lstUsers.Sort -1, -1, -1, -1, SortByRow, Sortkeys, SortKeyOrder
End Sub

Private Sub lstUsers_DblClick(ByVal Col As Long, ByVal Row As Long)

   ChangeOneUser
End Sub

Private Sub lstUsers_KeyPress(KeyAscii As Integer)
   
   Dim Fn As String

   Select Case KeyAscii
      Case 13
         ChangeOneUser
      Case KeyAsciiExportList
         If Client.SysSettings.ExportAllowMenu Then
            ExportListToFile ""
         End If
   End Select
End Sub

Private Sub lstUsers_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

   If Row > 0 Then
      RaiseEvent RightClick(CLng(lstUsers.GetRowItemData(Row)))
   End If
End Sub

Private Sub txtSearch_Change()

   If chkSearch.Value <> vbChecked Then
      chkSearch.Value = vbChecked
   Else
      UpdateList
   End If
End Sub

Private Sub txtSearch_GotFocus()

   SelectAllText ActiveControl
End Sub

Private Sub UserControl_Resize()

   fraUsers.Move 0, 0, fraUsers.Width, UserControl.ScaleHeight
   lstUsers.Move 120, 240, lstUsers.Width, UserControl.ScaleHeight - 280
End Sub
Private Sub ChangeOneUser()

   Client.UserMgr.GetUserFromId CurrUser, CLng(lstUsers.GetRowItemData(lstUsers.ActiveRow))
   EditCurrUser
End Sub
Private Sub UpdateList()

   Dim Temp As Long
   
   Temp = CurrentOrgId        'Force creation of a new list
   CurrentOrgId = -1
   GetData Temp
End Sub
