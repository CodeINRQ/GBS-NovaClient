VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.UserControl ucOrgPriority 
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   ScaleHeight     =   3000
   ScaleWidth      =   8220
   Begin VB.Frame fraOrgPriority 
      Caption         =   "Prioriteter"
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Tag             =   "1420101"
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdSave 
         Caption         =   "Spara"
         Height          =   300
         Left            =   6000
         TabIndex        =   1
         Tag             =   "1420102"
         Top             =   240
         Width           =   2055
      End
      Begin FPSpreadADO.fpSpread lstOrgPriority 
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5775
         _Version        =   458752
         _ExtentX        =   10186
         _ExtentY        =   4471
         _StockProps     =   64
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         RowHeaderDisplay=   0
         SpreadDesigner  =   "OrgPriority.ctx":0000
      End
   End
End
Attribute VB_Name = "ucOrgPriority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim CurrOrgId As Long

Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub
Public Sub ExportExcelFile(Fn As String)

   lstOrgPriority.ExportExcelBookEx Fn, "", ExcelSaveFlagNone
End Sub
Public Sub ExportToHtml(Fn As String)

   lstOrgPriority.ExportToHtml Fn, False, ""
End Sub
Public Sub ExportToXml(Fn As String)

   lstOrgPriority.ExportToXml Fn, "", "", ExportToXMLFormattedData, ""
End Sub
Public Sub ExportTextFile(Fn As String)

   lstOrgPriority.SaveTabFile Fn
End Sub
Public Sub ExportListToFile(DefFileName As String)

   Dim Fn As String
   Dim Ext As String
   
   If Len(DefFileName) = 0 Then
      DefFileName = Client.Texts.Txt(1420101, "Prioriteter")
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

Public Sub Init()

   lstOrgPriority.MaxRows = 1
   lstOrgPriority.ClearRange -1, -1, -1, -1, True
   RestoreSettings
   
End Sub
Sub SetCellValue(Row As Integer, Col As Integer, Txt As String)

   With lstOrgPriority
      .Row = Row
      .Col = .GetColFromID(CStr(Col))
      .Value = Txt
   End With
End Sub
Sub SetCellBool(Row As Integer, Col As Integer, BoolVaue As Boolean)

   With lstOrgPriority
      .Row = Row
      .Col = .GetColFromID(CStr(Col))
      .CellType = CellTypeCheckBox
      .TypeCheckType = TypeCheckTypeNormal
      .TypeVAlign = TypeVAlignCenter
      .TypeHAlign = TypeHAlignCenter
      .Value = BoolVaue
   End With
End Sub
Private Sub RestoreSettings()

   With lstOrgPriority
      .ReDraw = False
      .Reset
       
      .OperationMode = OperationModeNormal
      .UserColAction = UserColActionAutoSize
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
      .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 25
      .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
      .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
      .MaxCols = .Col
      
      SetCellValue 0, 0, "Id"
      SetCellValue 0, 1, Client.Texts.Txt(1420103, "Prioritet")
      SetCellValue 0, 2, Client.Texts.Txt(1420104, "Används")
      SetCellValue 0, 3, Client.Texts.Txt(1420105, "Förinst")
   
      .RowHeadersShow = False
   End With
End Sub
Private Sub ShowPriorityForOrg()

   Static LastOrgId As Long
   Dim OId As Long
   Dim Row As Integer
   Dim Priority As clsPriority
   Dim OrgPriority As clsOrgPriority
   Dim PIdx As Integer
   Dim PriorityEnabled As Boolean
   Dim PriorityDefault As Boolean
   
   Dim Org As clsOrg
   
   cmdSave.Enabled = False
   Set Org = Nothing
   Client.OrgMgr.GetOrgFromId Org, CurrOrgId
   lstOrgPriority.ClearRange -1, -1, -1, -1, True
   
   If CurrOrgId < 30000 Then
      If Not Org Is Nothing Then
         fraOrgPriority.Caption = Client.Texts.Txt(fraOrgPriority.Tag, "Prioriteter") & " " & Org.OrgText

         Client.PriorityMgr.Init
         
         lstOrgPriority.MaxRows = 0
         Row = 1
         
         For PIdx = 0 To Client.PriorityMgr.Count - 1
            Client.PriorityMgr.GetFromIndex Priority, PIdx
            Client.PriorityMgr.GetOrgPriorityFromId OrgPriority, CurrOrgId, Priority.PriorityId
            If Not OrgPriority Is Nothing Then
               PriorityEnabled = True
               PriorityDefault = OrgPriority.Def
            Else
               PriorityEnabled = False
               PriorityDefault = False
            End If
            lstOrgPriority.MaxRows = Row
            UpdateRowInList Row, Priority.PriorityId, Priority.PriortyText, PriorityEnabled, PriorityDefault
            Row = Row + 1
         Next PIdx
      End If
   End If
End Sub
Private Sub UpdateRowInList(Row As Integer, PriorityId As Integer, PriorityTxt As String, PriorityEnabled As Boolean, PriorityDefault As Boolean)

   Dim C As Integer

   With lstOrgPriority
      .SetRowItemData Row, CStr(PriorityId)
      .Row = Row

      C = 0:     SetCellValue Row, C, CStr(PriorityId)
      C = C + 1: SetCellValue Row, C, PriorityTxt
      C = C + 1: SetCellBool Row, C, PriorityEnabled
      C = C + 1: SetCellBool Row, C, PriorityDefault
   End With
End Sub

Public Sub NewOrg(OrgId As Long)

   If CurrOrgId <> OrgId Then
      CurrOrgId = OrgId
      ShowPriorityForOrg
   End If
End Sub

Private Sub cmdSave_Click()

   Dim R As Integer
   Dim E As Boolean
   Dim D As Boolean

   cmdSave.Enabled = False
   With lstOrgPriority
      Client.PriorityMgr.DeleteOrgPriorityByOrgId CurrOrgId
      
      For R = 1 To .MaxRows
         .Row = R
         .Col = 1: Debug.Print .Value
         .Col = 2: E = .Value
         .Col = 3: D = .Value
      
         If E Then
            Client.PriorityMgr.SaveOrgPriority CurrOrgId, CInt(.GetRowItemData(R)), D
         End If
      Next R
   End With
   Client.PriorityMgr.Init
   ShowPriorityForOrg
End Sub

Private Sub lstOrgPriority_EditChange(ByVal Col As Long, ByVal Row As Long)

   cmdSave.Enabled = True
End Sub

Private Sub lstOrgPriority_KeyPress(KeyAscii As Integer)

   Select Case KeyAscii
      Case KeyAsciiExportList
         If Client.SysSettings.ExportAllowMenu Then
            ExportListToFile ""
         End If
   End Select
End Sub

Private Sub lstOrgPriority_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   cmdSave.Enabled = True
End Sub
