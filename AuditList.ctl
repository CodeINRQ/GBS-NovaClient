VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.UserControl ucAuditList 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   ScaleHeight     =   4560
   ScaleWidth      =   9435
   Begin FPSpreadADO.fpSpread lstAudit 
      Height          =   4575
      HelpContextID   =   1350000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _Version        =   458752
      _ExtentX        =   16748
      _ExtentY        =   8070
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
      SpreadDesigner  =   "AuditList.ctx":0000
   End
End
Attribute VB_Name = "ucAuditList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const White = &HFFFFFF
Const Red = &HC0C0FF
Const Yellow = &HC0FFFF
Const Gray = &HC0C0C0

Public DictId        As Long

Private Sub UserControl_Resize()

   lstAudit.Move 0, 0, UserControl.ScaleWidth - 100, UserControl.ScaleHeight - 100
End Sub
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub
Public Sub RestoreSettings(Settings As String)

   SetupSpread lstAudit
   SetupSheet lstAudit
End Sub
Private Sub SetupSpread(Spread As fpSpread)

   With Spread
      .Reset
      .SheetCount = 1
      .TabStripPolicy = TabStripPolicyNever
      .BackColorStyle = BackColorStyleUnderGrid
   End With
End Sub

Private Sub SetupSheet(Spread As fpSpread)

   Dim Col As Integer
   
   With Spread
      .ReDraw = False
            
      .RowHeadersShow = False
      .OperationMode = OperationModeRead
      .UserColAction = UserColActionAutoSize
      .ColHeadersShow = True
      .ColHeaderDisplay = DispBlank
      .RowHeadersAutoText = DispBlank
      .AllowCellOverflow = False
      .AllowColMove = False
      
      .Col = -1
      .Row = -1
      .SelBackColor = &HC0C0C0
      .FontBold = False
      .TypeEditLen = 200
               
      .ShowScrollTips = ShowScrollTipsOff
      .TextTip = TextTipFloating
      
      Col = 0
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1350101, "Tid"), TypeHAlignLeft, 15, White
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1350102, "Text"), TypeHAlignLeft, 15, White
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1350103, "Status"), TypeHAlignLeft, 15, White
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1350104, "Användare"), TypeHAlignLeft, 12, White
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1350105, "Station"), TypeHAlignLeft, 12, White
      .MaxCols = Col
      
      .ReDraw = True
   End With
End Sub
Private Sub SetupColumn(Spread As fpSpread, Sheet As Integer, Col As Integer, HeaderText As String, HAlign As Integer, InitWidth As Integer, Color As Long)

   With Spread
      .Sheet = Sheet
      .Col = Col
      .Row = 0
      .Text = HeaderText
      .ColID = CStr(Col)
      .ColWidth(Col) = InitWidth
      .Row = -1
      .TypeHAlign = HAlign
      .BackColor = Color
   End With
End Sub
Private Sub SetupRow(Spread As fpSpread, Sheet As Integer, Row As Integer, Color As Long)

   With Spread
      .Sheet = Sheet
      .Col = -1
      .Row = Row
      .BackColor = Color
   End With
End Sub

Public Sub GetDataNow()

   Dim Audit As clsDictAudit
   Dim Row As Integer
   Dim Col As Integer
   
   ClearWorkBook lstAudit
   Row = 1
   Client.DictAuditMgr.CreateList DictId
   Do While Client.DictAuditMgr.GetNext(Audit)
      Col = 0
      Col = Col + 1: SetCellValue Row, Col, Format$(Audit.AuditTime, "ddddd ttttt")
      Col = Col + 1: SetCellValue Row, Col, Client.Texts.Txt(1370100 + Audit.AuditType, CStr(Audit.AuditType))
      Col = Col + 1: SetCellValue Row, Col, Client.Texts.Txt(1250100 + Audit.DictStatus, CStr(Audit.DictStatus))
      Col = Col + 1: SetCellValue Row, Col, Audit.UserShortName
      Col = Col + 1: SetCellValue Row, Col, Audit.StationId
      Row = Row + 1
   Loop
   Set Audit = Nothing
End Sub
Private Sub ClearWorkBook(Spread As fpSpread)

   Dim Sh As Integer
   
   With Spread
      For Sh = 1 To .SheetCount
         .Sheet = Sh
         .MaxRows = 0
         .ClearRange -1, -1, -1, -1, True
      Next Sh
   End With
End Sub
Sub SetCellValue(Row As Integer, Col As Integer, Txt As String)

   If lstAudit.MaxRows < Row Then
      lstAudit.MaxRows = Row
   End If

   lstAudit.Row = Row
   lstAudit.Col = lstAudit.GetColFromID(CStr(Col))
   lstAudit.Value = Txt
End Sub


