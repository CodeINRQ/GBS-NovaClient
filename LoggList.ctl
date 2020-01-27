VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ucLoggList 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   ScaleHeight     =   4560
   ScaleWidth      =   9435
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   375
      HelpContextID   =   1330000
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   56426497
      CurrentDate     =   38595
      MaxDate         =   401768
      MinDate         =   38353
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Visa"
      Height          =   255
      HelpContextID   =   1330000
      Left            =   3000
      TabIndex        =   4
      Tag             =   "1330110"
      Top             =   360
      Width           =   1335
   End
   Begin FPSpreadADO.fpSpread lstLogg 
      Height          =   3735
      HelpContextID   =   1330000
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   9495
      _Version        =   458752
      _ExtentX        =   16748
      _ExtentY        =   6588
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
      SpreadDesigner  =   "LoggList.ctx":0000
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   375
      HelpContextID   =   1330000
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   56426497
      CurrentDate     =   38595
      MaxDate         =   401768
      MinDate         =   38353
   End
   Begin VB.Label lblEndDate 
      Caption         =   "T o m:"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Tag             =   "1330109"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblStartDate 
      Caption         =   "Fr o m:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Tag             =   "1330108"
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "ucLoggList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const White = &HFFFFFF
Const Red = &HC0C0FF
Const Yellow = &HC0FFFF
Const Gray = &HC0C0C0

Private StartDate As Date
Private EndDate As Date
Public Sub Init()

   dtpStartDate = Format$(Now, "ddddd")
   dtpEndDate = Format$(Now, "ddddd")
   StartDate = Int(Now)
   EndDate = DateAdd("d", 1, StartDate)
End Sub
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub
Public Sub ExportExcelFile(Fn As String)

   lstLogg.ExportExcelBookEx Fn, "", ExcelSaveFlagNone
End Sub
Public Sub ExportToHtml(Fn As String)

   lstLogg.ExportToHtml Fn, False, ""
End Sub
Public Sub ExportToXml(Fn As String)

   lstLogg.ExportToXml Fn, "", "", ExportToXMLFormattedData, ""
End Sub
Public Sub ExportTextFile(Fn As String)

   lstLogg.SaveTabFile Fn
End Sub
Public Sub ExportListToFile(DefFileName As String)

   Dim Fn As String
   Dim Ext As String
   
   If Len(DefFileName) = 0 Then
      DefFileName = Client.Texts.Txt(1000424, "Logg")
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

Private Sub cmdGo_Click()

   GetDataNow
End Sub

Private Sub dtpEndDate_Change()

   EndDate = DateAdd("d", 1, DateSerial(dtpEndDate.Year, dtpEndDate.Month, dtpEndDate.Day))
End Sub

Private Sub dtpStartDate_Change()

   StartDate = DateSerial(dtpStartDate.Year, dtpStartDate.Month, dtpStartDate.Day)
End Sub

Private Sub lstLogg_KeyPress(KeyAscii As Integer)

   Select Case KeyAscii
      Case KeyAsciiExportList
         If Client.SysSettings.ExportAllowMenu Then
            ExportListToFile ""
         End If
   End Select
End Sub

Private Sub UserControl_Resize()

   lstLogg.Move 0, 840, UserControl.ScaleWidth, UserControl.ScaleHeight - 840
End Sub
Public Sub RestoreSettings(Settings As String)

   SetupSpread lstLogg
   SetupSheet lstLogg
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
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1330101, "Tid"), TypeHAlignLeft, 15, White
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1330102, "LoggId"), TypeHAlignLeft, 8, White
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1330103, "Nivå"), TypeHAlignLeft, 6, White
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1330104, "Text"), TypeHAlignLeft, 20, White
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1330105, "Användare"), TypeHAlignLeft, 12, White
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1330106, "Station"), TypeHAlignLeft, 12, White
      Col = Col + 1: SetupColumn Spread, 1, Col, Client.Texts.Txt(1330107, "Data"), TypeHAlignLeft, 20, White
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

Private Sub GetDataNow()

   Dim Logg As clsLogg
   Dim Row As Integer
   Dim Col As Integer
   
   ClearWorkBook lstLogg
   Row = 1
   Client.LoggMgr.CreateList StartDate, EndDate, 0, 1000
   Do While Client.LoggMgr.GetNext(Logg)
      Col = 0
      Col = Col + 1: SetCellValue Row, Col, Format$(Logg.LoggTime, "ddddd ttttt")
      Col = Col + 1: SetCellValue Row, Col, CStr(Logg.LoggId)
      Col = Col + 1: SetCellValue Row, Col, CStr(Logg.LoggLevel)
      Col = Col + 1: SetCellValue Row, Col, Client.Texts.Txt(Logg.LoggId, "")
      Col = Col + 1: SetCellValue Row, Col, CStr(Logg.UserShortName)
      Col = Col + 1: SetCellValue Row, Col, Logg.StationId
      Col = Col + 1: SetCellValue Row, Col, Logg.LoggData
      Row = Row + 1
   Loop
   Set Logg = Nothing
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

   If lstLogg.MaxRows < Row Then
      lstLogg.MaxRows = Row
   End If

   lstLogg.Row = Row
   lstLogg.Col = lstLogg.GetColFromID(CStr(Col))
   lstLogg.Value = Txt
End Sub

