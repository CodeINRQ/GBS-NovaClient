VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.UserControl ucHistList 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   ScaleHeight     =   4560
   ScaleWidth      =   9435
   Begin VB.CommandButton cmdGo 
      Caption         =   "Visa"
      Height          =   255
      HelpContextID   =   1140000
      Left            =   4560
      TabIndex        =   2
      Tag             =   "1140101"
      Top             =   30
      Width           =   1335
   End
   Begin VB.ComboBox cboHistYear 
      Height          =   315
      HelpContextID   =   1140000
      ItemData        =   "HistList.ctx":0000
      Left            =   3360
      List            =   "HistList.ctx":0002
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      HelpContextID   =   1140000
      ItemData        =   "HistList.ctx":0004
      Left            =   0
      List            =   "HistList.ctx":0006
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin FPSpreadADO.fpSpread lstHist 
      Height          =   4215
      HelpContextID   =   1140000
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   9495
      _Version        =   458752
      _ExtentX        =   16748
      _ExtentY        =   7435
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
      SpreadDesigner  =   "HistList.ctx":0008
   End
End
Attribute VB_Name = "ucHistList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const White = &HFFFFFF
Const Red = &HC0C0FF
Const Yellow = &HC0FFFF
Const Gray = &HC0C0C0

Const SheetNumber = 1
Const SheetLenSec = 2
Const SheetMax = 2

Const MaxRows = 1000
Const MaxCols = 15

Dim NumberOfFooterRow   As Integer
Dim HistType            As HistTypeEnum
Dim HistYear            As Integer

Dim CurrentOrgId As Long
Dim NextOrgId As Long
Dim CurrentHistType As HistTypeEnum
Dim CurrentHistYear As Integer

Private RowTotal(SheetMax, MaxRows) As Long
Private ColTotal(SheetMax, MaxCols) As Long
Private GrandTotal(SheetMax) As Long
Public Sub ExportExcelFile(Fn As String)

   lstHist.ExportExcelBookEx Fn, "", ExcelSaveFlagNone
End Sub
Public Sub ExportToHtml(Fn As String)

   lstHist.ExportToHtml Fn, False, ""
End Sub
Public Sub ExportToXml(Fn As String)

   lstHist.ExportToXml Fn, "", "", ExportToXMLFormattedData, ""
End Sub
Public Sub ExportTextFile(Fn As String)

   lstHist.SaveTabFile Fn
End Sub
Public Sub ExportListToFile(DefFileName As String)

   Dim Fn As String
   Dim Ext As String
   
   If Len(DefFileName) = 0 Then
      DefFileName = Client.Texts.Txt(1000405, "Historik")
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

   cboType.Clear
   cboType.AddItem Client.Texts.Txt(1140102, "Prioritet")
   cboType.AddItem Client.Texts.Txt(1140103, "Diktattyp")
   cboType.AddItem Client.Texts.Txt(1140104, "Organisation")
   cboType.AddItem Client.Texts.Txt(1140120, "Utskrivare")
   cboType.AddItem Client.Texts.Txt(1140121, "Utskrivande enhet")
   cboType.AddItem Client.Texts.Txt(1140122, "Medel dagar efter intalande")
   cboType.AddItem Client.Texts.Txt(1140123, "Medel dagar efter skriv senast")
   cboType.AddItem Client.Texts.Txt(1140124, "Intalare")
   cboType.ListIndex = 0
   HistType = 0
   
   cboHistYear.Clear
   cboHistYear.AddItem CStr(Year(Now))
   cboHistYear.AddItem CStr(Year(Now) - 1)
   cboHistYear.AddItem CStr(Year(Now) - 2)
   cboHistYear.AddItem CStr(Year(Now) - 3)
   cboHistYear.AddItem CStr(Year(Now) - 4)
   cboHistYear.AddItem CStr(Year(Now) - 5)
   cboHistYear.AddItem CStr(Year(Now) - 6)
   cboHistYear.ListIndex = 0
   HistYear = Year(Now)
End Sub
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub


Private Sub cmdGo_Click()

   HistYear = Year(Now) - cboHistYear.ListIndex
   HistType = cboType.ListIndex
   GetDataNow
End Sub

Private Sub lstHist_KeyPress(KeyAscii As Integer)

   Select Case KeyAscii
      Case KeyAsciiExportList
         If Client.SysSettings.ExportAllowMenu Then
            ExportListToFile ""
         End If
   End Select
End Sub

Private Sub UserControl_Resize()

   lstHist.Move 0, 360, UserControl.ScaleWidth, UserControl.ScaleHeight - 360
End Sub
Public Sub RestoreSettings(Settings As String)

   'Dim TempFilePath As String
   'Dim Ok As Boolean

   With lstHist
      'Ok = False
      'If Len(Settings) > 0 Then
      '   TempFilePath = WriteStringToTempFile(Settings)
      '   If .LoadFromFile(TempFilePath) Then
      '      Ok = True
      '   End If
      '   KillFileIgnoreError TempFilePath
      'End If
      'If Not Ok Then
         SetupSpread lstHist
         SetupSheet lstHist, 1, Client.Texts.Txt(1140105, "Antal")
         SetupSheet lstHist, 2, Client.Texts.Txt(1140106, "Längd")
      'End If
   End With
End Sub
Private Sub SetupSpread(Spread As fpSpread)

   With Spread
      .Reset
      .SheetCount = 2
      .TabStripPolicy = TabStripPolicyAlways
      .BackColorStyle = BackColorStyleUnderGrid
   End With
End Sub

Private Sub SetupSheet(Spread As fpSpread, Nr As Integer, Name As String)

   Dim Col As Integer
   
   With Spread
      .ReDraw = False
      .Sheet = Nr
      .SheetName = Name
            
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
      
      Col = 1: SetupColumn Spread, Nr, Col, "", TypeHAlignLeft, 20, vbButtonFace
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140107, "Jan"), TypeHAlignRight, 6, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140108, "Feb"), TypeHAlignRight, 6, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140109, "Mar"), TypeHAlignRight, 6, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140110, "Apr"), TypeHAlignRight, 6, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140111, "Maj"), TypeHAlignRight, 6, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140112, "Jun"), TypeHAlignRight, 6, Yellow
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140113, "Jul"), TypeHAlignRight, 6, Yellow
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140114, "Aug"), TypeHAlignRight, 6, Yellow
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140115, "Sep"), TypeHAlignRight, 6, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140116, "Okt"), TypeHAlignRight, 6, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140117, "Nov"), TypeHAlignRight, 6, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140118, "Dec"), TypeHAlignRight, 6, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1140119, "Summa"), TypeHAlignRight, 8, vbButtonFace
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
Private Sub ClearTotals()

   Dim I As Integer
   Dim Sh As Integer
   
   For Sh = 0 To SheetMax
      For I = 0 To MaxRows
         RowTotal(Sh, I) = 0
      Next I
      For I = 0 To MaxCols
         ColTotal(Sh, I) = 0
      Next I
      GrandTotal(Sh) = 0
   Next Sh
End Sub
Public Sub GetData(OrgId As Long)

   If NextOrgId <> OrgId Then
      NextOrgId = OrgId
      'ClearWorkBook lstHist
      GetDataNow
   End If
End Sub

Private Sub GetDataNow()

   Dim I As Integer
   Dim NumCol As Integer
   Dim W As Long
   Dim Hist As clsHistory
   Dim Row As Integer
   Dim Org As clsOrg
   Dim NumberOfRows As Integer
   Dim RowHeader As String
   Dim Col As Integer
   Dim Usr As clsUser
   Dim NumberAndLength As Boolean
   Dim AvgDays As Single
   
   If CurrentOrgId <> NextOrgId Or CurrentHistType <> HistType Or CurrentHistYear <> HistYear Then
      CurrentOrgId = NextOrgId
      CurrentHistType = HistType
      CurrentHistYear = HistYear

      ClearWorkBook lstHist
      ClearTotals
      Row = 1
      NumberAndLength = Not (CurrentHistType = htDaysFromCreated Or CurrentHistType = htDaysFromExpiry)
      If NumberAndLength Then
         lstHist.TabStripPolicy = TabStripPolicyAlways
      Else
         lstHist.ActiveSheet = 1
         lstHist.TabStripPolicy = TabStripPolicyNever
      End If
      Client.HistMgr.CreateList CurrentOrgId, CurrentHistType, CurrentHistYear
      Do While Client.HistMgr.ListNextItem(Hist)
         NumberOfRows = NumberOfRows + 1
         Col = 1
         Select Case HistType
            Case htPrio
               RowHeader = Client.PriorityMgr.TextFromId(Hist.Rowid)
            Case htDictType
               RowHeader = Client.DictTypeMgr.TextFromId(Hist.Rowid)
            Case htOrg
               RowHeader = Client.OrgMgr.TextFromId(Hist.Rowid)
            Case htTranscriber, htAuthor
               If Client.UserMgr.GetUserFromId(Usr, Hist.Rowid) Then
                  RowHeader = Usr.ShortName
               Else
                  RowHeader = ""
               End If
            Case htTranscriberOrg
               RowHeader = Client.OrgMgr.TextFromId(Hist.Rowid)
            Case htDaysFromCreated
               RowHeader = Client.PriorityMgr.TextFromId(Hist.Rowid)
            Case htDaysFromExpiry
               RowHeader = Client.PriorityMgr.TextFromId(Hist.Rowid)
            Case Else
               RowHeader = ""
         End Select
         If NumberAndLength Then
            SetCellString SheetNumber, Row, Col, RowHeader
            SetCellString SheetLenSec, Row, Col, RowHeader
            For I = 1 To 12
               Col = Col + 1: SetCell Row, Col, Hist.Number(I), Hist.SoundLenSec(I), True
            Next I
            Col = Col + 1: SetCell Row, Col, RowTotal(SheetNumber, Row), RowTotal(SheetLenSec, Row), False
         Else
            SetCellString SheetNumber, Row, Col, RowHeader
            For I = 1 To 12
               If Hist.SoundLenSec(I) > 0 Then
                  AvgDays = Hist.Number(I) / Hist.SoundLenSec(I)
               Else
                  AvgDays = 0
               End If
               Col = Col + 1: SetCell Row, Col, Format$(AvgDays, "0.0"), 0, False
            Next I
         End If
         Row = Row + 1
      Loop
      If NumberOfRows > 1 And NumberAndLength Then
         AddFooterRow
         NumberOfFooterRow = 1
      Else
         NumberOfFooterRow = 0
      End If
   End If
   Set Hist = Nothing
End Sub
Private Sub AddFooterRow()

   Dim Row As Integer
   Dim Col As Integer

   With lstHist
      Row = .MaxRows + 1
      SetCellString SheetNumber, Row, 1, "Summa"
      SetCellString SheetLenSec, Row, 1, "Summa"
      For Col = 2 To .MaxCols - 1
         SetCell Row, Col, ColTotal(SheetNumber, Col), ColTotal(SheetLenSec, Col), False
      Next Col
      SetCell Row, .MaxCols, GrandTotal(SheetNumber), GrandTotal(SheetLenSec), False
      SetupRow lstHist, SheetNumber, Row, vbButtonFace
      SetupRow lstHist, SheetLenSec, Row, vbButtonFace
   End With
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
Private Sub SetCell(Row As Integer, Col As Integer, Num As Long, LenSec As Long, AddToTotals As Boolean)

   SetCellNumber SheetNumber, Row, Col, Num
   If AddToTotals Then
      RowTotal(SheetNumber, Row) = RowTotal(SheetNumber, Row) + Num
      ColTotal(SheetNumber, Col) = ColTotal(SheetNumber, Col) + Num
      GrandTotal(SheetNumber) = GrandTotal(SheetNumber) + Num
   End If
   
   SetCellTime SheetLenSec, Row, Col, LenSec
   If AddToTotals Then
      RowTotal(SheetLenSec, Row) = RowTotal(SheetLenSec, Row) + LenSec
      ColTotal(SheetLenSec, Col) = ColTotal(SheetLenSec, Col) + LenSec
      GrandTotal(SheetLenSec) = GrandTotal(SheetLenSec) + LenSec
   End If
End Sub

Sub SetCellString(Sheet As Integer, Row As Integer, Col As Integer, Txt As String)

   With lstHist
      .Sheet = Sheet
      If .MaxRows < Row Then
         .MaxRows = Row
      End If
      .Row = Row
      .Col = .GetColFromID(CStr(Col))
      .Value = Txt
   End With
End Sub
Sub SetCellNumber(Sheet As Integer, Row As Integer, Col As Integer, Num As Long)

   Dim s As String

   If Num <> 0 Then
      s = CStr(Num)
   End If
   SetCellString Sheet, Row, Col, s
End Sub
Sub SetCellTime(Sheet As Integer, Row As Integer, Col As Integer, Sec As Long)

   Dim s As String

   If Sec > 0 Then
      s = FormatLength(Sec)
   End If
   SetCellString Sheet, Row, Col, s
End Sub

Public Function GetSetting() As String

   'Dim Pathname As String

   'ClearWorkBook lstHist
   'Pathname = CreateTempFileName("tmp")
   'lstHist.SaveToFile Pathname, False
   'GetSetting = ReadStringFromTempFile(Pathname)
   'KillFileIgnoreError Pathname
End Function
