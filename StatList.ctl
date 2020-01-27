VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.UserControl ucStatList 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   ScaleHeight     =   4560
   ScaleWidth      =   9435
   Begin VB.CommandButton cmdGo 
      Caption         =   "Visa"
      Height          =   255
      HelpContextID   =   1160000
      Left            =   3360
      TabIndex        =   0
      Tag             =   "1160101"
      Top             =   30
      Width           =   1335
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      HelpContextID   =   1160000
      ItemData        =   "StatList.ctx":0000
      Left            =   0
      List            =   "StatList.ctx":0002
      TabIndex        =   1
      Top             =   0
      Width           =   3255
   End
   Begin FPSpreadADO.fpSpread lstStat 
      Height          =   4215
      HelpContextID   =   1160000
      Left            =   0
      TabIndex        =   2
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
      SpreadDesigner  =   "StatList.ctx":0004
   End
End
Attribute VB_Name = "ucStatList"
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
Const MaxCols = 10

Dim NumberOfFooterRow As Integer

Dim CurrentOrgId As Long
Dim NextOrgId As Long

Private RowTotal(SheetMax, MaxRows) As Long
Private ColTotal(SheetMax, MaxCols) As Long
Private GrandTotal(SheetMax) As Long
Public Sub ExportExcelFile(Fn As String)

   lstStat.ExportExcelBookEx Fn, "", ExcelSaveFlagNone
End Sub
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub

Private Sub cmdGo_Click()

   GetDataNow
End Sub

Public Sub Init()

   cboType.AddItem Client.Texts.Txt(1160102, "Fördelning skriv-senast")
   cboType.ListIndex = 0
End Sub

Private Sub UserControl_Resize()

   lstStat.Move 0, 360, UserControl.ScaleWidth, UserControl.ScaleHeight - 360
End Sub
Public Sub RestoreSettings(Settings As String)

   'Dim TempFilePath As String
   'Dim Ok As Boolean

   With lstStat
      'Ok = False
      'If Len(Settings) > 0 Then
      '   TempFilePath = WriteStringToTempFile(Settings)
      '   If .LoadFromFile(TempFilePath) Then
      '      Ok = True
      '   End If
      '   KillFileIgnoreError TempFilePath
      'End If
      'If Not Ok Then
         SetupSpread lstStat
         SetupSheet lstStat, 1, Client.Texts.Txt(1160103, "Antal")
         SetupSheet lstStat, 2, Client.Texts.Txt(1160104, "Längd")
      'End If
   End With
End Sub
Private Sub SetupSpread(Spread As fpSpread)

   With Spread
      .Reset
      .SheetCount = 2
      .TabStripPolicy = TabStripPolicyAlways
      .BackColorStyle = BackColorStyleUnderGrid
      .SelBackColor = &HC0C0C0
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
      .FontBold = False
      .TypeEditLen = 200
               
      .ShowScrollTips = ShowScrollTipsOff
      .TextTip = TextTipFloating
      
      Col = 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1160105, "Organisation"), TypeHAlignLeft, 20, vbButtonFace
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1160106, "Förs >5"), TypeHAlignRight, 7, Red
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1160107, "Förs <5"), TypeHAlignRight, 7, Red
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1160108, "Idag"), TypeHAlignRight, 7, Yellow
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1160109, "1 dag"), TypeHAlignRight, 7, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1160110, "2 dagar"), TypeHAlignRight, 7, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1160111, "3-4 dagar"), TypeHAlignRight, 7, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1160112, "5-10 dagar"), TypeHAlignRight, 7, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1160113, ">10 dagar"), TypeHAlignRight, 7, White
      Col = Col + 1: SetupColumn Spread, Nr, Col, Client.Texts.Txt(1160114, "Summa"), TypeHAlignRight, 8, vbButtonFace
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
      'ClearWorkBook lstStat
      GetDataNow
   End If
End Sub
Private Sub GetDataNow()

   Dim I As Integer
   Dim NumCol As Integer
   Dim W As Long
   Dim Stat As clsStat
   Dim Row As Integer
   Dim NumberOfRows As Integer
   Dim Col As Integer
   Dim RowHeader As String
   
   If CurrentOrgId <> NextOrgId Then
      CurrentOrgId = NextOrgId

      ClearWorkBook lstStat
      ClearTotals
      Row = 1
      Client.StatMgr.CreateList NextOrgId
      Do While Client.StatMgr.ListNextItem(Stat)
         NumberOfRows = NumberOfRows + 1
         RowHeader = Client.OrgMgr.TextFromId(Stat.OrgId)
         Col = 1
         SetCellString SheetNumber, Row, Col, RowHeader
         SetCellString SheetLenSec, Row, Col, RowHeader
         Col = Col + 1: SetCell Row, Col, Stat.Num1, Stat.LenSec1, True
         Col = Col + 1: SetCell Row, Col, Stat.Num2, Stat.LenSec2, True
         Col = Col + 1: SetCell Row, Col, Stat.Num3, Stat.LenSec3, True
         Col = Col + 1: SetCell Row, Col, Stat.Num4, Stat.LenSec4, True
         Col = Col + 1: SetCell Row, Col, Stat.Num5, Stat.LenSec5, True
         Col = Col + 1: SetCell Row, Col, Stat.Num6, Stat.LenSec6, True
         Col = Col + 1: SetCell Row, Col, Stat.Num7, Stat.LenSec7, True
         Col = Col + 1: SetCell Row, Col, Stat.Num8, Stat.LenSec8, True
         Col = Col + 1: SetCell Row, Col, RowTotal(SheetNumber, Row), RowTotal(SheetLenSec, Row), False
         Row = Row + 1
      Loop
      If NumberOfRows > 1 Then
         AddFooterRow
         NumberOfFooterRow = 1
      Else
         NumberOfFooterRow = 0
      End If
   End If
   Set Stat = Nothing
End Sub
Private Sub AddFooterRow()

   Dim Row As Integer
   Dim Col As Integer

   With lstStat
      Row = .MaxRows + 1
      SetCellString SheetNumber, Row, 1, Client.Texts.Txt(1160114, "Summa")
      SetCellString SheetLenSec, Row, 1, Client.Texts.Txt(1160114, "Summa")
      For Col = 2 To .MaxCols - 1
         SetCell Row, Col, ColTotal(SheetNumber, Col), ColTotal(SheetLenSec, Col), False
      Next Col
      SetCell Row, .MaxCols, GrandTotal(SheetNumber), GrandTotal(SheetLenSec), False
      SetupRow lstStat, SheetNumber, Row, vbButtonFace
      SetupRow lstStat, SheetLenSec, Row, vbButtonFace
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
'Private Function RowTotal(Row) As Long
'
'   Dim I As Integer
'   Dim Res As Long
'
'   For I = 0 To CellUsedCol
'      Res = Res + Cell(CurrentUnit, Row, I)
'   Next I
'   RowTotal = Res
'End Function
'Private Function ColTotal(Col) As Long
'
'   Dim I As Integer
'   Dim Res As Long
'
'   For I = 0 To CellUsedRow
'      Res = Res + Cell(CurrentUnit, I, Col)
'   Next I
'   RowTotal = Res
'End Function


Sub SetCellString(Sheet As Integer, Row As Integer, Col As Integer, Txt As String)

   With lstStat
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

   Dim S As String

   If Num > 0 Then
      S = CStr(Num)
   End If
   SetCellString Sheet, Row, Col, S
End Sub
Sub SetCellTime(Sheet As Integer, Row As Integer, Col As Integer, Sec As Long)

   Dim S As String

   If Sec > 0 Then
      S = FormatLength(Sec)
   End If
   SetCellString Sheet, Row, Col, S
End Sub

Public Function GetSetting() As String

   'Dim Pathname As String

   'ClearWorkBook lstStat
   'Pathname = CreateTempFileName("tmp")
   'lstStat.SaveToFile Pathname, False
   'GetSetting = ReadStringFromTempFile(Pathname)
   'KillFileIgnoreError Pathname
End Function
