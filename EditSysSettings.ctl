VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.UserControl ucEditSysSettings 
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   ScaleHeight     =   2610
   ScaleWidth      =   8265
   Begin VB.Frame fraSettings 
      Caption         =   "Inställningar"
      Height          =   2535
      HelpContextID   =   1120000
      Left            =   0
      TabIndex        =   0
      Tag             =   "1120101"
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Återställ"
         Height          =   300
         Left            =   6000
         TabIndex        =   3
         Tag             =   "1120103"
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Spara"
         Height          =   300
         Left            =   6000
         TabIndex        =   2
         Tag             =   "1120102"
         Top             =   240
         Width           =   2055
      End
      Begin FPSpreadADO.fpSpread lstSettings 
         Height          =   2160
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5775
         _Version        =   458752
         _ExtentX        =   10186
         _ExtentY        =   3810
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "EditSysSettings.ctx":0000
      End
   End
End
Attribute VB_Name = "ucEditSysSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const White = &HFFFFFF

Private mSettings As clsStringStore

Public Event SaveClicked(Settings As clsStringStore)
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub

Public Property Set Settings(SS As clsStringStore)

   Set mSettings = SS
   FillSettingList lstSettings
End Property
Private Sub FillSettingList(Spread As fpSpread)

   Dim Se As String
   Dim Ke As String
   Dim Va As String
   Dim Row As Integer
   Dim Col As Integer
   
   
   SetupSpread lstSettings
   mSettings.Filter = ""
   Col = 1: SetupColumn Spread, 0, Col, "Section", TypeHAlignLeft, 15, White
   Col = Col + 1: SetupColumn Spread, 0, Col, "Key", TypeHAlignLeft, 15, White
   Col = Col + 1: SetupColumn Spread, 0, Col, "Value", TypeHAlignLeft, 15, White
   
   With Spread
      Row = 1
      Do While mSettings.GetNextFromFilter(Se, Ke, Va)
         Col = 1:       SetCellString Spread, 0, Row, Col, Se
         Col = Col + 1: SetCellString Spread, 0, Row, Col, Ke
         Col = Col + 1: SetCellString Spread, 0, Row, Col, Va
         Row = Row + 1
      Loop
      SortBySectionAndKey
   End With
End Sub
Private Sub SortBySectionAndKey()

   Dim Sortkeys As Variant
   Dim SortKeyOrder As Variant
   
   Sortkeys = Array(1, 2)
   SortKeyOrder = Array(1, 1)
   lstSettings.Sort -1, -1, -1, -1, SortByRow, Sortkeys, SortKeyOrder
End Sub

Sub SetCellString(Spread As fpSpread, Sheet As Integer, Row As Integer, Col As Integer, Txt As String)

   With Spread
      .Sheet = Sheet
      If .MaxRows < Row Then
         .MaxRows = Row
      End If
      .Row = Row
      .Col = .GetColFromID(CStr(Col))
      .CellType = CellTypeEdit
      .Value = Txt
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
Private Sub SetupSpread(Spread As fpSpread)

   With Spread
      .Reset
      .SheetCount = 1
      .MaxCols = 3
      .TabStripPolicy = TabStripPolicyNever
      .BackColorStyle = BackColorStyleUnderGrid
      .RowHeadersShow = False
      .OperationMode = OperationModeNormal
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
   End With
End Sub

Private Sub cmdRestore_Click()

   FillSettingList lstSettings
End Sub

Private Sub cmdSave_Click()

   Dim I As Integer
   Dim Se As String
   Dim Ke As String
   Dim Va As String

   Set mSettings = New clsStringStore
   For I = 1 To lstSettings.MaxRows
       Se = GetCellString(lstSettings, 0, I, 1)
       Ke = GetCellString(lstSettings, 0, I, 2)
       Va = GetCellString(lstSettings, 0, I, 3)
       If Len(Se) > 0 And Len(Ke) > 0 Then
          mSettings.AddString Se, Ke, Va
       End If
   Next I
   RaiseEvent SaveClicked(mSettings)
End Sub
Private Function GetCellString(Spread As fpSpread, Sheet As Integer, Row As Integer, Col As Integer) As String

   With Spread
      .Sheet = Sheet
      .Row = Row
      .Col = .GetColFromID(CStr(Col))
      GetCellString = .Value
   End With
End Function

