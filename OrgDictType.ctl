VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.UserControl ucOrgDictType 
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   ScaleHeight     =   2085
   ScaleWidth      =   8220
   Begin VB.Frame fraOrgDictType 
      Caption         =   "Diktattyper"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Tag             =   "1410101"
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdSave 
         Caption         =   "Spara"
         Height          =   300
         Left            =   6000
         TabIndex        =   1
         Tag             =   "1410102"
         Top             =   240
         Width           =   2055
      End
      Begin FPSpreadADO.fpSpread lstOrgDictType 
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5775
         _Version        =   458752
         _ExtentX        =   10186
         _ExtentY        =   2990
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
         SpreadDesigner  =   "OrgDictType.ctx":0000
      End
   End
End
Attribute VB_Name = "ucOrgDictType"
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

Public Sub Init()

   Dim I As Integer
   Dim LstIdx As Integer
   Dim ODt As clsOrgDictType
   Dim Row As Integer
   
   lstOrgDictType.MaxRows = 1
   lstOrgDictType.ClearRange -1, -1, -1, -1, True
   RestoreSettings
   
End Sub
Sub SetCellValue(Row As Integer, Col As Integer, Txt As String)

   With lstOrgDictType
      .Row = Row
      .Col = .GetColFromID(CStr(Col))
      .Value = Txt
   End With
End Sub
Sub SetCellBool(Row As Integer, Col As Integer, BoolVaue As Boolean)

   'If PicNr >= 0 Then
      With lstOrgDictType
         .Row = Row
         .Col = .GetColFromID(CStr(Col))
         .CellType = CellTypeCheckBox
         .TypeCheckType = TypeCheckTypeNormal
         .TypeVAlign = TypeVAlignCenter
         .TypeHAlign = TypeHAlignCenter
         .Value = BoolVaue
      End With
   'End If
End Sub
Private Sub RestoreSettings()

   With lstOrgDictType
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
      SetCellValue 0, 1, Client.Texts.Txt(1410103, "Diktattyp")
      SetCellValue 0, 2, Client.Texts.Txt(1410104, "Används")
      SetCellValue 0, 3, Client.Texts.Txt(1410105, "Förinst")
   
      .RowHeadersShow = False
   End With
End Sub
Private Sub ShowDictTypeForOrg()

   Static LastOrgId As Long
   Dim OId As Long
   Dim Row As Integer
   Dim DictType As clsDictType
   Dim OrgDictType As clsOrgDictType
   Dim DTIdx As Integer
   Dim DictTypeEnabled As Boolean
   Dim DictTypeDefault As Boolean
   
   Dim Org As clsOrg
   
   cmdSave.Enabled = False
   Set Org = Nothing
   Client.OrgMgr.GetOrgFromId Org, CurrOrgId
   lstOrgDictType.ClearRange -1, -1, -1, -1, True
   
   If CurrOrgId < 30000 Then
      If Not Org Is Nothing Then
         Client.DictTypeMgr.Init
         
         lstOrgDictType.MaxRows = 0
         Row = 1
         
         For DTIdx = 0 To Client.DictTypeMgr.Count - 1
            Client.DictTypeMgr.GetFromIndex DictType, DTIdx
            Client.DictTypeMgr.GetOrgDictTypeFromId OrgDictType, CurrOrgId, DictType.DictTypeId
            If Not OrgDictType Is Nothing Then
               DictTypeEnabled = True
               DictTypeDefault = OrgDictType.Def
            Else
               DictTypeEnabled = False
               DictTypeDefault = False
            End If
            lstOrgDictType.MaxRows = Row
            UpdateRowInList Row, DictType.DictTypeId, DictType.DictTypeText, DictTypeEnabled, DictTypeDefault
            Row = Row + 1
         Next DTIdx
      End If
   End If
End Sub
Private Sub UpdateRowInList(Row As Integer, DictTypeId As Integer, DictTypeTxt As String, DictTypeEnabled As Boolean, DictTypeDefault As Boolean)

   Dim C As Integer

   With lstOrgDictType
      .SetRowItemData Row, CStr(DictTypeId)
      .Row = Row

      C = 0:     SetCellValue Row, C, CStr(DictTypeId)
      C = C + 1: SetCellValue Row, C, DictTypeTxt
      C = C + 1: SetCellBool Row, C, DictTypeEnabled
      C = C + 1: SetCellBool Row, C, DictTypeDefault
   End With
End Sub

Public Sub NewOrg(OrgId As Long)

   If CurrOrgId <> OrgId Then
      CurrOrgId = OrgId
      ShowDictTypeForOrg
   End If
End Sub

Private Sub cmdSave_Click()

   Dim R As Integer
   Dim E As Boolean
   Dim D As Boolean

   cmdSave.Enabled = False
   With lstOrgDictType
      Client.DictTypeMgr.DeleteOrgDictTypeByOrgId CurrOrgId
      
      For R = 1 To .MaxRows
         .Row = R
         .Col = 1: Debug.Print .Value
         .Col = 2: E = .Value
         .Col = 3: D = .Value
      
         If E Then
            Client.DictTypeMgr.SaveOrgDictType CurrOrgId, CInt(.GetRowItemData(R)), D
         End If
      Next R
   End With
   Client.DictTypeMgr.Init
   ShowDictTypeForOrg
End Sub

Private Sub lstOrgDictType_EditChange(ByVal Col As Long, ByVal Row As Long)

   cmdSave.Enabled = True
End Sub

Private Sub lstOrgDictType_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   cmdSave.Enabled = True
End Sub
