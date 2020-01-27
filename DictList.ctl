VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.UserControl ucDictList 
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   ScaleHeight     =   4575
   ScaleWidth      =   9495
   Begin FPSpreadADO.fpSpread lstDict 
      Height          =   4575
      HelpContextID   =   1080000
      Left            =   -120
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
      SpreadDesigner  =   "DictList.ctx":0000
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   5
      Left            =   1920
      Picture         =   "DictList.ctx":0370
      Top             =   0
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   4
      Left            =   1440
      Picture         =   "DictList.ctx":0C3A
      Top             =   0
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgLater 
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   3
      Left            =   960
      Picture         =   "DictList.ctx":107C
      Top             =   0
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "DictList.ctx":1946
      Top             =   0
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   480
      Picture         =   "DictList.ctx":2210
      Top             =   0
      Width           =   480
      Visible         =   0   'False
   End
End
Attribute VB_Name = "ucDictList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click(DictId As Long)
Public Event DblClick(DictId As Long)
Public Event RightClick(DictId As Long)
Public Event ChangeNumberInList(TotalNumber As Long, NumberOfWarnings As Long, TotalLength As Long)

Dim SetNewSortOrder As Boolean
Dim NumberOfWarnings As Long
Dim TotalNumberInList As Long
Dim TotalLength As Long

Dim CurrentOrgId As Long
Dim CurrentTimeStamp As Double
Dim RowDictId(MaxNumberOfDictation) As Long

Private NewSearchFilter As Boolean
Private NewCurrPatientFilter As Boolean

Private Const ImgNrLater = 0
Private Const ImgNrNow = 1
Private Const ImgNrSoon = 2
Private Const ImgNrLocked = 3
Private Const ImgNrWarning = 4
Private Const ImgNrNote = 5

Const SS_SORT_ORDER_ASCENDING = 1

Const SS_BORDER_TYPE_NONE = 0
Const SS_BORDER_TYPE_LEFT = 1
Const SS_BORDER_TYPE_RIGHT = 2
Const SS_BORDER_TYPE_TOP = 4
Const SS_BORDER_TYPE_BOTTOM = 8
Const SS_BORDER_TYPE_OUTLINE = 16

Const SS_BORDER_STYLE_DEFAULT = 0
Const SS_BORDER_STYLE_SOLID = 1
Const SS_BORDER_STYLE_FINE_DOT = 13

Const SS_BDM_CURRENT_ROW = 4
Public Sub NewLanguage()

   Dim I As Integer
   
   For I = 0 To UserControl.Controls.Count - 1
      Client.Texts.ApplyToControl UserControl.Controls(I)
   Next I
End Sub
Public Sub ExportExcelFile(Fn As String)

   lstDict.ExportExcelBookEx Fn, "", ExcelSaveFlagNone
End Sub
Public Sub ExportToHtml(Fn As String)

   lstDict.ExportToHtml Fn, False, ""
End Sub
Public Sub ExportToXml(Fn As String)

   lstDict.ExportToXml Fn, "", "", ExportToXMLFormattedData, ""
End Sub
Public Sub ExportTextFile(Fn As String)

   lstDict.SaveTabFile Fn
End Sub
Public Sub ExportListToFile(DefFileName As String)

   Dim Fn As String
   Dim Ext As String
   
   If Len(DefFileName) = 0 Then
      DefFileName = Client.Texts.Txt(1000403, "Diktatlista")
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
Public Property Set SearchFilter(Flt As clsFilter)

   Set Client.DictMgr.SearchFilter = Flt
   NewSearchFilter = True
End Property
Public Property Set CurrPatientFilter(Flt As clsFilter)

   Set Client.DictMgr.CurrPatientFilter = Flt
   NewCurrPatientFilter = True
End Property

Private Sub lstDict_AfterUserSort(ByVal Col As Long)

   Dim Sortkeys As Variant
   Dim SortKeyOrder As Variant
   Static Desc As Boolean
   Dim WarnCol As Long
   
   Desc = lstDict.ColUserSortIndicator(Col) = ColUserSortIndicatorDescending
   lstDict.Col = Col
   If lstDict.ColID = "1" Then         'Flags och lock
      Col = lstDict.GetColFromID(6)    'Expiration date
   End If
   WarnCol = lstDict.GetColFromID("2")
   Sortkeys = Array(WarnCol, Col, 0)
   If Desc Then
      SortKeyOrder = Array(2, 2, 2)
   Else
      SortKeyOrder = Array(2, 1, 1)
   End If
   lstDict.Sort -1, -1, -1, -1, SortByRow, Sortkeys, SortKeyOrder
End Sub

Private Sub lstDict_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

   Dim PicCol As Integer
   
   PicCol = lstDict.GetColFromID("1")
   If PicCol >= Col1 And PicCol <= Col2 Then
      lstDict.ColWidth(PicCol) = 2
   Else
      PicCol = lstDict.GetColFromID("2")
      If PicCol >= Col1 And PicCol <= Col2 Then
         lstDict.ColWidth(PicCol) = 2
      Else
         PicCol = lstDict.GetColFromID("3")
         If PicCol >= Col1 And PicCol <= Col2 Then
            lstDict.ColWidth(PicCol) = 2
         End If
      End If
   End If
End Sub

Private Sub lstDict_DblClick(ByVal Col As Long, ByVal Row As Long)

   If Row > 0 Then
      RaiseEvent DblClick(CLng(lstDict.GetRowItemData(Row)))
   End If
End Sub

Private Sub lstDict_KeyPress(KeyAscii As Integer)

   Dim Fn As String

   Select Case KeyAscii
      Case 13
         RaiseEvent DblClick(CLng(lstDict.GetRowItemData(lstDict.ActiveRow)))
      Case KeyAsciiExportList
         If Client.SysSettings.ExportAllowMenu Then
            ExportListToFile ""
         End If
   End Select
End Sub

Private Sub lstDict_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

   If Row > 0 Then
      RaiseEvent RightClick(CLng(lstDict.GetRowItemData(Row)))
   End If
End Sub

Private Sub UserControl_Resize()

   lstDict.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub
Public Sub RestoreSettings(Settings As String, Ver As String)

   Dim TempFilePath As String
   Dim Ok As Boolean

   With lstDict
      Ok = False
      If Len(Settings) > 0 And Ver >= "1.20.0000" Then
         TempFilePath = WriteStringToTempFile(Settings)
         If .LoadFromFile(TempFilePath) Then
            Ok = True
         End If
         KillFileIgnoreError TempFilePath
      End If
      If Not Ok Then
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
         
         .Col = 0:        .ColID = CStr(.Col):   .ColWidth(.Col) = 5:   .CellType = CellTypeNumber:  .TypeNumberDecPlaces = 0
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 2
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 2
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 2
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 10
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 10
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 5
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 11
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 11
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 10
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .MaxCols = .Col
         SetNewSortOrder = True
      Else
         .Row = -1
         .Col = 0
         .CellType = CellTypeNumber
         .TypeNumberDecPlaces = 0
      End If
      
      SetCellValue 0, 0, Client.Texts.Txt(1080101, "Id")
      SetCellValue 0, 1, ""                                       'Expiration Flags och lock
      SetCellValue 0, 2, ""                                       'Warning
      SetCellValue 0, 3, ""                                       'note
      SetCellValue 0, 4, Client.Texts.Txt(1080102, "Personnr")
      SetCellValue 0, 5, Client.Texts.Txt(1080103, "Namn")
      SetCellValue 0, 6, Client.Texts.Txt(1080104, "Skriv senast")
      SetCellValue 0, 7, Client.Texts.Txt(1080113, "Prioritet")
      SetCellValue 0, 8, Client.Texts.Txt(1080105, "Organisation")
      SetCellValue 0, 9, Client.Texts.Txt(1080106, "Typ")
      SetCellValue 0, 10, Client.Texts.Txt(1080107, "Längd")
      SetCellValue 0, 11, Client.Texts.Txt(1080108, "Intalare")
      SetCellValue 0, 12, Client.Texts.Txt(1080109, "Intalat")
      SetCellValue 0, 13, Client.Texts.Txt(1080110, "Utskrivare")
      SetCellValue 0, 14, Client.Texts.Txt(1080115, "Utskrivet")
      SetCellValue 0, 15, Client.Texts.Txt(1080111, "Status")
      SetCellValue 0, 16, Client.Texts.Txt(1080112, "Används av")
   
      .RowHeadersShow = Client.SysSettings.DictListShowDictId
   End With
End Sub
Public Sub GetData(OrgId As Long)

   Dim I As Integer
   Dim NumCol As Integer
   Dim Dict As clsDict
   Dim Row As Integer
   Dim RowUpdated(MaxNumberOfDictation) As Boolean
   Dim ReSort As Boolean
   Dim PrevTimeStamp As Double
   Dim DictIdForSelectedRow As Long
   Dim TopRow As Long
   Dim LeftCol As Long
   Dim TooMany As Boolean
   Dim LastNumberOfWarnings As Long
   Dim LastTotalNumberInList As Long
   Dim LastTotalLength As Long
   Dim T As Variant
   Dim UpdateCurrentList As Boolean
   
   LastNumberOfWarnings = NumberOfWarnings
   LastTotalNumberInList = TotalNumberInList
   LastTotalLength = TotalLength
   
   UpdateCurrentList = CurrentOrgId = OrgId
   If OrgId = 30050 And NewSearchFilter Then
      UpdateCurrentList = False
   End If
   If OrgId = 30005 And NewCurrPatientFilter Then
      UpdateCurrentList = False
   End If
   
   If UpdateCurrentList Then
      ReSort = False
      
      TotalLength = 0
      DictIdForSelectedRow = lstDict.GetRowItemData(lstDict.SelModeIndex)
      TopRow = lstDict.TopRow
      LeftCol = lstDict.LeftCol
      PrevTimeStamp = CurrentTimeStamp
      
      RowDictIdCacheInit
      
      CurrentTimeStamp = Client.DictMgr.CreateList(OrgId, CurrentTimeStamp, TooMany)
      Do While Client.DictMgr.ListNextItem(Dict)
         TotalLength = TotalLength + Dict.SoundLength
         Row = FindRowFromRowDictIdCache(Dict.DictId)
         If Row <= 0 Then
            lstDict.MaxRows = lstDict.MaxRows + 1
            Row = lstDict.MaxRows
            UpdateRowInList Row, Dict
            ReSort = True
         Else
            If Dict.TimeStamp > PrevTimeStamp Then
               UpdateRowInList Row, Dict
               ReSort = True
            End If
         End If
         RowUpdated(Row) = True
      Loop
      
      For I = lstDict.MaxRows To 1 Step -1
         If Not RowUpdated(I) Then
            DeleteRowInList I
            ReSort = True
         End If
      Next I
      If ReSort Then
         lstDict.UserColAction = UserColActionSort
         lstDict.SelModeIndex = FindRowFromDictId(DictIdForSelectedRow)
         lstDict.TopRow = TopRow
         lstDict.LeftCol = LeftCol
      End If
   Else
      NewSearchFilter = False
      NewCurrPatientFilter = False
      CurrentOrgId = OrgId
      CurrentTimeStamp = 0
      
      'SetBusy
      TotalLength = 0
      CurrentTimeStamp = Client.DictMgr.CreateList(OrgId, CurrentTimeStamp, TooMany)
      
      lstDict.MaxRows = 0
      lstDict.ClearRange -1, -1, -1, -1, True
      
      NumberOfWarnings = 0
      Row = 1
      Do While Client.DictMgr.ListNextItem(Dict)
         TotalLength = TotalLength + Dict.SoundLength
         lstDict.MaxRows = Row
         UpdateRowInList Row, Dict
         Row = Row + 1
      Loop
      If SetNewSortOrder And lstDict.MaxRows > 0 Then
         SetNewSortOrder = False
         lstDict_AfterUserSort 5
      End If
      lstDict.UserColAction = UserColActionSort
      If TooMany Then
         MsgBox Client.Texts.Txt(1080114, "För många diktat funna. Avgränsa sökvillkor!"), vbOKOnly
      End If
   End If
   Set Dict = Nothing
   
   TotalNumberInList = lstDict.MaxRows
   If LastNumberOfWarnings <> NumberOfWarnings Or LastTotalNumberInList <> TotalNumberInList Or LastTotalLength <> TotalLength Then
      RaiseEvent ChangeNumberInList(TotalNumberInList, NumberOfWarnings, TotalLength)
   End If
End Sub
Private Function FindRowFromDictId(DictId As Long) As Long

   Dim I As Long
    
   For I = 1 To lstDict.MaxRows
      If CLng(lstDict.GetRowItemData(I)) = DictId Then
         FindRowFromDictId = I
         Exit For
      End If
   Next I
End Function
Private Sub DeleteRowInList(Row)

   lstDict.Row = Row
   lstDict.Col = lstDict.GetColFromID("2")
   If Len(lstDict.Text) > 0 Then
      NumberOfWarnings = NumberOfWarnings - 1
   End If
   lstDict.DeleteRows Row, 1
   lstDict.MaxRows = lstDict.MaxRows - 1
End Sub
Private Sub UpdateRowInList(Row As Integer, Dict As clsDict)

   Dim C As Integer
   Dim Mark As Integer
   Dim Warning As Integer
   Dim NotePic As Integer
   Dim PriorityText As String
   Dim Ddiff As Integer

   lstDict.SetRowItemData Row, CStr(Dict.DictId)
   
   lstDict.Row = Row
   lstDict.Col = -1
      
   Mark = 0
   If Len(Dict.LockedByUserShortName) > 0 Then
      Mark = ImgNrLocked
      lstDict.ForeColor = &H808080
      lstDict.FontItalic = True
   Else
      If Dict.StatusId < Transcribed Then
         Ddiff = DateDiff("d", Now, Dict.ExpiryDate)
         If Ddiff < 0 Then
            Mark = ImgNrNow
         ElseIf Ddiff = 0 Then
            Mark = ImgNrSoon
         Else
            Mark = ImgNrLater
         End If
      End If
      lstDict.ForeColor = 0
      lstDict.FontItalic = False
   End If
   
   Dim Priority As clsPriority
   Client.PriorityMgr.GetFromId Priority, Dict.PriorityId
   PriorityText = Priority.PriortyText
   If Priority.Colour > 0 Then
      lstDict.BackColor = Priority.Colour
   End If
   If Priority.Warning And Dict.StatusId < Transcribed Then
      NumberOfWarnings = NumberOfWarnings + 1
      Warning = ImgNrWarning
   Else
      Warning = 0
   End If
   Set Priority = Nothing
   
   If Len(Dict.Note) > 0 Then
      NotePic = ImgNrNote
   End If
   C = 0:     SetCellValue Row, C, Dict.DictId
   C = C + 1: SetCellPicture Row, C, Mark
   C = C + 1: SetCellPicture Row, C, Warning, CStr(Warning)
   C = C + 1: SetCellPicture Row, C, NotePic, Dict.Note
   C = C + 1: SetCellValue Row, C, Dict.Pat.PatIdFormatted
   C = C + 1: SetCellValue Row, C, Dict.Pat.PatName
   C = C + 1: SetCellValue Row, C, Format$(Dict.ExpiryDate, "ddddd")
   C = C + 1: SetCellValue Row, C, PriorityText
   C = C + 1: SetCellValue Row, C, Dict.OrgText
   C = C + 1: SetCellValue Row, C, Dict.DictTypeText
   C = C + 1: SetCellValue Row, C, FormatLength(Dict.SoundLength)
   C = C + 1: SetCellValue Row, C, Dict.AuthorShortName
   C = C + 1: SetCellValue Row, C, Format$(Dict.Created, "ddddd hh:nn")
   C = C + 1: SetCellValue Row, C, Dict.TranscriberShortName
   If Dict.TranscribedDate <> 0 Then
      C = C + 1: SetCellValue Row, C, Format$(Dict.TranscribedDate, "ddddd hh:nn")
   Else
      C = C + 1: SetCellValue Row, C, ""
   End If
   C = C + 1: SetCellValue Row, C, Dict.StatusText
   C = C + 1: SetCellValue Row, C, Dict.LockedByUserShortName
End Sub
Sub SetCellValue(Row As Integer, Col As Integer, Txt As String)

   lstDict.Row = Row
   lstDict.Col = lstDict.GetColFromID(CStr(Col))
   lstDict.Value = Txt
End Sub
Sub SetCellPicture(Row As Integer, Col As Integer, PicNr As Integer, Optional Text As String = "")

   lstDict.Row = Row
   lstDict.Col = lstDict.GetColFromID(CStr(Col))
   If PicNr > 0 Then
      lstDict.CellType = CellTypePicture
      lstDict.Text = Text
      lstDict.TypePictCenter = True
      lstDict.TypePictStretch = True
      lstDict.TypePictMaintainScale = True
      Set lstDict.TypePictPicture = imgIcon(PicNr)
   Else
      lstDict.CellType = CellTypeStaticText
      lstDict.Text = ""
   End If
End Sub

Private Function GetCellValue(Row As Integer, Col As Integer) As String

   lstDict.Row = Row
   lstDict.Col = lstDict.GetColFromID(CStr(Col))
   GetCellValue = lstDict.Text
End Function

Public Function GetSetting() As String

   Dim Pathname As String

   lstDict.MaxRows = 0
   lstDict.ClearRange -1, -1, -1, -1, True
   Pathname = CreateTempFileName("tmp")
   lstDict.SaveToFile Pathname, False
   GetSetting = ReadStringFromTempFile(Pathname)
   KillFileIgnoreError Pathname
End Function
Private Sub RowDictIdCacheInit()

   Dim Row As Integer
   Dim Idx As Integer
   Dim MaxRows As Integer
   
   MaxRows = lstDict.MaxRows
   For Row = 1 To MaxRows
      RowDictId(Row) = CLng(lstDict.GetRowItemData(Row))
   Next Row
   
   For Idx = MaxRows + 1 To UBound(RowDictId)
      RowDictId(Idx) = 0
   Next Idx

End Sub
Private Function FindRowFromRowDictIdCache(DictId As Long) As Long

   Dim Row As Long
    
   For Row = 1 To lstDict.MaxRows
      If RowDictId(Row) = DictId Then
         FindRowFromRowDictIdCache = Row
         Exit For
      End If
   Next Row
End Function

