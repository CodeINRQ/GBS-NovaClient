VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.UserControl ucDictList 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   ScaleHeight     =   4560
   ScaleWidth      =   9435
   Begin FPSpreadADO.fpSpread lstDict 
      Height          =   4575
      HelpContextID   =   1080000
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
      SpreadDesigner  =   "DictList.ctx":0000
   End
   Begin VB.Image imgLater 
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgLocked 
      Height          =   480
      Left            =   4080
      Picture         =   "DictList.ctx":0370
      Top             =   2040
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgSoon 
      Height          =   480
      Left            =   0
      Picture         =   "DictList.ctx":0C3A
      Top             =   0
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Image imgUrgent 
      Height          =   480
      Left            =   1200
      Picture         =   "DictList.ctx":1504
      Top             =   1680
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

Dim loadingdata As Boolean
Dim zoomindex As Integer
Dim lastOrg As Long

Dim CurrentOrgId As Long
Dim CurrentTimeStamp As Double

Private NewSearchFilter As Boolean
Private NewCurrPatientFilter As Boolean

Private Const ImgNrLocked = 0
Private Const ImgNrNow = 1
Private Const ImgNrSoon = 2
Private Const ImgNrLater = 3

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
   
   Desc = lstDict.ColUserSortIndicator(Col) = ColUserSortIndicatorDescending
   lstDict.Col = Col
   If lstDict.ColID = "1" Then
      Col = lstDict.GetColFromID(4)
   End If
   Sortkeys = Array(Col, 0)
   If Desc Then
      SortKeyOrder = Array(2, 2)
   Else
      SortKeyOrder = Array(1, 1)
   End If
   lstDict.Sort -1, -1, -1, -1, SortByRow, Sortkeys, SortKeyOrder
End Sub


Private Sub lstDict_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

   Dim PicCol As Integer
   
   PicCol = lstDict.GetColFromID("1")
   If PicCol >= Col1 And PicCol <= Col2 Then
      lstDict.ColWidth(PicCol) = 2
   End If
End Sub

Private Sub lstDict_DblClick(ByVal Col As Long, ByVal Row As Long)

   If Row > 0 Then
      RaiseEvent DblClick(CLng(lstDict.GetRowItemData(Row)))
   End If
End Sub

Private Sub lstDict_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      RaiseEvent DblClick(CLng(lstDict.GetRowItemData(lstDict.ActiveRow)))
   End If
End Sub

Private Sub lstDict_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

   If Row > 0 Then
      RaiseEvent RightClick(CLng(lstDict.GetRowItemData(Row)))
   End If
End Sub

Private Sub UserControl_Initialize()

   'isgrouped = True
   loadingdata = False
End Sub

Private Sub UserControl_Resize()

   lstDict.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub
Public Sub RestoreSettings(Settings As String)

   Dim TempFilePath As String
   Dim Ok As Boolean

   With lstDict
      Ok = False
      If Len(Settings) > 0 Then
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
         
         .Col = 0:        .ColID = CStr(.Col):   .ColWidth(.Col) = 5
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 2
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 10
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 5
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 11
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 10
         .Col = .Col + 1: .ColID = CStr(.Col):   .ColWidth(.Col) = 8
         .MaxCols = .Col
      End If
      
      SetCellValue 0, 0, Client.Texts.Txt(1080101, "Id")
      SetCellValue 0, 1, ""
      SetCellValue 0, 2, Client.Texts.Txt(1080102, "Personnr")
      SetCellValue 0, 3, Client.Texts.Txt(1080103, "Namn")
      SetCellValue 0, 4, Client.Texts.Txt(1080104, "Skriv senast")
      SetCellValue 0, 5, Client.Texts.Txt(1080105, "Organisation")
      SetCellValue 0, 6, Client.Texts.Txt(1080106, "Typ")
      SetCellValue 0, 7, Client.Texts.Txt(1080107, "Längd")
      SetCellValue 0, 8, Client.Texts.Txt(1080108, "Intalare")
      SetCellValue 0, 9, Client.Texts.Txt(1080109, "Intalat")
      SetCellValue 0, 10, Client.Texts.Txt(1080110, "Utskrivare")
      SetCellValue 0, 11, Client.Texts.Txt(1080111, "Status")
      SetCellValue 0, 12, Client.Texts.Txt(1080112, "Används av")
   
      .RowHeadersShow = Client.SysSettings.ShowDictId
   End With
End Sub
Public Sub GetData(OrgId As Long)

   Dim I As Integer
   Dim NumCol As Integer
   Dim W As Long
   Dim Dict As clsDict
   Dim Row As Integer
   Dim RowUpdated(MaxNumberOfDictation) As Boolean
   Dim ReSort As Boolean
   Dim PrevTimeStamp As Double
   Dim DictIdForSelectedRow As Long
   
   loadingdata = True
   
   If CurrentOrgId = OrgId And Not NewSearchFilter And Not NewCurrPatientFilter Then
      'Uppdate
      'For I = LBound(RowUpdated) To UBound(RowUpdated)     'Already False as local variable
      '   RowUpdated(I) = False
      'Next I
      ReSort = False
      
      DictIdForSelectedRow = lstDict.GetRowItemData(lstDict.SelModeIndex)
      PrevTimeStamp = CurrentTimeStamp
      CurrentTimeStamp = Client.DictMgr.CreateList(OrgId, CurrentTimeStamp)
      Do While Client.DictMgr.ListNextItem(Dict)
         Row = FindRowFromDictId(Dict)
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
            lstDict.DeleteRows I, 1
            lstDict.MaxRows = lstDict.MaxRows - 1
            ReSort = True
         End If
      Next I
      If ReSort Then
         lstDict.UserColAction = UserColActionSort
         For I = 1 To lstDict.MaxRows
            If lstDict.GetRowItemData(I) = DictIdForSelectedRow Then
               lstDict.SelModeIndex = I
               Exit For
            End If
         Next I
      End If
   Else
      NewSearchFilter = False
      NewCurrPatientFilter = False
      CurrentOrgId = OrgId
      CurrentTimeStamp = 0
      
      'SetBusy
      CurrentTimeStamp = Client.DictMgr.CreateList(OrgId, CurrentTimeStamp)
      
      lstDict.MaxRows = 0
      lstDict.ClearRange -1, -1, -1, -1, True
      Row = 1
      Do While Client.DictMgr.ListNextItem(Dict)
         lstDict.MaxRows = Row
         UpdateRowInList Row, Dict
         Row = Row + 1
      Loop
      lstDict.UserColAction = UserColActionSort
      'SetNotBusy
   End If
   loadingdata = False
   Set Dict = Nothing
End Sub
Private Function FindRowFromDictId(Dict As clsDict) As Long

   Dim I As Long
    
   With lstDict
      For I = 1 To .MaxRows
         If CLng(.GetRowItemData(I)) = Dict.DictId Then
            FindRowFromDictId = I
            Exit For
         End If
      Next I
   End With
End Function

Private Sub UpdateRowInList(Row As Integer, Dict As clsDict)

   Dim C As Integer
   Dim Mark As Integer
   Dim Ddiff As Integer

   With lstDict
      .SetRowItemData Row, CStr(Dict.DictId)
      
      Mark = -1
      If Len(Dict.LockedByUserShortName) > 0 Then
         Mark = ImgNrLocked
         .Row = Row
         .Col = -1
         .ForeColor = &H808080
         .FontItalic = True
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
         .Row = Row
         .Col = -1
         .ForeColor = 0
         .FontItalic = False
      End If
      
      C = 0:     SetCellValue Row, C, Dict.DictId
      C = C + 1: SetCellPicture Row, C, Mark
      C = C + 1: SetCellValue Row, C, Dict.Pat.PatIdFormatted
      C = C + 1: SetCellValue Row, C, Dict.Pat.PatName
      C = C + 1: SetCellValue Row, C, Format$(Dict.ExpiryDate, "ddddd")
      C = C + 1: SetCellValue Row, C, Dict.OrgText
      C = C + 1: SetCellValue Row, C, Dict.DictTypeText
      C = C + 1: SetCellValue Row, C, FormatLength(Dict.SoundLength)
      C = C + 1: SetCellValue Row, C, Dict.AuthorShortName
      C = C + 1: SetCellValue Row, C, Format$(Dict.Created, "ddddd hh:nn")
      
'      If Dict.Changed <> 0 Then
'         C = C + 1: SetCellValue Row, C, Format$(Dict.Changed, "ddddd hh:nn")
'      Else
'         C = C + 1: SetCellValue Row, C, ""
'      End If
      
      C = C + 1: SetCellValue Row, C, Dict.TranscriberShortName
      C = C + 1: SetCellValue Row, C, Dict.StatusText
      C = C + 1: SetCellValue Row, C, Dict.LockedByUserShortName
   End With
End Sub
Sub SetCellValue(Row As Integer, Col As Integer, Txt As String)

   With lstDict
      .Row = Row
      .Col = .GetColFromID(CStr(Col))
      .Value = Txt
   End With
End Sub
Sub SetCellPicture(Row As Integer, Col As Integer, PicNr As Integer)

   'If PicNr >= 0 Then
      With lstDict
         .Row = Row
         .Col = .GetColFromID(CStr(Col))
         .CellType = CellTypePicture
         .TypePictCenter = True
         .TypePictStretch = True
         .TypePictMaintainScale = True
         '.TypePictPicture = .LoadPicture(App.Path & "\images\" & PicName & ".ico", PictureTypeICO)
         Select Case PicNr
            Case ImgNrLocked
               Set .TypePictPicture = imgLocked.Picture
            Case ImgNrNow
               Set .TypePictPicture = imgUrgent.Picture
            Case ImgNrSoon
               Set .TypePictPicture = imgSoon.Picture
            Case ImgNrLater
               Set .TypePictPicture = imgLater.Picture
         End Select
      End With
   'End If
End Sub

Private Function GetCellValue(Row As Integer, Col As Integer) As String

   With lstDict
      .Row = Row
      .Col = .GetColFromID(CStr(Col))
      GetCellValue = .Text
   End With
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
