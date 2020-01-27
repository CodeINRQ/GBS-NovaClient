VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TEXTS_MAX = 600

Private IdTable(TEXTS_MAX) As Long
Private TextTable(TEXTS_MAX) As String
Private ToolTipTable(TEXTS_MAX) As String
Private NumberOfTexts As Integer
Private mCtfpath As String

Public Property Let CTFPath(ByVal Path As String)

   mCtfpath = Path
End Property

Private Sub ChangeHelpFile(LanguageCode As String)

   Dim Fn As String

   Fn = App.Path & "\" & App.EXEName & LanguageCode & ".hlp"
   If Not FileExist(Fn) Then
      Fn = App.Path & "\" & App.EXEName & "EN.hlp"
   End If
   On Error Resume Next
   App.HelpFile = Fn
End Sub
Private Function FileExist(Pn As String) As Boolean

   On Error Resume Next
   FileExist = FileLen(Pn) > 0
End Function
Public Sub NewLanguage(Språkkod As String)

   Dim Fil As Integer
   Dim Pos1 As Integer
   Dim Pos2 As Integer
   Dim L As Integer
   Dim Id As String
   Dim s As String
   Dim FileName As String
   Dim Path As String


   ChangeHelpFile Språkkod
   Path = CreateTempPath()
   FileName = Path & "Texts." & Språkkod

   On Error Resume Next

   Kill FileName
   
   frmMain.MhZip.QUnZip = "-s1qa2ws3ed """ & mCtfpath & """ texts." & Språkkod & " """ & Path & """"

   On Error GoTo 0

   NumberOfTexts = 0
   Fil = FreeFile
   On Error Resume Next
   Open FileName For Input As Fil
   If Err = 0 Then
   
      Do While Not EOF(Fil)
         Line Input #Fil, s
         Pos1 = InStr(s, Chr$(9))
         If Pos1 > 1 And Left$(s, 1) <= "9" Then
            Pos2 = InStr(Pos1 + 1, s, Chr$(9))
            If Pos2 > 0 Then
               L = Pos2 - Pos1 - 1
            Else
               L = 1000
            End If
            If NumberOfTexts <= UBound(IdTable) Then
               On Error Resume Next
               IdTable(NumberOfTexts) = CLng(Left$(s, Pos1 - 1))
               If Err = 0 Then
                  TextTable(NumberOfTexts) = StringReplace(mId$(s, Pos1 + 1, L), Chr$(133), "...")
                  If Pos2 > 0 Then
                     ToolTipTable(NumberOfTexts) = StringReplace(mId$(s, Pos2 + 1), Chr$(133), "...")
                  End If
                  NumberOfTexts = NumberOfTexts + 1
               End If
               On Error GoTo 0
            End If
         End If
      Loop
   
      Close Fil
   End If
   On Error Resume Next
   Kill FileName
   On Error GoTo 0

   SortTextTable
   ApplyToAllForms
End Sub

Private Sub SortTextTable()

   Dim I As Integer
   Dim J As Integer
   Dim Id As Long
   Dim Txt As String

   For I = 0 To NumberOfTexts - 2
      For J = I + 1 To NumberOfTexts - 1
         If IdTable(I) > IdTable(J) Then
            Id = IdTable(I)
            IdTable(I) = IdTable(J)
            IdTable(J) = Id
            Txt = TextTable(I)
            TextTable(I) = TextTable(J)
            TextTable(J) = Txt
         End If
      Next J
   Next I
End Sub

Private Sub ApplyToAllForms()

   Dim I As Integer

   For I = 0 To Forms.Count - 1
      ApplyToOneForm Forms(I)
   Next I
End Sub

Public Sub ApplyToOneForm(F As Form)

   Dim I As Integer
   Dim C As Control
   Dim s As String

   On Error Resume Next
   F.Caption = Txt(CLng(F.Tag), F.Caption)
   On Error GoTo ApplyToOneForm_Err
   For I = 0 To F.Controls.Count - 1
      Set C = F.Controls(I)
      ApplyToControl C
      If Left$(C.Name, 2) = "uc" Then
         C.NewLanguage
      End If
WithoutTag:
   Next I
   Exit Sub

ApplyToOneForm_Err:
   Resume WithoutTag
End Sub
Public Sub ApplyToControl(C As Control)

   On Error Resume Next
   If Len(C.Tag) > 0 Then
       C.Caption = Txt(CLng(C.Tag), C.Caption)
       C.ToolTipText = ToolTip(CLng(C.Tag), C.ToolTipText)
    End If
    'Debug.Print C.Name
End Sub
Public Function Txt(Id As Long, ByVal Def As String) As String

   Dim I As Integer

   Do While IdTable(I) < Id And I < NumberOfTexts
      I = I + 1
   Loop
   If IdTable(I) = Id Then
      If Len(TextTable(I)) > 0 Then
         Txt = TextTable(I)
      Else
         Txt = Def
      End If
   Else
      Txt = Def
   End If
End Function
Public Function ToolTip(Id As Long, ByVal Def As String) As String

   Dim I As Integer

   Do While IdTable(I) < Id And I < NumberOfTexts
      I = I + 1
   Loop
   If IdTable(I) = Id Then
      If Len(ToolTipTable(I)) > 0 Then
         ToolTip = ToolTipTable(I)
      Else
         ToolTip = Def
      End If
   Else
      ToolTip = Def
   End If
End Function

