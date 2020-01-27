Attribute VB_Name = "modCareTalkClient"
Option Explicit

Public Const API_ACCESS_CODE = "dsfkkd8jd,.,sdf88h3%&%&¤iyt"

Private Const SYNCHRONIZE = &H100000
Private Const WAIT_OBJECT_0 = 0
Private Const WAIT_TIMEOUT = &H102

Private Const STATUS_TIMEOUT = &H102&
Private Const INFINITE = -1& ' Infinite interval
Private Const QS_KEY = &H1&
Private Const QS_MOUSEMOVE = &H2&
Private Const QS_MOUSEBUTTON = &H4&
Private Const QS_POSTMESSAGE = &H8&
Private Const QS_TIMER = &H10&
Private Const QS_PAINT = &H20&
Private Const QS_SENDMESSAGE = &H40&
Private Const QS_HOTKEY = &H80&
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT _
        Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON _
        Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Declare Function MsgWaitForMultipleObjects Lib "user32" _
        (ByVal nCount As Long, pHandles As Long, _
        ByVal fWaitAll As Long, ByVal dwMilliseconds _
        As Long, ByVal dwWakeMask As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) _
   As Long
Public Const GW_HWNDPREV = 3

Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long

Declare Function GetWindow Lib "user32" _
  (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Declare Function SetForegroundWindow Lib "user32" _
  (ByVal hWnd As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
                        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const MAX_BUFFER_LENGTH = 256

Private Declare Function GetDriveType Lib "kernel32" _
Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Client As clsClient

Private Declare Function OpenProcess Lib "kernel32" ( _
   ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
   ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
   ByVal hObject As Long) As Long

Public Enum StatusEnum
   BeingRecorded = 20
   Recorded = 30
   BeingTrancribed = 50
   WaitForSign = 60
   Transcribed = 70
   SoundDeleted = 80
End Enum

Public Enum ClientTypeEnum
   ClientType_GrundigNova = 0
   ClientType_CareTalk = 1
   ClientType_LegalTalk = 2
End Enum

Public ApplicationVersion As String
Public GlobalCommandLine As String
Public GlobalAutostart As Boolean
Public ReadyForApiCalls As Boolean
Public RecorderInUse As Boolean

Public StartUpServer As String
Public StartUpDatabase As String
Public StartUpUserLoginName As String
Public StartUpPassword As String
Public StartUpExtPassword As String
Public StartUpExtSystem As String
Public StartUpLoginResult As Integer
Public StartUpFormMainIsLoaded As Integer    '0=no, 1=being, 2=yes

Public LastfrmDictTop As Long
Public LastfrmDictLeft As Long

Public Const MaxNumberOfDictation = 30000

Public Const TraceTitle_Entry = "Entry"
Public Const TraceTitle_Exit = "Exit"
Public Const TraceTitle_Err = "Error"
Public Const TraceTitle_Event = "Event"

Public Const KeyAsciiExportList = 5

Public Function GetDigtaDSSFolder() As String

   Dim DigtaDrive As String
   Dim s As String
    
   DigtaDrive = ""   'First try
   Do While GetNextPossibleDrive(DigtaDrive)
      s = ""
      On Error Resume Next
      s = Dir(DigtaDrive & ":\DSS", vbDirectory)
      On Error GoTo 0
      If s = "DSS" Then
         GetDigtaDSSFolder = DigtaDrive & ":\DSS\"
         Exit Function
      End If
   Loop
End Function
Private Function GetNextPossibleDrive(ByRef DriveLetter As String) As Boolean

   Dim DriveType As Long
   Dim s As String
   Dim Pos As Integer
   
   s = Client.SysSettings.ImportDSSDrives
   If Len(s) = 0 Then
      If Len(DriveLetter) = 0 Then
         DriveLetter = "C"    'skip A and B, they are probably diskette drives, I think...
      Else
         DriveLetter = Chr$(Asc(DriveLetter) + 1)
      End If
      Do While DriveLetter <= "Z"
         DriveType = GetDriveType(DriveLetter & ":\")
         If DriveType = 2 Then
            GetNextPossibleDrive = True
            Exit Function
         End If
         DriveLetter = Chr$(Asc(DriveLetter) + 1)
      Loop
      GetNextPossibleDrive = False
   Else
      If Len(DriveLetter) = 0 Then
         Pos = 1
      Else
         Pos = InStr(s, DriveLetter)
         Pos = Pos + 1
      End If
      If Pos > Len(s) Then
         GetNextPossibleDrive = False
      Else
         DriveLetter = mId$(s, Pos, 1)
         GetNextPossibleDrive = True
      End If
   End If
End Function
Public Sub Trc(Loc As String, Value As String)

   'Debug.Print Loc & ": " & Value
End Sub

Function nvl(Value As Variant, InsteadOfNull As Variant) As Variant

   If IsNull(Value) Then
      nvl = InsteadOfNull
   Else
      nvl = Value
   End If
End Function
Public Function CreateTempFileName(ExtensionExclDot As String) As String

   Dim s As String
   
   s = CreateTempPath() & CStr(CLng(Timer)) & CStr(RndNumber(1, 30000)) & "." & ExtensionExclDot
   KillFileIgnoreError s
   CreateTempFileName = s
End Function
Private Function RndNumber(Min As Integer, Max As Integer) As Integer

   RndNumber = Int(Rnd * (Max - Min + 1)) + Min
End Function

Public Function CreateTempFolder(FolderName As String) As String

   Dim P As String
   
   P = CreateTempPath() & FolderName
   On Error Resume Next
   MkDir P
   CreateTempFolder = P
End Function

Public Function CreateTempPath() As String

    Dim strBufferString As String
    Dim lngResult As Long
    strBufferString = String(MAX_BUFFER_LENGTH, "X")
    lngResult = GetTempPath(MAX_BUFFER_LENGTH, strBufferString)
    CreateTempPath = mId(strBufferString, 1, lngResult)
End Function

Public Sub KillFileIgnoreError(FileName As String)

   On Error Resume Next
   Kill FileName
End Sub
Public Function StringReplace(ByVal str As String, SubStrToReplace As String, InsertInstead As String) As String

   Dim Pos As Integer
   
   Pos = InStr(str, SubStrToReplace)
   Do While Pos > 0
      str = Left$(str, Pos - 1) & InsertInstead & mId$(str, Pos + Len(SubStrToReplace))
      'Pos = 0
      On Error Resume Next
      Pos = InStr(Pos + Len(InsertInstead), str, SubStrToReplace)
      On Error GoTo 0
   Loop
   StringReplace = str
End Function
Public Function CheckPatId(ByVal PatId As String) As Boolean

   'Returns true for a Correct number, else false
   Dim Siffra(9) As Integer
   Dim Resultat As Integer
   Dim I As Integer
   Dim Century As String
   
   If Not Client.SysSettings.DictInfoMandatoryPatId And Len(PatId) = 0 Then
      CheckPatId = True
      Exit Function
   End If
   
   If Client.SysSettings.DictInfoAlfaInPatid Then
      If Client.SysSettings.DictInfoMandatoryPatId Then
         CheckPatId = Len(PatId) > 0
         Exit Function
      Else
         CheckPatId = True
         Exit Function
      End If
   End If
   
   'Remove "-" if there is one
   PatId = FormatPatIdForStoring(PatId)
   
   If Client.SysSettings.DictInfoMandatoryPatIdCentury Then
      'check length
      If Len(PatId) <> 12 Then
         CheckPatId = False
         Exit Function
      End If
   End If
   
   If Len(PatId) = 12 Then
      Century = Left$(PatId, 2)
      PatId = mId$(PatId, 3)
      If (Century <> "19" And Century <> "20") Then
         CheckPatId = False
         Exit Function
      End If
   Else
      If Len(PatId) <> 10 Then
         CheckPatId = False
         Exit Function
      End If
   End If
   
   If Not Client.SysSettings.DictInfoPatIdChecksum Then
      CheckPatId = True
      Exit Function
   End If
   
   'split in strings
   For I = 1 To 9
      Siffra(I) = CInt(mId(PatId, I, 1))
   Next
   
   'double number in odd positions
   For I = 0 To 9 Step 2
      Siffra(I + 1) = Siffra(I + 1) * 2
   Next
   
   'add to digits strings and add strings
   For I = 1 To 9
      If Siffra(I) >= 10 Then
         Resultat = Resultat + Siffra(I) - 9
      Else
         Resultat = Resultat + Siffra(I)
      End If
   Next
    
   I = CInt(mId(PatId, Len(PatId), 1))

   If (10 - Resultat Mod 10) Mod 10 = I Then
      CheckPatId = True
   Else
      CheckPatId = False
   End If
End Function
Public Function CheckPatname(PName As String) As Boolean

   If Client.SysSettings.DictInfoMandatoryPatName Then
      CheckPatname = Len(PName) > 0
   Else
      CheckPatname = True
   End If
End Function
Public Function FormatLength(Sec As Long) As String

   Dim Mins As Integer
   Dim Hours As Integer
   Dim Secs As Integer
   Dim s As String

   Secs = Sec Mod 60
   Mins = (Sec \ 60) Mod 60
   Hours = (Sec \ 60) \ 60
   If Hours <> 0 Then
      s = Format$(Hours, "0") & ":"
   End If
   FormatLength = s & Format$(Mins, "00") & ":" & Format$(Secs, "00")
End Function
Public Function WriteStringToTempFile(s As String) As String

   Dim Pathname As String
   
   Pathname = CreateTempFileName("tmp")
   WriteStringToFile s, Pathname
   WriteStringToTempFile = Pathname
End Function
Public Sub WriteStringToFile(s As String, Pathname As String, Optional Append As Boolean)

   Dim F As Integer
   
   F = FreeFile
   Open Pathname For Binary Access Write As #F
   If Append Then
      Seek #F, LOF(F) + 1
   End If
   Put #F, , s
   Close #F
End Sub
Public Function ReadStringFromTempFile(Pathname As String, Optional MaxLength As Long = 0, Optional Offset As Long = 0) As String

   Dim F As Integer
   Dim s As String
   
   If MaxLength <= 0 Then
      MaxLength = FileLen(Pathname)
   End If
   s = Space$(MaxLength)
   F = FreeFile
   Open Pathname For Binary Access Read As #F
   If Offset > 0 Then
      Get #F, Offset, s
   Else
      Get #F, , s
   End If
   Close #F
   ReadStringFromTempFile = s
End Function


Private Function CommandString(ByRef CommandLine As String) As String

   Dim Pos As Integer

   CommandLine = Trim$(CommandLine)
   Pos = InStr(CommandLine, " ")
   If Pos > 0 Then
      CommandString = Left$(CommandLine, Pos - 1)
      CommandLine = mId$(CommandLine, Pos + 1)
   Else
      CommandString = CommandLine
      CommandLine = ""
   End If
End Function

Public Function CommandValue(KeyWithoutSlash As String, Default As String) As String

   Dim CommandLine As String
   Dim s As String

   CommandLine = GlobalCommandLine
   s = CommandString(CommandLine)
   Do While s <> ""
      If UCase$(s) = "/" & UCase$(KeyWithoutSlash) Then
         CommandValue = CommandString(CommandLine)
         Exit Function
      End If
      s = CommandString(CommandLine)
   Loop
   CommandValue = Default
End Function
Sub GotoPrevInstance()

   Dim OldTitle As String
   Dim PrevHndl As Long
   Dim result As Long

   On Error Resume Next
   'Save the title of the application.
   OldTitle = App.Title

   'Rename the title of this application so FindWindow
   'will not find this application instance.
   App.Title = "unwanted instance1"

   'Attempt to get window handle using VB4 class name.
   PrevHndl = FindWindow("ThunderRTMain", OldTitle)
   App.Title = "unwanted instance2"

   'Check for no success.
   If PrevHndl = 0 Then
      'Attempt to get window handle using VB5 class name.
      PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
   End If
   App.Title = "unwanted instance3"

   'Check if found
   If PrevHndl = 0 Then
      'Attempt to get window handle using VB6 class name
      PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
   End If
   App.Title = "unwanted instance4"

   'Check if found
   If PrevHndl = 0 Then
      'No previous instance found.
      Exit Sub
   End If
   App.Title = "unwanted instance5"

   'Get handle to previous window.
   PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
   App.Title = "unwanted instance6"

   'Restore the program.
   result = OpenIcon(PrevHndl)
   App.Title = "unwanted instance7"

   'Activate the application.
   result = SetForegroundWindow(PrevHndl)
   App.Title = "unwanted instance8"

   'End the application.
   Unload frmMain
   App.Title = "unwanted instance9"
   End
   App.Title = "unwanted instance10"
End Sub
Sub Main()
   
   If App.StartMode = 0 Then
      Load frmMain
   End If
End Sub
Public Function GetStationName() As String
   
   Dim s As String
   
   s = Space(512)
   GetComputerName s, Len(s)
   GetStationName = Environ("USERDOMAIN") & "\" & Trim$(s)
End Function

Public Function MsgWaitObj(Interval As Long, _
                           Optional hObj As Long = 0&, _
                           Optional nObj As Long = 0&) As Long
                           
   Dim T As Long, T1 As Long
   
   If Interval <> INFINITE Then
       T = MyGetTickCount()
       On Error Resume Next
       T = T + Interval
       ' Overflow prevention
       If Err <> 0& Then
           If T > 0& Then
               T = ((T + &H80000000) _
               + Interval) + &H80000000
           Else
               T = ((T - &H80000000) _
               + Interval) - &H80000000
           End If
       End If
       On Error GoTo 0
       ' T contains now absolute time of the end of interval
   Else
       T1 = INFINITE
   End If
   Do
       If Interval <> INFINITE Then
           T1 = MyGetTickCount()
           On Error Resume Next
        T1 = T - T1
           ' Overflow prevention
           If Err <> 0& Then
               If T > 0& Then
                   T1 = ((T + &H80000000) _
                   - (T1 - &H80000000))
               Else
                   T1 = ((T - &H80000000) _
                   - (T1 + &H80000000))
               End If
           End If
           On Error GoTo 0
           ' T1 contains now the remaining interval part
           If IIf((T1 Xor Interval) > 0&, _
               T1 > Interval, T1 < 0&) Then
               ' Interval expired
               ' during DoEvents
               MsgWaitObj = STATUS_TIMEOUT
               Exit Function
           End If
       End If
       ' Wait for event, interval expiration
       ' or message appearance in thread queue
       MsgWaitObj = MsgWaitForMultipleObjects(nObj, _
               hObj, 0&, T1, QS_ALLINPUT)
       ' Let's message be processed
       DoEvents
       If MsgWaitObj <> nObj Then Exit Function
       ' It was message - continue to wait
   Loop
End Function
Public Function IsLoaded(FormName As String) As Boolean

   Dim sFormName As String
   Dim F As Form
   
   sFormName = UCase$(FormName)
   
   For Each F In Forms
      If UCase$(F.Name) = sFormName Then
        IsLoaded = True
        Exit Function
      End If
   Next
End Function
Sub SelectAllText(C As Control)

   On Error Resume Next
   C.Selstart = 0
   C.SelLength = Len(C.Text)
End Sub

Public Function GetExportFileName(DefFileName As String) As String

   Dim Filter As String
   Dim Pos As Integer
   
   Filter = Client.Texts.Txt(1000901, "Excel-filer") & " (*.xls)|*.xls|"
   Filter = Filter & Client.Texts.Txt(1000905, "Text-filer") & " (*.txt)|*.txt|"
   Filter = Filter & Client.Texts.Txt(1000903, "Html-filer") & " (*.htm)|*.htm|"
   Filter = Filter & Client.Texts.Txt(1000904, "Xml-filer") & " (*.xml)|*.xml|"
   Filter = Filter & Client.Texts.Txt(1000902, "Alla filer") & " (*.*)|*.*"
   
   frmMain.CDialog.FileName = DefFileName
   frmMain.CDialog.InitDir = ""
   frmMain.CDialog.CancelError = True
   frmMain.CDialog.DefaultExt = "xls"
   frmMain.CDialog.DialogTitle = Client.Texts.Txt(1000900, "Exportera")
   frmMain.CDialog.Filter = Filter
   frmMain.CDialog.FilterIndex = 1
   frmMain.CDialog.Flags = cdlOFNExplorer Or cdlOFNOverwritePrompt
   frmMain.CDialog.HelpFile = ""
   frmMain.CDialog.HelpCommand = 0
   frmMain.CDialog.HelpContext = 0
   On Error Resume Next
   frmMain.CDialog.Action = 2
   If Err <> 0 Then
      Exit Function
   End If
   On Error GoTo 0

   GetExportFileName = frmMain.CDialog.FileName
End Function
Public Function TryToCopyFile(Source As String, Dest As String) As Boolean

   On Error GoTo TryToCopyFile_Err
   FileCopy Source, Dest
   TryToCopyFile = True
   Exit Function
   
TryToCopyFile_Err:
   TryToCopyFile = False
   Exit Function
End Function
Public Function FormatPatIdForStoring(ByVal s As String) As String

   s = StringReplace(s, "-", "")
   s = StringReplace(s, "/", "")
   s = StringReplace(s, "\", "")
   s = StringReplace(s, ".", "")
   s = StringReplace(s, ",", "")
   s = StringReplace(s, "+", "")
   FormatPatIdForStoring = s
End Function
Public Function RC4(ByVal Expression As String, ByVal Password As String) As String
   On Error Resume Next
   Dim RB(0 To 255) As Integer, x As Long, y As Long, Z As Long, Key() As Byte, ByteArray() As Byte, Temp As Byte
   If Len(Password) = 0 Then
       Exit Function
   End If
   If Len(Expression) = 0 Then
       Exit Function
   End If
   If Len(Password) > 256 Then
       Key() = StrConv(Left$(Password, 256), vbFromUnicode)
   Else
       Key() = StrConv(Password, vbFromUnicode)
   End If
   For x = 0 To 255
       RB(x) = x
   Next x
   x = 0
   y = 0
   Z = 0
   For x = 0 To 255
       y = (y + RB(x) + Key(x Mod Len(Password))) Mod 256
       Temp = RB(x)
       RB(x) = RB(y)
       RB(y) = Temp
   Next x
   x = 0
   y = 0
   Z = 0
   ByteArray() = StrConv(Expression, vbFromUnicode)
   For x = 0 To Len(Expression)
       y = (y + 1) Mod 256
       Z = (Z + RB(y)) Mod 256
       Temp = RB(y)
       RB(y) = RB(Z)
       RB(Z) = Temp
       ByteArray(x) = ByteArray(x) Xor (RB((RB(y) + RB(Z)) Mod 256))
   Next x
   RC4 = StrConv(ByteArray, vbUnicode)
End Function

Public Sub ShellAndWait(ByVal program_name As String, ByVal window_style As VbAppWinStyle)

   Dim process_id As Long
   Dim process_handle As Long

   ' Start the program.
   process_id = Shell(program_name, window_style)

   ' Wait for the program to finish.
   ' Get the process handle.
   process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
   If process_handle <> 0 Then
      WaitForSingleObject process_handle, INFINITE
      CloseHandle process_handle
   End If
End Sub

Public Function FileExists(ByVal sFileName As String) As Boolean

   FileExists = Len(Dir(sFileName)) > 0
End Function
Public Function MyGetTickCount() As Long
   
   Dim Count As Currency
   Static Offset As Currency
   
   On Error GoTo Fel
   Count = GetTickCount()
   If Offset = 0 Then
      Offset = -Count + 1
   End If
   MyGetTickCount = CLng(Count + Offset)
   Exit Function
   
Fel:
   MsgBox "MyGetTickCount " & CStr(Count) & " " & CStr(Offset)
   Exit Function
End Function
