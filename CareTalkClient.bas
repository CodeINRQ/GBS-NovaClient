Attribute VB_Name = "modCareTalkClient"
Option Explicit

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

Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long

Declare Function GetWindow Lib "user32" _
  (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Declare Function SetForegroundWindow Lib "user32" _
  (ByVal hwnd As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
                        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const MAX_BUFFER_LENGTH = 256

Private Declare Function GetDriveType Lib "kernel32" _
Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Client As clsClient

Public Enum StatusEnum
   BeingRecorded = 20
   Recorded = 30
   BeingTrancribed = 50
   WaitForSign = 60
   Transcribed = 70
   SoundDeleted = 80
End Enum

Public ApplicationVersion As String
Public GlobalCommandLine As String
Public ReadyForApiCalls As Boolean

Public StartUpServer As String
Public StartUpDatabase As String
Public StartUpUserLoginName As String
Public StartUpPassword As String
Public StartUpExtPassword As String
Public StartUpExtSystem As String
Public StartUpLoginResult As Integer
Public StartUpFormMainIsLoaded As Integer    '0=no, 1=being, 2=yes

Public Const MaxNumberOfDictation = 30000

Public Const TraceTitle_Entry = "Entry"
Public Const TraceTitle_Exit = "Exit"
Public Const TraceTitle_Err = "Error"
Public Const TraceTitle_Event = "Event"

Public Function GetDigtaDSSFolder() As String

   Dim DigtaDrive As String
   Dim S As String
    
   DigtaDrive = "C"    'skip A and B, they are probably diskette drives, I think...
   Do While GetNextRemovableDrive(DigtaDrive)
      S = ""
      On Error Resume Next
      S = Dir(DigtaDrive & ":\DSS", vbDirectory)
      On Error GoTo 0
      If S = "DSS" Then
         GetDigtaDSSFolder = DigtaDrive & ":\DSS\"
         Exit Function
      Else
         DigtaDrive = Chr$(Asc(DigtaDrive) + 1)
      End If
   Loop
End Function
Private Function GetNextRemovableDrive(ByRef DriveLetter As String) As Boolean

   Dim DriveType As Long
   
   Do While DriveLetter <= "Z"
      DriveType = GetDriveType(DriveLetter & ":\")
      If DriveType = 2 Then
         GetNextRemovableDrive = True
         Exit Function
      End If
      DriveLetter = Chr$(Asc(DriveLetter) + 1)
   Loop
   GetNextRemovableDrive = False
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

    CreateTempFileName = CreateTempPath() & CStr(CLng(Timer)) & "." & ExtensionExclDot
End Function
Public Function CreateTempPath() As String

    Dim strBufferString As String
    Dim lngResult As Long
    strBufferString = String(MAX_BUFFER_LENGTH, "X")
    lngResult = GetTempPath(MAX_BUFFER_LENGTH, strBufferString)
    CreateTempPath = mId(strBufferString, 1, lngResult)
End Function
Public Sub KillFileIgnoreError(Filename As String)

   On Error Resume Next
   Kill Filename
End Sub
Public Function StringReplace(ByVal Str As String, SubStrToReplace As String, InsertInstead As String) As String

   Dim Pos As Integer
   
   Pos = InStr(Str, SubStrToReplace)
   Do While Pos > 0
      Str = Left$(Str, Pos - 1) & InsertInstead & mId$(Str, Pos + Len(SubStrToReplace))
      'Pos = 0
      On Error Resume Next
      Pos = InStr(Pos + Len(InsertInstead), Str, SubStrToReplace)
      On Error GoTo 0
   Loop
   StringReplace = Str
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
   PatId = StringReplace(PatId, "-", "")
   
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
   Dim S As String

   Secs = Sec Mod 60
   Mins = (Sec \ 60) Mod 60
   Hours = (Sec \ 60) \ 60
   If Hours <> 0 Then
      S = Format$(Hours, "0") & ":"
   End If
   FormatLength = S & Format$(Mins, "00") & ":" & Format$(Secs, "00")
End Function
Public Function WriteStringToTempFile(S As String) As String

   Dim Pathname As String
   Dim f As Integer
   
   Pathname = CreateTempFileName("tmp")
   f = FreeFile
   Open Pathname For Binary Access Write As #f
   Put #f, , S
   Close #f
   WriteStringToTempFile = Pathname
End Function
Public Function ReadStringFromTempFile(Pathname As String) As String

   Dim f As Integer
   Dim S As String
   
   S = Space$(FileLen(Pathname))
   f = FreeFile
   Open Pathname For Binary Access Read As #f
   Get #f, , S
   Close #f
   ReadStringFromTempFile = S
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
   Dim S As String

   CommandLine = GlobalCommandLine
   S = CommandString(CommandLine)
   Do While S <> ""
      If UCase$(S) = "/" & UCase$(KeyWithoutSlash) Then
         CommandValue = CommandString(CommandLine)
         Exit Function
      End If
      S = CommandString(CommandLine)
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
   
   Dim S As String
   
   S = Space(512)
   GetComputerName S, Len(S)
   GetStationName = Trim$(S)
End Function

Public Function MsgWaitObj(Interval As Long, _
                           Optional hObj As Long = 0&, _
                           Optional nObj As Long = 0&) As Long
                           
   Dim T As Long, T1 As Long
   
   If Interval <> INFINITE Then
       T = GetTickCount()
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
           T1 = GetTickCount()
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
   Dim f As Form
   
   sFormName = UCase$(FormName)
   
   For Each f In Forms
      If UCase$(f.Name) = sFormName Then
        IsLoaded = True
        Exit Function
      End If
   Next
End Function

