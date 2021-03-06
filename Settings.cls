VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private mIniFilePath As String

Public Property Let IniFilePath(ByVal Path As String)

   mIniFilePath = Path
End Property
Public Property Get IniFilePath() As String

   IniFilePath = mIniFilePath
End Property

Function GetBool(Section As String, Entry As String, Default As Boolean) As Boolean
   
   Dim Tmp As String
   Dim s As String * 10
   Dim I As Integer
   Dim V�rde As Boolean

   If Len(mIniFilePath) > 0 Then
      I = GetPrivateProfileString(Section, Entry, "", s, Len(s), mIniFilePath)
      Tmp = Left$(s, I)
   End If
   If Tmp = "" Then
      V�rde = Default
   Else
      V�rde = UCase$(Left$(Tmp, 1)) = "Y"
   End If
   GetBool = V�rde
   'If Debug_SkapaIniFil Then
   '   WriteBool Section, Entry, V�rde
   'End If
End Function

Function GetLong(Section As String, Entry As String, Default As Long) As Long

   Dim I As Integer
   Dim s As String * 10
   Dim Tmp As String
   Dim V�rde As Long

   If Len(mIniFilePath) > 0 Then
      I = GetPrivateProfileString(Section, Entry, "", s, Len(s), mIniFilePath)
      Tmp = Left$(s, I)
   End If
   V�rde = Default
   On Error Resume Next
   V�rde = CLng(Tmp)
   On Error GoTo 0
   GetLong = V�rde
   'If Debug_SkapaIniFil Then
   '   WriteLong Section, Entry, V�rde
   'End If
End Function

Function GetFolder(Section As String, Entry As String, Default As String) As String

   Dim Katalog As String

   Katalog = GetString(Section, Entry, Default)
   If Right$(Katalog, 1) <> "\" Then
      Katalog = Katalog & "\"
   End If
   GetFolder = Katalog
   'If Debug_SkapaIniFil Then
   '   WriteString Section, Entry, Katalog
   'End If
End Function

Function GetString(Section As String, Entry As String, Default As String) As String

   Dim I As Integer
   Dim s As String * 1024
   Dim V�rde As String

   If Len(mIniFilePath) > 0 Then
      I = GetPrivateProfileString(Section, Entry, Default, s, Len(s), mIniFilePath)
      V�rde = Left$(s, I)
   Else
      V�rde = Default
   End If
   GetString = V�rde
   'If Debug_SkapaIniFil Then
   '   WriteString Section, Entry, V�rde
   'End If
End Function

Sub WriteBool(Section As String, Entry As String, Value As Boolean)

   Dim V�rde As String

   If Value Then
      V�rde = "Yes"
   Else
      V�rde = "No"
   End If
   WriteString Section, Entry, V�rde
End Sub

Sub WriteLong(Section As String, Entry As String, Value As Long)

   WriteString Section, Entry, Format$(Value)
End Sub

Sub WriteString(Section As String, Entry As String, Value As String)

   Dim I As Integer
   Dim Eno As Long

   On Error GoTo WriteString_Err
   If Len(mIniFilePath) > 0 Then
      I = WritePrivateProfileString(Section, Entry, Value, mIniFilePath)
   End If
   Exit Sub
   
WriteString_Err:
   Eno = Err.Number
   ErrorHandle "1300101", Eno, 1300101, "Inst�llningar kan inte sparas", False
   Resume Next
End Sub


