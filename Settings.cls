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
   Dim S As String * 10
   Dim I As Integer
   Dim Värde As Boolean

   If Len(mIniFilePath) > 0 Then
      I = GetPrivateProfileString(Section, Entry, "", S, Len(S), mIniFilePath)
      Tmp = Left$(S, I)
   End If
   If Tmp = "" Then
      Värde = Default
   Else
      Värde = UCase$(Left$(Tmp, 1)) = "Y"
   End If
   GetBool = Värde
   'If Debug_SkapaIniFil Then
   '   WriteBool Section, Entry, Värde
   'End If
End Function

Function GetLong(Section As String, Entry As String, Default As Long) As Long

   Dim I As Integer
   Dim S As String * 10
   Dim Tmp As String
   Dim Värde As Long

   If Len(mIniFilePath) > 0 Then
      I = GetPrivateProfileString(Section, Entry, "", S, Len(S), mIniFilePath)
      Tmp = Left$(S, I)
   End If
   Värde = Default
   On Error Resume Next
   Värde = CLng(Tmp)
   On Error GoTo 0
   GetLong = Värde
   'If Debug_SkapaIniFil Then
   '   WriteLong Section, Entry, Värde
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
   Dim S As String * 1024
   Dim Värde As String

   If Len(mIniFilePath) > 0 Then
      I = GetPrivateProfileString(Section, Entry, Default, S, Len(S), mIniFilePath)
      Värde = Left$(S, I)
   Else
      Värde = Default
   End If
   GetString = Värde
   'If Debug_SkapaIniFil Then
   '   WriteString Section, Entry, Värde
   'End If
End Function

Sub WriteBool(Section As String, Entry As String, Value As Boolean)

   Dim Värde As String

   If Value Then
      Värde = "Yes"
   Else
      Värde = "No"
   End If
   WriteString Section, Entry, Värde
End Sub

Sub WriteLong(Section As String, Entry As String, Value As Long)

   WriteString Section, Entry, Format$(Value)
End Sub

Sub WriteString(Section As String, Entry As String, Value As String)

   Dim I As Integer

   On Error GoTo WriteString_Err
   If Len(mIniFilePath) Then
      I = WritePrivateProfileString(Section, Entry, Value, mIniFilePath)
   End If
   Exit Sub
   
WriteString_Err:
   ErrorHandle "1300101", Err, 1300101, "Inställningar kan inte sparas", False
   Resume Next
End Sub


