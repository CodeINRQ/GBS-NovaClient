VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustom_Standard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsCustom

Public Function clsCustom_CheckDictationForMandatoryData(ByVal OkByFormTest As Boolean, Dict As clsDict, BeforeSave As Boolean) As Boolean
   
   Dim Ok As Boolean
   
   Ok = OkByFormTest
   If BeforeSave And Not Ok Then
      MsgBox Client.Texts.Txt(1030101, "Uppgifterna är inte kompletta!"), vbCritical
   End If
   clsCustom_CheckDictationForMandatoryData = Ok
End Function

Public Function clsCustom_CheckPatId(ByVal PatId As String) As Boolean

   'Returns true for a Correct number, else false
   Dim Siffra(9) As Integer
   Dim Resultat As Integer
   Dim I As Integer
   Dim Century As String
   
   If Not Client.SysSettings.DictInfoMandatoryPatId And Len(PatId) = 0 Then
      clsCustom_CheckPatId = True
      Exit Function
   End If
   
   If Client.SysSettings.DictInfoAlfaInPatid Then
      If Client.SysSettings.DictInfoMandatoryPatId Then
         clsCustom_CheckPatId = Len(PatId) > 0
         Exit Function
      Else
         clsCustom_CheckPatId = True
         Exit Function
      End If
   End If
   
   'Remove "-" if there is one
   PatId = clsCustom_FormatPatIdForStoring(PatId)
   
   If Client.SysSettings.DictInfoMandatoryPatIdCentury Then
      'check length
      If Len(PatId) <> 12 Then
         clsCustom_CheckPatId = False
         Exit Function
      End If
   End If
   
   If Len(PatId) = 12 Then
      Century = Left$(PatId, 2)
      PatId = mId$(PatId, 3)
      If (Century <> "19" And Century <> "20") Then
         clsCustom_CheckPatId = False
         Exit Function
      End If
   Else
      If Len(PatId) <> 10 Then
         clsCustom_CheckPatId = False
         Exit Function
      End If
   End If
   
   If Not Client.SysSettings.DictInfoPatIdChecksum Then
      clsCustom_CheckPatId = True
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
      clsCustom_CheckPatId = True
   Else
      clsCustom_CheckPatId = False
   End If
End Function
Public Function clsCustom_CheckPatname(ByVal PName As String) As Boolean

   If Client.SysSettings.DictInfoMandatoryPatName Then
      clsCustom_CheckPatname = Len(PName) > 0
   Else
      clsCustom_CheckPatname = True
   End If
End Function
Public Function clsCustom_CheckCurrentPatientTextbox(ByVal PatId As String) As Boolean

   clsCustom_CheckCurrentPatientTextbox = clsCustom_CheckPatId(PatId)
End Function

Public Function clsCustom_CheckCurrentPatientBeforeUse(ByVal PatId As String) As Boolean

   clsCustom_CheckCurrentPatientBeforeUse = clsCustom_CheckPatId(PatId)
End Function

Public Function clsCustom_RndPatId() As String

   Dim S As String
   Dim I As Integer
   
   If Rnd > 0.9 Then
      S = "20" & Format$(Int(Rnd * 5), "00")
   Else
      S = "19" & Format$(Int(Rnd * 100), "00")
   End If
   S = S & Format$(Int(Rnd * 12) + 1, "00")
   S = S & Format$(Int(Rnd * 28) + 1, "00")
   S = S & Format$(Int(Rnd * 1000), "000")
   For I = 0 To 9
      If clsCustom_CheckPatId(S & Chr$(Asc("0") + I)) Then
         clsCustom_RndPatId = S & Chr$(Asc("0") + I)
         Exit Function
      End If
   Next I
   clsCustom_RndPatId = S & "0"
End Function
Public Function clsCustom_RndPatName() As String

   clsCustom_RndPatName = RndFirstName() & " " & RndLastName()
End Function
Private Function RndFirstName() As String

   Dim I As Integer
   
   I = Int(Rnd * 21)
   
   Select Case I
      Case 0:  RndFirstName = "Jenny"
      Case 1:  RndFirstName = "Lars"
      Case 2:  RndFirstName = "Eva"
      Case 3:  RndFirstName = "Frida"
      Case 4:  RndFirstName = "Sven"
      Case 5:  RndFirstName = "Per"
      Case 6:  RndFirstName = "Björn"
      Case 7:  RndFirstName = "Olof"
      Case 8:  RndFirstName = "Matilda"
      Case 9:  RndFirstName = "Sverker"
      Case 10: RndFirstName = "Ulf"
      Case 11: RndFirstName = "Ture"
      Case 12: RndFirstName = "Charlotte"
      Case 13: RndFirstName = "Pelle"
      Case 14: RndFirstName = "Ludvig"
      Case 15: RndFirstName = "Adam"
      Case 16: RndFirstName = "Svante"
      Case 17: RndFirstName = "Lotta"
      Case 18: RndFirstName = "Lena"
      Case 19: RndFirstName = "Emma"
      Case 20: RndFirstName = "Josephin"
   End Select
End Function
Private Function RndLastName() As String

   Dim I As Integer
   
   I = Int(Rnd * 21)
   
   Select Case I
      Case 0:  RndLastName = "Andersson"
      Case 1:  RndLastName = "Petersson"
      Case 2:  RndLastName = "Larsson"
      Case 3:  RndLastName = "Blomgren"
      Case 4:  RndLastName = "Svensson"
      Case 5:  RndLastName = "Persson"
      Case 6:  RndLastName = "Lindström"
      Case 7:  RndLastName = "Hålgersson"
      Case 8:  RndLastName = "Lundström"
      Case 9:  RndLastName = "Zetterström"
      Case 10: RndLastName = "Green"
      Case 11: RndLastName = "Hagberg"
      Case 12: RndLastName = "Grip"
      Case 13: RndLastName = "Storm"
      Case 14: RndLastName = "Johansson"
      Case 15: RndLastName = "Ringqvist"
      Case 16: RndLastName = "Carlsson"
      Case 17: RndLastName = "Stolpe"
      Case 18: RndLastName = "Fransson"
      Case 19: RndLastName = "Nilsson"
      Case 20: RndLastName = "Petersson"
   End Select
End Function
Public Function clsCustom_FormatPatId(PatId As String) As String
   
   If Len(PatId) = 12 Then
      clsCustom_FormatPatId = Left$(PatId, 8) & "-" & mId$(PatId, 9)
   ElseIf Len(PatId) = 10 Then
      clsCustom_FormatPatId = Left$(PatId, 6) & "-" & mId$(PatId, 7)
   Else
      clsCustom_FormatPatId = PatId
   End If
End Function
Public Function clsCustom_FormatPatIdForStoring(ByVal S As String) As String

   S = StringReplace(S, "-", "")
   S = StringReplace(S, "/", "")
   S = StringReplace(S, "\", "")
   S = StringReplace(S, ".", "")
   S = StringReplace(S, ",", "")
   S = StringReplace(S, "+", "")
   clsCustom_FormatPatIdForStoring = S
End Function
