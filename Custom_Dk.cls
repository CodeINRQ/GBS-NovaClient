VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustom_Dk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsCustom

Public Function clsCustom_CheckDictationForMandatoryData(ByVal OkByFormTest As Boolean, Dict As clsDict, BeforeSave As Boolean) As Boolean
   
   Dim Ok As Boolean
   Dim OkPatId As Boolean
   
   Ok = OkByFormTest
   If BeforeSave Then
      If Not clsCustom_CheckPatId(Dict.Pat.PatId) And Not Dict.NoPatient Then
         If Len(Dict.Pat.PatId) > 0 Then
            OkPatId = MsgBox(Client.Texts.Txt(1030122, "Är cpr-nummret korrekt?"), vbYesNo) = vbYes
            If OkPatId Then
               Ok = True
            Else
               Exit Function
            End If
         End If
      End If
      If Dict.OrgId = 0 Then
         MsgBox Client.Texts.Txt(1030123, "Välj avdelning och diktattyp"), vbExclamation
         Exit Function
      End If
      If Not Ok Then
         MsgBox Client.Texts.Txt(1030101, "Uppgifterna är inte kompletta!"), vbCritical
      End If
   End If

   clsCustom_CheckDictationForMandatoryData = Ok
End Function

Public Function clsCustom_CheckPatId(ByVal PatId As String) As Boolean
   
   Dim Result As Integer
   Dim S As String
   
   On Error GoTo CheckPatId_Err
   
   S = clsCustom_FormatPatIdForStoring(PatId)
   
   If Not Client.SysSettings.DictInfoMandatoryPatId And Len(PatId) = 0 Then
      clsCustom_CheckPatId = True
      Exit Function
   End If

   If Len(S) <> 10 Then
      clsCustom_CheckPatId = False
      Exit Function
   End If
   
   'If Not Client.SysSettings.DictInfoPatIdChecksum Then
   '   clsCustom_CheckPatId = True
   '   Exit Function
   'End If
   
   Result = 0
   Result = Result + GetPositionValue(S, 1, 4)
   Result = Result + GetPositionValue(S, 2, 3)
   Result = Result + GetPositionValue(S, 3, 2)
   Result = Result + GetPositionValue(S, 4, 7)
   Result = Result + GetPositionValue(S, 5, 6)
   Result = Result + GetPositionValue(S, 6, 5)
   Result = Result + GetPositionValue(S, 7, 4)
   Result = Result + GetPositionValue(S, 8, 3)
   Result = Result + GetPositionValue(S, 9, 2)
   Result = Result + GetPositionValue(S, 10, 1)

   If (Result Mod 11) = 0 Then
      clsCustom_CheckPatId = True
   Else
      clsCustom_CheckPatId = False
   End If
   Exit Function
   
CheckPatId_Err:
   clsCustom_CheckPatId = False
   Exit Function
End Function

Private Function GetPositionValue(S As String, Pos As Integer, Weight As Integer) As Integer

   GetPositionValue = (Asc(mId(S, Pos, 1)) - 48) * Weight
End Function

Public Function clsCustom_CheckPatname(ByVal PName As String) As Boolean

   If Client.SysSettings.DictInfoMandatoryPatName Then
      clsCustom_CheckPatname = Len(PName) > 0
   Else
      clsCustom_CheckPatname = True
   End If
End Function

Public Function clsCustom_CheckCurrentPatientTextbox(ByVal PatId As String) As Boolean

   clsCustom_CheckCurrentPatientTextbox = Len(PatId) >= 10
End Function
Public Function clsCustom_CheckCurrentPatientBeforeUse(ByVal PatId As String) As Boolean

   Dim Ok As Boolean
   
   Ok = clsCustom_CheckPatId(PatId)
   If Ok Then
      clsCustom_CheckCurrentPatientBeforeUse = True
   Else
      Ok = MsgBox(Client.Texts.Txt(1030122, "Är cpr-nummret korrekt?"), vbYesNo) = vbYes
      clsCustom_CheckCurrentPatientBeforeUse = Ok
   End If
End Function
Public Function clsCustom_RndPatId() As String

   clsCustom_RndPatId = ""
End Function
Public Function clsCustom_RndPatName() As String

   clsCustom_RndPatName = RndFirstName() & " " & RndLastName()
End Function
Private Function RndFirstName() As String

'   Dim I As Integer
'
'   I = Int(Rnd * 21)
'
'   Select Case I
'      Case 0:  RndFirstName = "Jenny"
'      Case 1:  RndFirstName = "Lars"
'      Case 2:  RndFirstName = "Eva"
'      Case 3:  RndFirstName = "Frida"
'      Case 4:  RndFirstName = "Sven"
'      Case 5:  RndFirstName = "Per"
'      Case 6:  RndFirstName = "Björn"
'      Case 7:  RndFirstName = "Olof"
'      Case 8:  RndFirstName = "Matilda"
'      Case 9:  RndFirstName = "Sverker"
'      Case 10: RndFirstName = "Ulf"
'      Case 11: RndFirstName = "Ture"
'      Case 12: RndFirstName = "Charlotte"
'      Case 13: RndFirstName = "Pelle"
'      Case 14: RndFirstName = "Ludvig"
'      Case 15: RndFirstName = "Adam"
'      Case 16: RndFirstName = "Svante"
'      Case 17: RndFirstName = "Lotta"
'      Case 18: RndFirstName = "Lena"
'      Case 19: RndFirstName = "Emma"
'      Case 20: RndFirstName = "Josephin"
'   End Select
End Function
Private Function RndLastName() As String

'   Dim I As Integer
'
'   I = Int(Rnd * 21)
'
'   Select Case I
'      Case 0:  RndLastName = "Andersson"
'      Case 1:  RndLastName = "Petersson"
'      Case 2:  RndLastName = "Larsson"
'      Case 3:  RndLastName = "Blomgren"
'      Case 4:  RndLastName = "Svensson"
'      Case 5:  RndLastName = "Persson"
'      Case 6:  RndLastName = "Lindström"
'      Case 7:  RndLastName = "Hålgersson"
'      Case 8:  RndLastName = "Lundström"
'      Case 9:  RndLastName = "Zetterström"
'      Case 10: RndLastName = "Green"
'      Case 11: RndLastName = "Hagberg"
'      Case 12: RndLastName = "Grip"
'      Case 13: RndLastName = "Storm"
'      Case 14: RndLastName = "Johansson"
'      Case 15: RndLastName = "Ringqvist"
'      Case 16: RndLastName = "Carlsson"
'      Case 17: RndLastName = "Stolpe"
'      Case 18: RndLastName = "Fransson"
'      Case 19: RndLastName = "Nilssson"
'      Case 20: RndLastName = "Petersson"
'   End Select
End Function
Public Function clsCustom_FormatPatId(PatId As String) As String
   
   If Len(PatId) = 10 Then
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
