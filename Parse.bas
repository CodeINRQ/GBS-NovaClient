Attribute VB_Name = "modParse"
Option Explicit

Public Function FindString(ByVal S As String, ByVal ParseString As String) As String

   Dim Res As String
   Dim StartPosString As String
   Dim EndPosString As String
   Dim StartPos As Integer
   Dim EndPos As Integer
   
   If Len(ParseString) = 0 Then
      FindString = S
      Exit Function
   End If
   
   StartPosString = ConsumeToNextChar(ParseString, ";")
   EndPosString = ConsumeToNextChar(ParseString, ";")
   
   StartPos = FindPos(S, StartPosString)
   If StartPos > 0 Then
      S = mId$(S, StartPos)
   Else
      S = ""
   End If
   EndPos = FindPos(S, EndPosString)
   If EndPos > 0 Then
      S = Left$(S, EndPos - 1)
   End If
   FindString = Trim$(S)
End Function
Private Function FindPos(ByVal S As String, ByVal PosString As String) As Integer

   Dim FirstPos As Integer
   Dim MAXLEN As Integer
   Dim DelimitString As String
   Dim DelimitCount As Integer
   Dim StartNumeric As Integer
   Dim StartAlfa As Integer
   Dim StartControl As Integer
   
   Dim Ret As Integer
   Dim I As Integer
   Dim Pos As Integer
   
   If Len(PosString) = 0 Then
      Exit Function
   End If
   
   On Error Resume Next
   FirstPos = CInt(ConsumeToNextChar(PosString, ","))
   MAXLEN = CInt(ConsumeToNextChar(PosString, ","))
   DelimitString = ConsumeToNextChar(PosString, ",")
   DelimitCount = CInt(ConsumeToNextChar(PosString, ","))
   StartNumeric = CInt(ConsumeToNextChar(PosString, ","))
   StartAlfa = CInt(ConsumeToNextChar(PosString, ","))
   StartControl = CInt(ConsumeToNextChar(PosString, ","))
   
   S = mId$(S, FirstPos, MAXLEN)
   Ret = 0
   
   If Len(DelimitString) > 0 Then
      Pos = 1
      For I = 1 To DelimitCount
         Pos = InStr(Pos, S, DelimitString)
         If Pos = 0 Then Exit For
      Next I
      If Pos > 0 Then
         Ret = Pos
      End If
   End If
   
   If StartNumeric > 0 Then
      I = FindFirstNumeric(S, Ret + 1)
      If I > 0 Then
         Ret = I
      End If
   End If
   
   If StartAlfa > 0 Then
      I = FindFirstAlfa(S, Ret + 1)
      If I > 0 Then
         Ret = I
      End If
   End If
   
   If StartControl > 0 Then
      I = FindFirstControl(S, Ret + 1)
      If I > 0 Then
         Ret = I
      End If
   End If
   
   If Ret = 0 Then
      Ret = MAXLEN + 1
   End If
   FindPos = Ret
End Function
Private Function FindFirstNumeric(S, StartPos) As Integer

   Dim C As String
   Dim I As Integer
   
   For I = StartPos To Len(S)
      C = mId$(S, I, 1)
      If C >= "0" And C <= "9" Then
         FindFirstNumeric = I
         Exit For
      End If
   Next I
End Function
Private Function FindFirstAlfa(S, StartPos) As Integer

   Dim C As String
   Dim I As Integer
   
   For I = StartPos To Len(S)
      C = mId$(S, I, 1)
      If C >= "A" Then
         FindFirstAlfa = I
         Exit For
      End If
   Next I
End Function
Private Function FindFirstControl(S, StartPos) As Integer

   Dim C As String
   Dim I As Integer
   
   For I = StartPos To Len(S)
      C = mId$(S, I, 1)
      If C < "0" Then
         FindFirstControl = I
         Exit For
      End If
   Next I
End Function
Private Function ConsumeToNextChar(ByRef S As String, C As String) As String

   Dim Pos As Integer
   
   Pos = InStr(S, C)
   If Pos > 0 Then
      ConsumeToNextChar = Left$(S, Pos - 1)
      S = mId$(S, Pos + 1)
   Else
      ConsumeToNextChar = S
      S = ""
   End If
End Function

