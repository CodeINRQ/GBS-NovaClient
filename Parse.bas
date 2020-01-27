Attribute VB_Name = "modParse"
Option Explicit

Public Function FindString(ByVal s As String, ByVal ParseString As String) As String

   Dim Res As String
   Dim StartPosString As String
   Dim EndPosString As String
   Dim StartPos As Integer
   Dim EndPos As Integer
   
   If Len(ParseString) = 0 Then
      FindString = s
      Exit Function
   End If
   
   StartPosString = ConsumeToNextChar(ParseString, ";")
   EndPosString = ConsumeToNextChar(ParseString, ";")
   
   StartPos = FindPos(s, StartPosString)
   If StartPos > 0 Then
      s = mId$(s, StartPos)
   Else
      s = ""
   End If
   EndPos = FindPos(s, EndPosString)
   If EndPos > 0 Then
      s = Left$(s, EndPos - 1)
   End If
   FindString = Trim$(s)
End Function
Private Function FindPos(ByVal s As String, ByVal PosString As String) As Integer

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
   
   s = mId$(s, FirstPos, MAXLEN)
   Ret = 0
   
   If Len(DelimitString) > 0 Then
      Pos = 1
      For I = 1 To DelimitCount
         Pos = InStr(Pos, s, DelimitString)
         If Pos = 0 Then Exit For
      Next I
      If Pos > 0 Then
         Ret = Pos
      End If
   End If
   
   If StartNumeric > 0 Then
      I = FindFirstNumeric(s, Ret + 1)
      If I > 0 Then
         Ret = I
      End If
   End If
   
   If StartAlfa > 0 Then
      I = FindFirstAlfa(s, Ret + 1)
      If I > 0 Then
         Ret = I
      End If
   End If
   
   If StartControl > 0 Then
      I = FindFirstControl(s, Ret + 1)
      If I > 0 Then
         Ret = I
      End If
   End If
   
   If Ret = 0 Then
      Ret = MAXLEN + 1
   End If
   FindPos = Ret
End Function
Private Function FindFirstNumeric(s, StartPos) As Integer

   Dim C As String
   Dim I As Integer
   
   For I = StartPos To Len(s)
      C = mId$(s, I, 1)
      If C >= "0" And C <= "9" Then
         FindFirstNumeric = I
         Exit For
      End If
   Next I
End Function
Private Function FindFirstAlfa(s, StartPos) As Integer

   Dim C As String
   Dim I As Integer
   
   For I = StartPos To Len(s)
      C = mId$(s, I, 1)
      If C >= "A" Then
         FindFirstAlfa = I
         Exit For
      End If
   Next I
End Function
Private Function FindFirstControl(s, StartPos) As Integer

   Dim C As String
   Dim I As Integer
   
   For I = StartPos To Len(s)
      C = mId$(s, I, 1)
      If C < "0" Then
         FindFirstControl = I
         Exit For
      End If
   Next I
End Function
Public Function ConsumeToNextChar(ByRef s As String, C As String) As String

   Dim Pos As Integer
   
   Pos = InStr(s, C)
   If Pos > 0 Then
      ConsumeToNextChar = Left$(s, Pos - 1)
      s = mId$(s, Pos + 1)
   Else
      ConsumeToNextChar = s
      s = ""
   End If
End Function
