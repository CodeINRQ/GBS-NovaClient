VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStore As String
Private mStoreFiltered As String
Private mFilterSection As String

Private Const RecDelimit = "¤"
Private Const FieldDelemit = "#"

Public Property Let Filter(Section As String)

   mStoreFiltered = mStore
   mFilterSection = NormalizedKey(Section)
End Property
Public Function GetNextFromFilter(ByRef Section As String, ByRef Key As String, ByRef Value As String) As Boolean

   Dim Rec As String

   Do While Len(mStoreFiltered) > 0
      Rec = GetRecord(mStoreFiltered)
      SplitRecord Rec, Section, Key, Value
      If Section = mFilterSection Or Len(mFilterSection) = 0 Then
         GetNextFromFilter = True
         Exit Function
      End If
   Loop
   GetNextFromFilter = False
End Function
Public Sub AddString(Section As String, Key As String, Value As String)

   RemoveFromStore Section, Key
   AddToStore Section, Key, Value
End Sub
Public Sub AddBool(Section As String, Key As String, Value As Boolean)

   Dim BValue As String
   
   If Value Then
      BValue = "Y"
   Else
      BValue = "N"
   End If
   AddString Section, Key, BValue
End Sub
Public Sub AddLong(Section As String, Key As String, Value As Long)

   Dim LValue As String
   
   LValue = CStr(Value)
   AddString Section, Key, LValue
End Sub
Public Function GetString(Section As String, Key As String, Default As String) As String

   Dim S As String
   Dim Record As String
   Dim Sec As String
   Dim K As String
   Dim V As String
   
   S = mStore
   Do While Len(S) > 0
      Record = GetRecord(S)
      SplitRecord Record, Sec, K, V
      If NormalizedKey(Section) = NormalizedKey(Sec) And NormalizedKey(Key) = NormalizedKey(K) Then
         GetString = V
         Exit Function
      End If
   Loop
   GetString = Default
   AddString Section, Key, Default
End Function
Public Function GetBool(Section As String, Key As String, Default As Boolean) As Boolean

   Dim BDefault As String
   
   If Default Then
      BDefault = "Y"
   Else
      BDefault = "N"
   End If
   GetBool = UCase$(GetString(Section, Key, BDefault)) = "Y"
End Function
Public Function GetLong(Section As String, Key As String, Default As Long) As Long

   On Error Resume Next
   GetLong = CLng(GetString(Section, Key, CStr(Default)))
End Function
Public Property Get Serialized() As String

   Serialized = mStore
End Property
Public Property Let Serialized(Serialized As String)

   mStore = StringReplace(StringReplace(Serialized, vbLf, ""), vbCr, "")
End Property
Private Function NormalizedKey(S As String) As String

   NormalizedKey = UCase$(Trim$(S))
End Function
Private Sub AddToStore(Section As String, Key As String, Value As String)

   mStore = mStore & Section & FieldDelemit & Key & FieldDelemit & Value & RecDelimit
End Sub
Private Function GetRecord(ByRef S) As String

   Dim P As Integer
   
   P = InStr(S, RecDelimit)
   If P > 0 Then
      GetRecord = Left$(S, P)
      S = Mid$(S, P + 1)
   End If
End Function
Private Sub SplitRecord(ByVal R As String, ByRef S As String, ByRef K As String, ByRef V As String)

   Dim P
   
   P = InStr(R, FieldDelemit)
   If P > 0 Then
      S = Left$(R, P - 1)
      R = Mid$(R, P + 1)
   Else
      S = ""
   End If
   
   P = InStr(R, FieldDelemit)
   If P > 0 Then
      K = Left$(R, P - 1)
      R = Mid$(R, P + 1)
   Else
      K = ""
   End If
   
   P = InStr(R, RecDelimit)
   If P > 0 Then
      V = Left$(R, P - 1)
   Else
      V = ""
   End If
End Sub
Private Sub RemoveFromStore(Section As String, Key As String)

   Dim R As String
   Dim S As String
   Dim K As String
   Dim V As String
   Dim NewStore As String
   Dim OldStore As String
   Dim P As Integer
   
   OldStore = mStore
   Do While Len(OldStore) > 0
      P = InStr(OldStore, RecDelimit)
      If P > 0 Then
         R = Left$(OldStore, P)
         OldStore = Mid$(OldStore, P + 1)
         SplitRecord R, S, K, V
         If NormalizedKey(Section) <> NormalizedKey(S) Or NormalizedKey(Key) <> NormalizedKey(K) Then
            NewStore = NewStore & R
         End If
      Else
         NewStore = NewStore & OldStore
         OldStore = ""
      End If
   Loop
   mStore = NewStore
End Sub

Private Sub Class_Initialize()

   mStore = ""
End Sub
