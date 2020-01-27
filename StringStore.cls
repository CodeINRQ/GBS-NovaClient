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

Public Sub EmptyStore()

   mStore = ""
End Sub

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

   Dim s As String
   Dim Record As String
   Dim Sec As String
   Dim K As String
   Dim V As String
   
   s = mStore
   Do While Len(s) > 0
      Record = GetRecord(s)
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
Private Function NormalizedKey(s As String) As String

   NormalizedKey = UCase$(Trim$(s))
End Function
Private Sub AddToStore(Section As String, Key As String, Value As String)

   mStore = mStore & Section & FieldDelemit & Key & FieldDelemit & Value & RecDelimit
End Sub
Private Function GetRecord(ByRef s) As String

   Dim P As Integer
   
   P = InStr(s, RecDelimit)
   If P > 0 Then
      GetRecord = Left$(s, P)
      s = mId$(s, P + 1)
   End If
End Function
Private Sub SplitRecord(ByVal r As String, ByRef s As String, ByRef K As String, ByRef V As String)

   Dim P
   
   P = InStr(r, FieldDelemit)
   If P > 0 Then
      s = Left$(r, P - 1)
      r = mId$(r, P + 1)
   Else
      s = ""
   End If
   
   P = InStr(r, FieldDelemit)
   If P > 0 Then
      K = Left$(r, P - 1)
      r = mId$(r, P + 1)
   Else
      K = ""
   End If
   
   P = InStr(r, RecDelimit)
   If P > 0 Then
      V = Left$(r, P - 1)
   Else
      V = ""
   End If
End Sub
Private Sub RemoveFromStore(Section As String, Key As String)

   Dim r As String
   Dim s As String
   Dim K As String
   Dim V As String
   Dim NewStore As String
   Dim OldStore As String
   Dim P As Integer
   
   OldStore = mStore
   Do While Len(OldStore) > 0
      P = InStr(OldStore, RecDelimit)
      If P > 0 Then
         r = Left$(OldStore, P)
         OldStore = mId$(OldStore, P + 1)
         SplitRecord r, s, K, V
         If NormalizedKey(Section) <> NormalizedKey(s) Or NormalizedKey(Key) <> NormalizedKey(K) Then
            NewStore = NewStore & r
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
