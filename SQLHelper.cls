VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum FldTypeEnum
   fldString
   fldStringNullIfBlank
   fldBoolean
   fldBooleanNullIfFalse
   fldDateTimeNullIfZero
   fldLong
   fldLongNullIfZero
   fldInteger
   fldIntegerNullIfZero
End Enum

Private Const MaxNumberOfMappings = 200

Private mTable As String

Private mFieldName(MaxNumberOfMappings) As String
Private mFieldValue(MaxNumberOfMappings) As String
Private mKeyField(MaxNumberOfMappings) As Boolean
Private mIdentityField(MaxNumberOfMappings) As Boolean

Private mNumberOfMappings As Integer

Public Property Let Table(Value As String)

   mTable = Value
End Property

Public Sub Clear()

   mNumberOfMappings = 0
End Sub
Public Sub AddFieldMapping(FieldName As String, FieldValue As Variant, FldType As FldTypeEnum, Optional KeyField As Boolean = False, Optional IdentityField As Boolean = False)

   If mNumberOfMappings <= UBound(mFieldName) Then
      mFieldName(mNumberOfMappings) = FieldName
      mFieldValue(mNumberOfMappings) = ConvertValue(FieldValue, FldType)
      mKeyField(mNumberOfMappings) = KeyField
      mIdentityField(mNumberOfMappings) = IdentityField
      mNumberOfMappings = mNumberOfMappings + 1
   End If
End Sub

Public Function CreateUpdateStmt() As String

   Dim SQL As String
   Dim WhereClause As String
   Dim I As Integer
   
   SQL = "UPDATE " & mTable & " SET "
   
   For I = 0 To mNumberOfMappings - 1
      
      If mKeyField(I) Then
         If Len(WhereClause) > 0 Then
            WhereClause = WhereClause & " AND "
         End If
         WhereClause = WhereClause & mFieldName(I) & "=" & mFieldValue(I)
      Else
         SQL = SQL & mFieldName(I) & "=" & mFieldValue(I) & ","
      End If
   Next I
   
   SQL = Left(SQL, Len(SQL) - 1) 'Remove trailing comma
   
   SQL = SQL & " WHERE " & WhereClause
   
   CreateUpdateStmt = SQL
End Function
Public Function CreateDeleteStmt() As String

   Dim SQL As String
   Dim WhereClause As String
   Dim I As Integer
   
   SQL = "DELETE " & mTable
   
   For I = 0 To mNumberOfMappings - 1
      
      If mKeyField(I) Then
         If Len(WhereClause) > 0 Then
            WhereClause = WhereClause & " AND "
         End If
         WhereClause = WhereClause & mFieldName(I) & "=" & mFieldValue(I)
      End If
   Next I
   
   SQL = SQL & " WHERE " & WhereClause
   
   CreateDeleteStmt = SQL
End Function
Public Function CreateInsertStmt() As String

   Dim SQL As String
   Dim I As Integer
   
   SQL = "INSERT INTO " & mTable & " ("
   
   For I = 0 To mNumberOfMappings - 1
      If Not mIdentityField(I) Then
         SQL = SQL & mFieldName(I) & ","
      End If
   Next I
   
   SQL = Left(SQL, Len(SQL) - 1) 'Remove trailing comma
   
   SQL = SQL & ") VALUES ("
   For I = 0 To mNumberOfMappings - 1
      If Not mIdentityField(I) Then
         SQL = SQL & mFieldValue(I) & ","
      End If
   Next I
   
   SQL = Left(SQL, Len(SQL) - 1) 'Remove trailing comma
   
   SQL = SQL & ")"
   
   CreateInsertStmt = SQL
End Function

Private Function ConvertValue(Value As Variant, FldType As FldTypeEnum) As String

   Dim Res As String
   
   Select Case FldType
      Case fldString
         Res = "'" & Value & "'"
      Case fldStringNullIfBlank
         If Len(Value) = 0 Then
            Res = "NULL"
         Else
            Res = "'" & Value & "'"
         End If
      Case fldBoolean
         If Value Then
            Res = "1"
         Else
            Res = "0"
         End If
      Case fldBooleanNullIfFalse
         If Value Then
            Res = "1"
         Else
            Res = "NULL"
         End If
      Case fldDateTimeNullIfZero
         If Value = 0 Then
            Res = "NULL"
         Else
            Res = Client.Server.SQLDateTimeString(CDate(Value))
         End If
      Case fldLong, fldInteger
         Res = CStr(Value)
      Case fldLongNullIfZero, fldIntegerNullIfZero
         If Value = 0 Then
            Res = "NULL"
         Else
            Res = CStr(Value)
         End If
   End Select
   ConvertValue = Res
End Function
