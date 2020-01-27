VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimeKeeping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Values(100) As Long    '0 is new 100 old
Private AccValue As Double
Private NumberOfValues As Long
Private IsStarted As Boolean
Private StartValue As Long
Private MaxValue As Long
Private MinValue As Long

Private Sub Class_Initialize()

   AccValue = 0
   NumberOfValues = 0
End Sub
Public Sub StartMeasure()

   StartValue = GetTickCount()
   IsStarted = True
End Sub
Public Sub StopMeasure()

   Dim StopValue As Long
   Dim MeasuredValue As Long

   StopValue = GetTickCount()
   If IsStarted Then
      MeasuredValue = StopValue - StartValue
      If MeasuredValue < 0 Then
         MeasuredValue = 0
      End If
      SaveValue MeasuredValue
   End If
   IsStarted = False
End Sub
Private Sub SaveValue(NewValue As Long)

   Dim I As Integer

   AccValue = AccValue + NewValue
   NumberOfValues = NumberOfValues + 1
   If NumberOfValues <= UBound(Values) Then
      For I = NumberOfValues To 1 Step -1
         Values(I) = Values(I - 1)
      Next I
   Else
      For I = UBound(Values) To 1 Step -1
         Values(I) = Values(I - 1)
      Next I
   End If
   Values(0) = NewValue
   If NumberOfValues = 1 Then
      MaxValue = NewValue
      MinValue = NewValue
   Else
      If NewValue > MaxValue Then
         MaxValue = NewValue
      End If
      If NewValue < MinValue Then
         MinValue = NewValue
      End If
   End If
End Sub
Public Property Get NumberOfMeasurements() As Long

   NumberOfMeasurements = NumberOfValues
End Property
Public Property Get TotalTimeInMilliSec() As Double

   TotalTimeInMilliSec = AccValue
End Property
Public Property Get MaxValuesInSlidingWindow() As Long

   MaxValuesInSlidingWindow = UBound(Values)
End Property
Public Property Get MeanTimeInMilliSec() As Long

   If NumberOfValues > 0 Then
      MeanTimeInMilliSec = CLng(AccValue / NumberOfValues)
   Else
      MeanTimeInMilliSec = 0
   End If
End Property
Public Property Get SlidingMeanTimeInMilliSec(NumberOfMeasurmentsInSlidingMeanValue As Long, Optional ExcludeExtremeValues As Boolean = False) As Long

   Dim NumOfM As Long
   Dim Tot As Double
   Dim NumInTot As Integer
   Dim I As Integer
   
   NumOfM = NumberOfMeasurmentsInSlidingMeanValue
   If NumberOfMeasurements > UBound(Values) Then
      NumOfM = UBound(Values)
   End If
   If NumOfM > NumberOfValues Then
      NumOfM = NumberOfValues
   End If
   If NumOfM > 0 Then
      For I = 0 To NumOfM - 1
         If ExcludeExtremeValues Then
            If Values(I) < MaxValue And Values(I) > MinValue Then
               Tot = Tot + Values(I)
               NumInTot = NumInTot + 1
            End If
         Else
            Tot = Tot + Values(I)
            NumInTot = NumInTot + 1
         End If
      Next I
      If NumInTot > 0 Then
         SlidingMeanTimeInMilliSec = CLng(Tot / NumInTot)
      Else
         SlidingMeanTimeInMilliSec = 0
      End If
   Else
      SlidingMeanTimeInMilliSec = 0
   End If
End Property
Public Property Get LastMeasurement() As Long

   LastMeasurement = Values(0)
End Property
