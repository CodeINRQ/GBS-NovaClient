VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmShowMe 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   11456
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmShowMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowNow(Url As String, Width As Integer, Height As Integer)

   If Len(Url) > 0 Then
      If Height > 0 And Width > 0 Then
         Me.Move 0, 0, Width, Height
      End If
      wb.Navigate Url
      Me.Show vbModal
      SetWindowTopMostAndForeground Me
   End If
End Sub

Private Sub Form_Load()

   CenterAndTranslateForm Me, frmMain
End Sub

Private Sub Form_Resize()

   wb.Move 0, 0, Me.Width, Me.Height
End Sub
