VERSION 5.00
Begin VB.Form frmCalibMic 
   Caption         =   "Kalibrering av mikrofon"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   Icon            =   "CalibMic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7545
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1400100"
   Begin CareTalk.ucVUmeter ucVUmeter1 
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   318
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Visa volymreglage"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Tag             =   "1400101"
      Top             =   3240
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Spara inst�llningar"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Tag             =   "1400102"
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.Label lblText 
         Alignment       =   2  'Center
         Height          =   2535
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmCalibMic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function VCStoreIni _
    Lib "Helper.dll" _
    Alias "_VolumeControl_StoreIni@4" (ByVal IniFileName As String) As Long

Private Declare Function VCPrepareSettingsForPlayer _
    Lib "Helper.dll" _
    Alias "_VolumeControl_PrepareSettingsForPlayer@8" (ByVal bWithRecorder As Long, _
                                                       ByVal IniFileName As String) As Long

Private Declare Function VCResetSettingsForPlayer _
    Lib "Helper.dll" _
    Alias "_VolumeControl_ResetSettingsForPlayer@0" () As Long
    
Private Declare Function VCBeginRecord _
    Lib "Helper.dll" _
    Alias "_VolumeControl_OnBeginRecord@0" () As Long

Private Declare Function VCEndRecord _
    Lib "Helper.dll" _
    Alias "_VolumeControl_OnEndRecord@0" () As Long

Private Declare Function VCShowRecordSettingsDialog _
    Lib "Helper.dll" _
    Alias "_VolumeControl_ShowRecordSettingsDialog@0" () As Long

Private Declare Function VCUnShowRecordSettingsDialog _
    Lib "Helper.dll" _
    Alias "_VolumeControl_UnShowRecordSettingsDialog@0" () As Long

Private WithEvents G As clsDSSRecorder
Attribute G.VB_VarHelpID = -1
Private TemFileName As String

Private Sub chkShow_Click()

   If chkShow.Value = vbChecked Then
      VCShowRecordSettingsDialog
   Else
      VCUnShowRecordSettingsDialog
   End If
End Sub

Private Sub cmdSave_Click()

   SaveCalibration
   Unload Me
End Sub

Private Sub Form_Load()

   Dim Hw As Gru_Harware

   CenterAndTranslateForm Me, frmMain

   lblText.Caption = vbLf & vbLf & Client.Texts.Txt(1400103, "S�kning efter mikrofonen sker...")
   Me.Show
   WindowFloating Me, True
   
   RecorderInUse = True
   Set G = Client.DSSRec
   G.GetHardWare Hw
   
   If Hw <> GRU_HW_RECORD Then
      MsgBox Client.Texts.Txt(1400104, "Kan inte hitta mikrofonen")
   Else
      lblText.Caption = HelpText()
      
      VCShowRecordSettingsDialog
      
      TemFileName = CreateTempFileName("dss")
      KillFileIgnoreError TemFileName
      G.LoadFile TemFileName, 0, 1
      G.Rec False
      Me.ucVUmeter1.Value = 0
   End If
   
Form_Load_Exit:
   Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

   CleanUpBeforeClosing
End Sub
Private Sub CleanUpBeforeClosing()

   If Not G Is Nothing Then
      G.PlayStop
      G.CloseFile
      Set G = Nothing
   End If
   VCUnShowRecordSettingsDialog
   KillFileIgnoreError TemFileName
   RecorderInUse = False
End Sub
Private Sub G_GruEvent(EventType As Gru_Event, Data As Long)

   On Error Resume Next
   Select Case EventType
      Case GRU_INPUTCHANGE
         Me.ucVUmeter1.Value = Data
   End Select
End Sub
Private Function HelpText() As String

   Dim Res As String
   
'   Res = "St�ll skjutreglaget p� mikrofonen p� l�ge START." & vbLf & vbLf & _
'       "St�ll in reglagen f�r ing�ngsniv�n till line-in resp. mikrofonen " & vbLf & _
'       "s� att niv�indikatorn stannar inom det gr�na omr�det." & vbLf & vbLf & _
'       "Om ing�ngsniv�n inte kan st�llas in optimalt: g� till Alternativ (avancerade) " & vbLf & _
'       "i ljudegenskaperna och justera eventuella mikrofonf�rst�rkare." & vbLf & vbLf & _
'       "Vi rekommenderar att mikrofonk�nsligheten st�lls in p� den l�gsta niv�n." & vbLf & _
'       "P� s� s�tt reduceras biljuden under inspelningen."
       
   Res = Client.Texts.Txt(1400105, "Normalt ska reglaget f�r k�nslighet p� mikrofonen st�llas i sitt l�gsta l�ge.") & vbLf & _
         Client.Texts.Txt(1400106, "D� reduceras biljud under inspelningen. St�ll sedan skjutreglaget p� mikrofonen p� l�ge START (r�tt fast sken).") & vbLf & vbLf & _
         Client.Texts.Txt(1400107, "St�ll huvudreglaget f�r inspelningsniv� p� max. Reglaget kan kallas t ex Wave in, Master.") & vbLf & _
         Client.Texts.Txt(1400108, "Kontrollera att mikrofoning�ngen �r aktiv. Det ska sitta en bock i valet som kan kallas V�lj, Anv�nd eller Select.") & vbLf & vbLf & _
         Client.Texts.Txt(1400109, "Tala i mikronen i normal samtalston (samma r�stl�ge som kommer anv�ndas vid diktering). Kontrollera niv�indikatorn nedan") & vbLf & _
         Client.Texts.Txt(1400110, "och justera inspelningsniv�n f�r mikrofon s� att stapeln inte �verskrider halva det gr�na omr�det vid normalt tal.")

   HelpText = Res
End Function


