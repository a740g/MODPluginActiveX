VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{7F4D3A79-E3C0-4F36-A9B5-A9F7068D2550}#24.0#0"; "ModPlugin32.ocx"
Begin VB.Form frmModPlug32Test 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ModPlugin ActiveX Test Application"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "frmModPlug32Test.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin ModPlugin32.ModPlugin ModPlugin1 
      Height          =   735
      Left            =   60
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1296
      BackColor       =   12632256
      SpectrumHighColor=   16711935
      Repeat          =   -1  'True
      UpdateInterval  =   100
      AutoPlay        =   -1  'True
   End
   Begin VB.HScrollBar hsbVolume 
      Height          =   255
      Left            =   2940
      Max             =   100
      TabIndex        =   14
      Top             =   2280
      Width           =   2595
   End
   Begin VB.HScrollBar hsbSeek 
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "&Stop"
      Height          =   495
      Left            =   2970
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Default         =   -1  'True
      Height          =   495
      Left            =   1590
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   60
      TabIndex        =   7
      Top             =   2580
      Width           =   5655
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Wolf"
         Height          =   495
         Left            =   4320
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Thunder"
         Height          =   495
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Exit"
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Finale"
         Height          =   495
         Left            =   2280
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   495
      Left            =   210
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   495
      Left            =   4350
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlModFile 
      Left            =   5160
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ModPlugin32.PlaySound PlaySound2 
      Left            =   480
      Top             =   960
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin ModPlugin32.PlaySound PlaySound1 
      Left            =   480
      Top             =   960
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Volume"
      Height          =   195
      Left            =   2940
      TabIndex        =   16
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Seek:"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label lblPosition 
      AutoSize        =   -1  'True
      Caption         =   "00:00:00/00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1665
      TabIndex        =   1
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Set your options."
      Height          =   195
      Left            =   4380
      TabIndex        =   2
      Top             =   1020
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Right click above."
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   1020
      Width           =   1290
   End
End
Attribute VB_Name = "frmModPlug32Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bDoSeek As Boolean

Private Sub cmdAbout_Click()
    ModPlugin1.About
End Sub

Private Sub cmdLoad_Click()
    With cdlModFile
        .CancelError = True
        .DefaultExt = ".mod"
        .Filter = "All supported|*.mod;*.mdz;*.s3m;*.s3z;*.xm;*.xmz;*.it;*.itz;*.wav;*.zip|ProTracker|*.mod;*.mdz|ScreamTracker|*.s3m;*.s3z|FastTracker|*.xm;*.xmz|ImpulseTracker|*.it;*.itz|PCM Waveform audio|*.wav|Zip compressed|*.zip|All files|*.*"
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        On Error GoTo errHandleOpen
        .ShowOpen
        On Error GoTo 0
    End With
    
    ModPlugin1.File = cdlModFile.FileName

errHandleOpen:
    Exit Sub
End Sub

Private Sub cmdPlay_Click()
    ModPlugin1.Play
End Sub

Private Sub cmdStop_Click()
    ModPlugin1.Pause
End Sub

Private Sub Form_Load()
    cmdPlay.Enabled = False
    cmdStop.Enabled = False
    hsbVolume.Value = ModPlugin1.Volume
    hsbSeek.Enabled = False
    PlaySound2.Location = Resource
    PlaySound2.Play "Thunder"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PlaySound2.Location = Resource
    PlaySound2.Wait = True
    PlaySound2.Play "Wolf"
    PlaySound2.Wait = False
End Sub

Private Sub Command1_Click()
    PlaySound2.Location = FileName
    PlaySound2.Play "SystemStart"
End Sub

Private Sub Command10_Click()
    PlaySound1.Location = Resource
    PlaySound1.Repeat = True
    PlaySound1.Play App.Path & "\fruity44.wav"
    PlaySound1.Repeat = False
End Sub

Private Sub Command3_Click()
    PlaySound2.Location = Resource
    Screen.MousePointer = vbHourglass
    PlaySound2.Wait = True
    PlaySound2.Play "Wolf"
    PlaySound2.Wait = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command4_Click()
    PlaySound2.Location = Resource
    PlaySound2.Repeat = True
    PlaySound2.Play "Thunder"
    PlaySound2.Repeat = False
End Sub

Private Sub Command6_Click()
    PlaySound2.Location = FileName
    PlaySound2.Play "SystemExit"
End Sub

Private Sub hsbSeek_Change()
    If (bDoSeek) Then
        ModPlugin1.Position = hsbSeek.Value
        bDoSeek = False
    End If
End Sub

Private Sub hsbSeek_Scroll()
    bDoSeek = True
End Sub

Private Sub hsbVolume_Scroll()
    ModPlugin1.Volume = hsbVolume.Value
End Sub

Private Sub ModPlugin1_Loaded(ByVal bLoaded As Boolean)
    Me.Caption = App.Title & IIf(bLoaded, " [loaded]", " [error]")
    If (bLoaded) Then
        hsbSeek.Min = 0
        hsbSeek.Max = ModPlugin1.MaxPosition
        cmdPlay.Enabled = True
        cmdStop.Enabled = False
        hsbSeek.Enabled = True
    Else
        cmdPlay.Enabled = False
        cmdStop.Enabled = False
        hsbSeek.Enabled = False
    End If
End Sub

Private Sub ModPlugin1_Played(ByVal lCurSec As Long, ByVal lTotSec As Long, ByVal lCurPos As Long, ByVal lTotPos As Long)
    lblPosition.Caption = Format(TimeSerial(0, 0, lCurSec), "hh:mm:ss") & "/" & Format(TimeSerial(0, 0, lTotSec), "hh:mm:ss")
    hsbSeek.Value = lCurPos
End Sub

Private Sub ModPlugin1_Status(ByVal bPlaying As Boolean)
    Me.Caption = App.Title & IIf(bPlaying, " [playing]", " [stopped]")
    If (bPlaying) Then
        cmdStop.Enabled = True
        cmdPlay.Enabled = False
    Else
        cmdStop.Enabled = False
        cmdPlay.Enabled = True
    End If
End Sub

Private Sub ModPlugin1_VolumeChanged(ByVal bVolume As Byte)
    hsbVolume.Value = bVolume
End Sub

Private Sub PlaySound2_LoadResource(ByVal sName As String, cBuffer() As Byte)
    ' Must be static!!!
    Static cBuf() As Byte
    
    ' Free any memory allocated
    Erase cBuf
    
    'MsgBox "Loading " & sName & " from resource file..." & vbCrLf & "We could have even loaded the data off a file too!", vbInformation
    
    ' Load resource data
    cBuf = LoadResData(sName, "Wave")
    
    ' Pass buffer address to control
    cBuffer = cBuf
End Sub

Private Sub PlaySound1_LoadResource(ByVal sName As String, cBuffer() As Byte)
    Dim i As Integer
    Static cBuf() As Byte
    
    i = FreeFile
    Open sName For Binary Access Read As i
    
    ReDim cBuf(LOF(i)) As Byte
    
    Get i, , cBuf
    
    Close i
    
    cBuffer = cBuf
End Sub

