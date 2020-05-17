VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3735
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrCoolBar 
      Interval        =   10
      Left            =   5220
      Top             =   60
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Click here for a surprise!"
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   5
      ToolTipText     =   "Closes the dialog box and saves any changes you have made."
      Top             =   2745
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   6
      ToolTipText     =   "Starts the Windows System Information application."
      Top             =   3195
      Width           =   1245
   End
   Begin VB.Label lblModPluginVersion 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   60
   End
   Begin VB.Image imgCoolBar 
      Height          =   165
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   5730
   End
   Begin VB.Image imgCoolBar 
      Height          =   165
      Index           =   1
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   5730
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   60
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   105
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   90
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":030A
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGVALSYSINFOLOC = "MSINFO"
Private Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Private Const gREGVALSYSINFO = "PATH"

Private Sub cmdSysInfo_Click()
    StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    picIcon.Picture = Me.Icon
    lblTitle.Caption = App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription.Caption = App.FileDescription & vbCrLf & vbCrLf & vbCrLf & App.LegalCopyright & vbCrLf & "mailto: v_2samg@hotmail.com"
    lblModPluginVersion.Caption = "ModPlugin Version " & ModPluginGetVersion
    imgCoolBar(1).Picture = LoadResPicture("GRADIENT_BAR", vbResBitmap)
    imgCoolBar(0).Picture = LoadResPicture("GRADIENT_BAR", vbResBitmap)
End Sub

Private Sub StartSysInfo()
    Dim rc As Long
    Dim SysInfoPath As String
    
    On Error GoTo SysInfoErr
  
    ' Try To Get System Info Program Path\Name From Registry...
    SysInfoPath = GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO)
    If SysInfoPath = vbNullString Then
        ' Try To Get System Info Program Path Only From Registry...
        SysInfoPath = GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC)
        If SysInfoPath <> vbNullString Then
            ' Validate Existance Of Known 32 Bit File Version
            If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            ' Error - File Can Not Be Found...
            Else
                GoTo SysInfoErr
            End If
        ' Error - Registry Entry Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    End If
    
    Shell SysInfoPath, vbNormalFocus
    
    Exit Sub
    
SysInfoErr:
    MsgBox "System Information is unavailable at this time.", vbExclamation
End Sub

Private Sub picIcon_Click()
    frmCoolStars.Show vbModal, Me
    Unload frmCoolStars
    Set frmCoolStars = Nothing
End Sub

Private Sub tmrCoolBar_Timer()
    ' Scrolling crap...
    If (imgCoolBar(1).Left < 0) Then
        imgCoolBar(1).Left = Me.ScaleWidth - 1
    End If
    
    ' Move from right to left to avoid flickering
    imgCoolBar(1).Left = imgCoolBar(1).Left - Screen.TwipsPerPixelX
    imgCoolBar(0).Left = imgCoolBar(1).Left - imgCoolBar(0).Width
    
    ' Release rest of time slice
    DoEvents
End Sub
