VERSION 5.00
Begin VB.Form frmCoolStars 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Credits"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   FillStyle       =   0  'Solid
   Icon            =   "frmCoolStars.frx":0000
   LinkTopic       =   "frmCoolStars"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrUpdate 
      Interval        =   25
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblCoolText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   7680
      TabIndex        =   0
      Top             =   5760
      UseMnemonic     =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frmCoolStars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum StarShape
    Dot
    Box
    Round
End Enum

' Star type
Private Type StarType
    x As Single     ' precision required
    y As Single     ' precision required
    z As Long
    c As Long
    t As StarShape
End Type

Private Const MAXSTARS = 256       ' maximum no. of stars
Private Const DIVIDER = 1024

' The star array
Private Star(1 To MAXSTARS) As StarType
' Mode details
Private ScreenWidth As Integer, ScreenHeight As Integer

Private i As Long
Private a As Long, b As Long
Private cTop As Long

Private Sub Form_Load()
    Static cData() As Byte
    Dim i As Long
    Dim sTemp As String
    
    ' Erase data
    Erase cData
    
    ' Load the text data into a byte array and convert it to text
    cData = LoadResData("COOL_TEXT", "TEXT")
    
    For i = LBound(cData) To UBound(cData)
        sTemp = sTemp & Chr(cData(i))
    Next
    
    ' Load the sound and play it
    cData = LoadResData("RAINDROP", "WAVE")
    sndPlaySoundMemory VarPtr(cData(0)), SND_MEMORY Or SND_ASYNC Or SND_LOOP
    
    lblCoolText.Caption = App.Title & vbCrLf & _
        "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
        App.LegalCopyright & vbCrLf & _
        App.Comments & vbCrLf & vbCrLf & vbCrLf & sTemp
End Sub

Private Sub Form_Resize()
    ScreenWidth = Me.ScaleWidth
    ScreenHeight = Me.ScaleHeight
    a = ScreenWidth \ 2
    b = ScreenHeight \ 2
    lblCoolText.Left = (Me.ScaleWidth \ 2) - (lblCoolText.Width \ 2)
    cTop = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sndPlaySoundMemory 0, 0
End Sub

Private Sub tmrUpdate_Timer()
    Dim wh As Long
    
    lblCoolText.Top = cTop
    cTop = cTop - 1
    
    If (cTop < -lblCoolText.Height) Then
        cTop = Me.ScaleHeight
    End If
    
    ' Move the stars
    For i = 1 To MAXSTARS
        Me.FillColor = QBColor(0)
        wh = (Star(i).z - DIVIDER) \ 20
        Select Case Star(i).t
            Case Round
                If (wh > 0) Then Me.Circle (Star(i).x, Star(i).y), wh, QBColor(0)
            Case Box
                Me.Line (Star(i).x, Star(i).y)-Step(wh, wh), QBColor(0), BF
            Case Else
                Me.PSet (Star(i).x, Star(i).y), QBColor(0)
        End Select
        If (Star(i).x > ScreenWidth - 2 Or Star(i).x < 1 Or Star(i).y > ScreenHeight - 2 Or Star(i).y < 1) Then
            Star(i).x = Rnd * ScreenWidth
            Star(i).y = Rnd * ScreenHeight
            Star(i).z = DIVIDER
            Star(i).c = RGB(Fix(Rnd * 256), Fix(Rnd * 256), Fix(Rnd * 256))
            Star(i).t = Rnd * 3
        End If
        Star(i).z = Star(i).z + 1
        Star(i).x = (((Star(i).x) - a) * (Star(i).z / DIVIDER)) + a
        Star(i).y = (((Star(i).y) - b) * (Star(i).z / DIVIDER)) + b
        Me.FillColor = Star(i).c
        wh = (Star(i).z - DIVIDER) \ 20
        Select Case Star(i).t
            Case Round
                Me.Circle (Star(i).x, Star(i).y), wh, Star(i).c
            Case Box
                Me.Line (Star(i).x, Star(i).y)-Step(wh, wh), Star(i).c, BF
            Case Else
                Me.PSet (Star(i).x, Star(i).y), Star(i).c
        End Select
    Next
    
    ' Relinquish processor time
    DoEvents
End Sub
