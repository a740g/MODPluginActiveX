VERSION 5.00
Begin VB.UserControl ModPlugin 
   CanGetFocus     =   0   'False
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   LockControls    =   -1  'True
   PropertyPages   =   "ctlModPlugin.ctx":0000
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
   ToolboxBitmap   =   "ctlModPlugin.ctx":0023
   Begin VB.Timer tmrUpdate 
      Left            =   2617
      Top             =   157
   End
End
Attribute VB_Name = "ModPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' ModPlugin 1.91+ Control for Visual Basic
' Copyright (c) Samuel Gomes, 2002-2005
' mailto: v_2samg@hotmail.com

Option Explicit

'Default Property Values:
Const m_def_MixingSoundRate = 44100
Const m_def_MixingStereo = True
Const m_def_Mixing16Bit = True
Const m_def_MixingOverSampling = True
Const m_def_MixingBassExpansion = True
Const m_def_MixingDolbySurround = True
Const m_def_AutoPlay = False
Const m_def_ControlStereo = True
Const m_def_UpdateInterval = 500
Const m_def_BackColor = &H0&
Const m_def_SpectrumLowColor = &HFF0000
Const m_def_VUMeterLowColor = &HFF00&
Const m_def_SpectrumHighColor = &HFF&
Const m_def_VUMeterHighColor = &HFF&
Const m_def_PositionTime = 0
Const m_def_File = ""
Const m_def_Repeat = False
Const m_def_Volume = 100
Const m_def_Title = "Loading..."
Const m_def_Position = 0
Const m_def_Ready = False
Const m_def_Playing = False
Const m_def_Version = 0
Const m_def_Length = 0
Const m_def_MaxPosition = 0

'Property Variables:
Dim m_MixingSoundRate As Long
Dim m_MixingStereo As Boolean
Dim m_Mixing16Bit As Boolean
Dim m_MixingOverSampling As Boolean
Dim m_MixingBassExpansion As Boolean
Dim m_MixingDolbySurround As Boolean
Dim m_AutoPlay As Boolean
Dim m_ControlStereo As Boolean
Dim m_UpdateInterval As Integer
Dim m_BackColor As OLE_COLOR
Dim m_SpectrumLowColor As OLE_COLOR
Dim m_VUMeterLowColor As OLE_COLOR
Dim m_SpectrumHighColor As OLE_COLOR
Dim m_VUMeterHighColor As OLE_COLOR
Dim m_PositionTime As Long
Dim m_File As String
Dim m_Repeat As Boolean
Dim m_Volume As Byte
Dim m_Title As String
Dim m_Position As Long
Dim m_Ready As Boolean
Dim m_Playing As Boolean
Dim m_Version As Single
Dim m_Length As Long
Dim m_MaxPosition As Long

'Event Declarations:
Event VolumeChanged(ByVal bVolume As Byte)
Attribute VolumeChanged.VB_Description = "Fired when the volume level is changed."
Event Loaded(ByVal bLoaded As Boolean)
Attribute Loaded.VB_Description = "Fired when a song is loaded and ready for playback."
Event Played(ByVal lCurSec As Long, ByVal lTotSec As Long, ByVal lCurPos As Long, ByVal lTotPos As Long)
Attribute Played.VB_Description = "Fired after each 'UpdateInterval'. See UpdateInterval property. Here lCSec is the current time played in seconds, lTSec is the totol time in seconds, lCPos is the current position in the Mod and lTPos is the total positions in the Mod."
Event Status(ByVal bPlaying As Boolean)
Attribute Status.VB_Description = "Fired when the status of song playback is changed (playing/stopped)."

'Private variables
Dim hMP As Long             ' ModPlugin handle
Dim bIsPlaying As Boolean   ' Flag to hack start/stop (UI element 3)

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get File() As String
Attribute File.VB_Description = "Mod file specification."
Attribute File.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    File = m_File
End Property

Public Property Let File(ByVal New_File As String)
    Dim lCtr As Long
    
    m_File = New_File
    PropertyChanged "File"
    
    RestartModPlugin
    
    'Debug.Print "Loading song..."
    If (ModPlug_Load(hMP, m_File) = 0) Then
        Err.Raise vbObjectError + 516, UserControl.Name, "Failed to load module file"
        Exit Property       ' what?!!! no harm done...
    End If
    
    ' Hold on till the MOD is up and loaded.
    Do
        DoEvents
        lCtr = lCtr + 1     ' infinite loop precaution
    Loop Until (Ready Or lCtr > 1073741824 Or (Not Ambient.UserMode))
    
    ' Fire the event now
    RaiseEvent Loaded(Ready)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,0,0
Public Property Get Repeat() As Boolean
Attribute Repeat.VB_Description = "Loop the song?"
Attribute Repeat.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    Repeat = m_Repeat
End Property

Public Property Let Repeat(ByVal New_Repeat As Boolean)
    If Ambient.UserMode Then Err.Raise 382
    m_Repeat = New_Repeat
    PropertyChanged "Repeat"
    
    RestartModPlugin
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,100
Public Property Get Volume() As Byte
Attribute Volume.VB_Description = "Playback volume."
Attribute Volume.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    'Debug.Print "Getting playback volume..."
    m_Volume = ModPlug_GetVolume(hMP)
    m_Volume = IIf(m_Volume > 100, 100, m_Volume)
    m_Volume = IIf(m_Volume < 0, 0, m_Volume)
    Volume = m_Volume
End Property

Public Property Let Volume(ByVal New_Volume As Byte)
    m_Volume = New_Volume
    PropertyChanged "Volume"
    m_Volume = IIf(m_Volume > 100, 100, m_Volume)
    m_Volume = IIf(m_Volume < 0, 0, m_Volume)
    
    'Debug.Print "Setting playback volume..."
    If (ModPlug_SetVolume(hMP, m_Volume) = 0) Then
        Err.Raise vbObjectError + 517, UserControl.Name, "Failed to set playback volume"
    End If
    
    ' Volume changed. Fire event.
    RaiseEvent VolumeChanged(m_Volume)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get Title() As String
Attribute Title.VB_Description = "General song title. 'Loading...' default."
Attribute Title.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    Title = m_Title
End Property

Public Property Let Title(ByVal New_Title As String)
    If Ambient.UserMode Then Err.Raise 382
    m_Title = New_Title
    PropertyChanged "Title"
    
    RestartModPlugin
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Play() As Boolean
Attribute Play.VB_Description = "Plays the loaded song."
    If Not (Ready) Then
        'Debug.Print "Reloading song..."
        File = m_File
    End If
    
    'Debug.Print "Playing song..."
    Play = (ModPlug_Play(hMP) <> 0)
    
    ' Save the playing flag
    bIsPlaying = Playing
    
    ' Fire status event
    RaiseEvent Status(bIsPlaying)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,0
Public Property Get Position() As Long
Attribute Position.VB_Description = "Current song position."
Attribute Position.VB_MemberFlags = "400"
    'Debug.Print "Getting playback position..."
    m_Position = ModPlug_GetCurrentPosition(hMP)
    Position = m_Position
End Property

Public Property Let Position(ByVal New_Position As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    m_Position = New_Position
    PropertyChanged "Position"
    
    'Debug.Print "Setting playback position..."
    If (ModPlug_SetCurrentPosition(hMP, m_Position) = 0) Then
        Err.Raise vbObjectError + 518, UserControl.Name, "Failed to set playback position"
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,2,0
Public Property Get Ready() As Boolean
Attribute Ready.VB_Description = "Is song loaded and ready to play?"
Attribute Ready.VB_MemberFlags = "400"
    'Debug.Print "Getting ready status..."
    m_Ready = (ModPlug_IsReady(hMP) <> 0)
    Ready = m_Ready
End Property

Public Property Let Ready(ByVal New_Ready As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_Ready = New_Ready
    PropertyChanged "Ready"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,1,2,0
Public Property Get Version() As Single
Attribute Version.VB_Description = "ModPlugin version number."
Attribute Version.VB_MemberFlags = "400"
    'Debug.Print "Getting ModPlugin version..."
    m_Version = ModPluginGetVersion
    Version = m_Version
End Property

Public Property Let Version(ByVal New_Version As Single)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_Version = New_Version
    PropertyChanged "Version"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Pause() As Boolean
Attribute Pause.VB_Description = "Pauses song playback."
    'Debug.Print "Stopping song playback..."
    Pause = (ModPlug_Stop(hMP) <> 0)
    
    ' Save playing flag
    bIsPlaying = Playing
    
    ' Fire status event
    RaiseEvent Status(bIsPlaying)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,2,0
Public Property Get Length() As Long
Attribute Length.VB_Description = "Song length in seconds."
Attribute Length.VB_MemberFlags = "400"
    'Debug.Print "Getting module length..."
    m_Length = ModPlug_GetSongLength(hMP) \ 1000
    Length = m_Length
End Property

Public Property Let Length(ByVal New_Length As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_Length = New_Length
    PropertyChanged "Length"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,2,0
Public Property Get MaxPosition() As Long
Attribute MaxPosition.VB_Description = "Maximum play position of the song."
Attribute MaxPosition.VB_MemberFlags = "400"
    'Debug.Print "Getting maximum module position..."
    m_MaxPosition = ModPlug_GetMaxPosition(hMP)
    MaxPosition = m_MaxPosition
End Property

Public Property Let MaxPosition(ByVal New_MaxPosition As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_MaxPosition = New_MaxPosition
    PropertyChanged "MaxPosition"
End Property

' Timer event to help hack some UI elements
Private Sub tmrUpdate_Timer()
    Dim bVol As Byte, lPos As Long, bPly As Boolean
    
    ' Volume update event (volume bar) hack (UI element 1)
    ' Get current volume directly
    bVol = ModPlug_GetVolume(hMP)
    ' Check if there is a volume mismatch
    If (bVol <> m_Volume) Then
        ' Call property to raise the event and set m_Volume
        Volume = bVol
    End If
    
    ' Played time/position event (seek bar) hack (UI element 2)
    ' Get current position directly
    lPos = ModPlug_GetCurrentPosition(hMP)
    ' Check if there is a position mismatch
    If (lPos <> m_Position) Then
        ' Fire event and call properties to set m_Position
        RaiseEvent Played(PositionTime, Length, Position, MaxPosition)
    End If
    
    ' Start/Stop button hack (UI element 3) last & final hack
    ' Get playing status
    bPly = Playing
    ' Check if this flag matches our global flag
    If (bPly <> bIsPlaying) Then
        ' We have a mismatch... simulate play/pause
        If (bPly) Then
            ' play started
            Play
        Else
            ' playback stopped
            Pause
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    RestartModPlugin
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_SpectrumLowColor = m_def_SpectrumLowColor
    m_VUMeterLowColor = m_def_VUMeterLowColor
    m_SpectrumHighColor = m_def_SpectrumHighColor
    m_VUMeterHighColor = m_def_VUMeterHighColor
    m_PositionTime = m_def_PositionTime
    m_File = m_def_File
    m_Repeat = m_def_Repeat
    m_Volume = m_def_Volume
    m_Title = m_def_Title
    m_Position = m_def_Position
    m_Ready = m_def_Ready
    m_Playing = m_def_Playing
    m_Version = m_def_Version
    m_Length = m_def_Length
    m_MaxPosition = m_def_MaxPosition
    m_UpdateInterval = m_def_UpdateInterval
    m_ControlStereo = m_def_ControlStereo
    m_MixingSoundRate = m_def_MixingSoundRate
    m_MixingStereo = m_def_MixingStereo
    m_Mixing16Bit = m_def_Mixing16Bit
    m_MixingOverSampling = m_def_MixingOverSampling
    m_MixingBassExpansion = m_def_MixingBassExpansion
    m_MixingDolbySurround = m_def_MixingDolbySurround
    m_AutoPlay = m_def_AutoPlay
    
    'Debug.Print "Initialized properties."
    
    RestartModPlugin
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_SpectrumLowColor = PropBag.ReadProperty("SpectrumLowColor", m_def_SpectrumLowColor)
    m_VUMeterLowColor = PropBag.ReadProperty("VUMeterLowColor", m_def_VUMeterLowColor)
    m_SpectrumHighColor = PropBag.ReadProperty("SpectrumHighColor", m_def_SpectrumHighColor)
    m_VUMeterHighColor = PropBag.ReadProperty("VUMeterHighColor", m_def_VUMeterHighColor)
    m_PositionTime = PropBag.ReadProperty("PositionTime", m_def_PositionTime)
    m_File = PropBag.ReadProperty("File", m_def_File)
    m_Repeat = PropBag.ReadProperty("Repeat", m_def_Repeat)
    m_Volume = PropBag.ReadProperty("Volume", m_def_Volume)
    m_Title = PropBag.ReadProperty("Title", m_def_Title)
    m_Position = PropBag.ReadProperty("Position", m_def_Position)
    m_Ready = PropBag.ReadProperty("Ready", m_def_Ready)
    m_Playing = PropBag.ReadProperty("Playing", m_def_Playing)
    m_Version = PropBag.ReadProperty("Version", m_def_Version)
    m_Length = PropBag.ReadProperty("Length", m_def_Length)
    m_MaxPosition = PropBag.ReadProperty("MaxPosition", m_def_MaxPosition)
    m_UpdateInterval = PropBag.ReadProperty("UpdateInterval", m_def_UpdateInterval)
    m_ControlStereo = PropBag.ReadProperty("ControlStereo", m_def_ControlStereo)
    m_MixingSoundRate = PropBag.ReadProperty("MixingSoundRate", m_def_MixingSoundRate)
    m_MixingStereo = PropBag.ReadProperty("MixingStereo", m_def_MixingStereo)
    m_Mixing16Bit = PropBag.ReadProperty("Mixing16Bit", m_def_Mixing16Bit)
    m_MixingOverSampling = PropBag.ReadProperty("MixingOverSampling", m_def_MixingOverSampling)
    m_MixingBassExpansion = PropBag.ReadProperty("MixingBassExpansion", m_def_MixingBassExpansion)
    m_MixingDolbySurround = PropBag.ReadProperty("MixingDolbySurround", m_def_MixingDolbySurround)
    m_AutoPlay = PropBag.ReadProperty("AutoPlay", m_def_AutoPlay)
    
    'Debug.Print "Properties read from property bag."
    
    RestartModPlugin
End Sub

Private Sub UserControl_Resize()
    If Not (Ambient.UserMode) Then
        RestartModPlugin
    End If
End Sub

Private Sub UserControl_Terminate()
    If (hMP <> 0) Then
        If (Playing) Then
            Pause
        End If
    
        ' Shutdown timer
        tmrUpdate.Interval = 0
        
        If (ModPlug_Destroy(hMP) = 0) Then
            Err.Raise vbObjectError + 515, UserControl.Name, "Failed to destroy ModPlugin object"
        End If
        
        'Debug.Print "ModPlugin destroyed."
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("SpectrumLowColor", m_SpectrumLowColor, m_def_SpectrumLowColor)
    Call PropBag.WriteProperty("VUMeterLowColor", m_VUMeterLowColor, m_def_VUMeterLowColor)
    Call PropBag.WriteProperty("SpectrumHighColor", m_SpectrumHighColor, m_def_SpectrumHighColor)
    Call PropBag.WriteProperty("VUMeterHighColor", m_VUMeterHighColor, m_def_VUMeterHighColor)
    Call PropBag.WriteProperty("PositionTime", m_PositionTime, m_def_PositionTime)
    Call PropBag.WriteProperty("File", m_File, m_def_File)
    Call PropBag.WriteProperty("Repeat", m_Repeat, m_def_Repeat)
    Call PropBag.WriteProperty("Volume", m_Volume, m_def_Volume)
    Call PropBag.WriteProperty("Title", m_Title, m_def_Title)
    Call PropBag.WriteProperty("Position", m_Position, m_def_Position)
    Call PropBag.WriteProperty("Ready", m_Ready, m_def_Ready)
    Call PropBag.WriteProperty("Playing", m_Playing, m_def_Playing)
    Call PropBag.WriteProperty("Version", m_Version, m_def_Version)
    Call PropBag.WriteProperty("Length", m_Length, m_def_Length)
    Call PropBag.WriteProperty("MaxPosition", m_MaxPosition, m_def_MaxPosition)
    Call PropBag.WriteProperty("UpdateInterval", m_UpdateInterval, m_def_UpdateInterval)
    Call PropBag.WriteProperty("ControlStereo", m_ControlStereo, m_def_ControlStereo)
    Call PropBag.WriteProperty("MixingSoundRate", m_MixingSoundRate, m_def_MixingSoundRate)
    Call PropBag.WriteProperty("MixingStereo", m_MixingStereo, m_def_MixingStereo)
    Call PropBag.WriteProperty("Mixing16Bit", m_Mixing16Bit, m_def_Mixing16Bit)
    Call PropBag.WriteProperty("MixingOverSampling", m_MixingOverSampling, m_def_MixingOverSampling)
    Call PropBag.WriteProperty("MixingBassExpansion", m_MixingBassExpansion, m_def_MixingBassExpansion)
    Call PropBag.WriteProperty("MixingDolbySurround", m_MixingDolbySurround, m_def_MixingDolbySurround)
    Call PropBag.WriteProperty("AutoPlay", m_AutoPlay, m_def_AutoPlay)
    
    'Debug.Print "Properties written to property bag."
End Sub

Private Sub RestartModPlugin()
    Dim sCreateStr As String
    
    ' Restart timer
    tmrUpdate.Interval = m_UpdateInterval
    
    ' Forcefully shutdown timer if in design mode
    On Error Resume Next
    If Not (Ambient.UserMode) Then
        tmrUpdate.Interval = 0
    End If
    On Error GoTo 0
    
    ' Shutdown if active
    If (hMP <> 0) Then
        If (Playing) Then
            Pause
        End If
    
        If (ModPlug_Destroy(hMP) = 0) Then
            Err.Raise vbObjectError + 515, UserControl.Name, "Failed to destroy ModPlugin object"
        End If
        
        'Debug.Print "ModPlugin destroyed."
    End If
    
    ' Set registry entries
    MixingSoundRate = MixingSoundRate
    MixingStereo = MixingStereo
    Mixing16Bit = Mixing16Bit
    MixingOverSampling = MixingOverSampling
    MixingBassExpansion = MixingBassExpansion
    MixingDolbySurround = MixingDolbySurround
    AutoPlay = AutoPlay
        
    ' Restart
    sCreateStr = "volume|" & m_Volume & "|" _
                & "loop|" & LCase(Trim(Str(m_Repeat))) & "|" _
                & "controls|" & IIf(m_ControlStereo, "stereo", "mono") & "|" _
                & "title|" & StrConv(Trim(m_Title), vbProperCase) & "|" _
                & "bgcolor|" & RGBToMPColor(m_BackColor) & "|" _
                & "spcolor|" & RGBToMPColor(m_SpectrumLowColor) & "|" _
                & "spcolorhi|" & RGBToMPColor(m_SpectrumHighColor) & "|" _
                & "vucolor|" & RGBToMPColor(m_VUMeterLowColor) & "|" _
                & "vucolorhi|" & RGBToMPColor(m_VUMeterHighColor) & "|"
                
    'Debug.Print "ModPlugin creation options: "; sCreateStr
    
    hMP = ModPlug_Create(sCreateStr)
    If (hMP = 0) Then
        Err.Raise vbObjectError + 513, UserControl.Name, "Failed to create ModPlugin object"
    End If
    
    If (ModPlug_SetWindow(hMP, UserControl.hWnd) = 0) Then
        Err.Raise vbObjectError + 514, UserControl.Name, "Failed to attach ModPlugin to window"
    End If
    
    ' Fire volume event
    RaiseEvent VolumeChanged(Volume)
    
    'Debug.Print "ModPlugin created."
End Sub

' Converts a color from RGB to ModPlugin color
Private Function RGBToMPColor(ByVal lRGB As OLE_COLOR) As String
    RGBToMPColor = "#" & Right("0" & Hex(RGBRed(lRGB)), 2) & Right("0" & Hex(RGBGreen(lRGB)), 2) & Right("0" & Hex(RGBBlue(lRGB)), 2)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,1,0,&h000000
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets the background color."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    If Ambient.UserMode Then Err.Raise 382
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    RestartModPlugin
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,1,0,&hff0000
Public Property Get SpectrumLowColor() As OLE_COLOR
Attribute SpectrumLowColor.VB_Description = "Lower spectrum color."
    SpectrumLowColor = m_SpectrumLowColor
End Property

Public Property Let SpectrumLowColor(ByVal New_SpectrumLowColor As OLE_COLOR)
    If Ambient.UserMode Then Err.Raise 382
    m_SpectrumLowColor = New_SpectrumLowColor
    PropertyChanged "SpectrumLowColor"
    
    RestartModPlugin
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,1,0,&h00ff00
Public Property Get VUMeterLowColor() As OLE_COLOR
Attribute VUMeterLowColor.VB_Description = "Lower vu-meter color."
    VUMeterLowColor = m_VUMeterLowColor
End Property

Public Property Let VUMeterLowColor(ByVal New_VUMeterLowColor As OLE_COLOR)
    If Ambient.UserMode Then Err.Raise 382
    m_VUMeterLowColor = New_VUMeterLowColor
    PropertyChanged "VUMeterLowColor"
    
    RestartModPlugin
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,1,0,&h0000ff
Public Property Get SpectrumHighColor() As OLE_COLOR
Attribute SpectrumHighColor.VB_Description = "Upper spectrum color."
    SpectrumHighColor = m_SpectrumHighColor
End Property

Public Property Let SpectrumHighColor(ByVal New_SpectrumHighColor As OLE_COLOR)
    If Ambient.UserMode Then Err.Raise 382
    m_SpectrumHighColor = New_SpectrumHighColor
    PropertyChanged "SpectrumHighColor"
    
    RestartModPlugin
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,1,0,&h0000ff
Public Property Get VUMeterHighColor() As OLE_COLOR
Attribute VUMeterHighColor.VB_Description = "Upper vu-meter color."
    VUMeterHighColor = m_VUMeterHighColor
End Property

Public Property Let VUMeterHighColor(ByVal New_VUMeterHighColor As OLE_COLOR)
    If Ambient.UserMode Then Err.Raise 382
    m_VUMeterHighColor = New_VUMeterHighColor
    PropertyChanged "VUMeterHighColor"
    
    RestartModPlugin
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,0
Public Property Get PositionTime() As Long
Attribute PositionTime.VB_Description = "Current song position in seconds."
Attribute PositionTime.VB_MemberFlags = "400"
    On Error Resume Next
    m_PositionTime = CSng(Length) * (CSng(Position) / CSng(MaxPosition))
    On Error GoTo 0
    PositionTime = m_PositionTime
End Property

Public Property Let PositionTime(ByVal New_PositionTime As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    m_PositionTime = New_PositionTime
    PropertyChanged "PositionTime"
    
    On Error Resume Next
    Position = CSng(MaxPosition) * (CSng(m_PositionTime) / CSng(Length))
    On Error GoTo 0
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,500
Public Property Get UpdateInterval() As Integer
Attribute UpdateInterval.VB_Description = "Changes the rate at which the Played event is fired. Defaults to 500 ms. i.e. 1/2 sec. Thefore the event is fired 2 times a second."
Attribute UpdateInterval.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    UpdateInterval = m_UpdateInterval
End Property

Public Property Let UpdateInterval(ByVal New_UpdateInterval As Integer)
    m_UpdateInterval = New_UpdateInterval
    PropertyChanged "UpdateInterval"
    
    If (Playing) Then
        tmrUpdate.Interval = m_UpdateInterval
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub About()
Attribute About.VB_Description = "Shows the about dialog."
Attribute About.VB_UserMemId = -552
    'Debug.Print "Showing the about dialog..."
    frmAbout.Show vbModal, Me
    Unload frmAbout
    Set frmAbout = Nothing
    'Debug.Print "Unloaded about dialog."
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,2,False
Public Property Get Playing() As Boolean
Attribute Playing.VB_Description = "Is the song playing?"
Attribute Playing.VB_MemberFlags = "400"
    'Debug.Print "Getting playback status..."
    m_Playing = (ModPlug_IsPlaying(hMP) <> 0)
    Playing = m_Playing
End Property

Public Property Let Playing(ByVal New_Playing As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_Playing = New_Playing
    PropertyChanged "Playing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,0,True
Public Property Get ControlStereo() As Boolean
Attribute ControlStereo.VB_Description = "True for stereo control, false for mono controls."
Attribute ControlStereo.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    ControlStereo = m_ControlStereo
End Property

Public Property Let ControlStereo(ByVal New_ControlStereo As Boolean)
    If Ambient.UserMode Then Err.Raise 382
    m_ControlStereo = New_ControlStereo
    PropertyChanged "ControlStereo"
    
    RestartModPlugin
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,44100
Public Property Get MixingSoundRate() As Long
Attribute MixingSoundRate.VB_Description = "Returns and sets the sound playback rate."
Attribute MixingSoundRate.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    MixingSoundRate = m_MixingSoundRate
End Property

Public Property Let MixingSoundRate(ByVal New_MixingSoundRate As Long)
    m_MixingSoundRate = New_MixingSoundRate
    PropertyChanged "MixingSoundRate"
    If (m_MixingSoundRate <= 0) Then m_MixingSoundRate = 44100
    ModPluginSetSoundRate m_MixingSoundRate
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MixingStereo() As Boolean
Attribute MixingStereo.VB_Description = "Enables/Disables stereo sound mixing."
Attribute MixingStereo.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    MixingStereo = m_MixingStereo
End Property

Public Property Let MixingStereo(ByVal New_MixingStereo As Boolean)
    m_MixingStereo = New_MixingStereo
    PropertyChanged "MixingStereo"
    If (m_MixingStereo) Then
        ModPluginSetSettings ModPluginGetSettings Or MPMIX_STEREO
    Else
        ModPluginSetSettings ModPluginGetSettings And (Not MPMIX_STEREO)
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Mixing16Bit() As Boolean
Attribute Mixing16Bit.VB_Description = "Eables/Disabled 16-bit sound mixing."
Attribute Mixing16Bit.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    Mixing16Bit = m_Mixing16Bit
End Property

Public Property Let Mixing16Bit(ByVal New_Mixing16Bit As Boolean)
    m_Mixing16Bit = New_Mixing16Bit
    PropertyChanged "Mixing16Bit"
    If (m_Mixing16Bit) Then
        ModPluginSetSettings ModPluginGetSettings Or MPMIX_16BIT
    Else
        ModPluginSetSettings ModPluginGetSettings And (Not MPMIX_16BIT)
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MixingOverSampling() As Boolean
Attribute MixingOverSampling.VB_Description = "Enables/Disables sound over sampling."
Attribute MixingOverSampling.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    MixingOverSampling = m_MixingOverSampling
End Property

Public Property Let MixingOverSampling(ByVal New_MixingOverSampling As Boolean)
    m_MixingOverSampling = New_MixingOverSampling
    PropertyChanged "MixingOverSampling"
    If (m_MixingOverSampling) Then
        ModPluginSetSettings ModPluginGetSettings And (Not MPMIX_DISABLE_OVERSAMPLING)
    Else
        ModPluginSetSettings ModPluginGetSettings Or MPMIX_DISABLE_OVERSAMPLING
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MixingBassExpansion() As Boolean
Attribute MixingBassExpansion.VB_Description = "Enables/Disables bass expansion during sound mixing."
Attribute MixingBassExpansion.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    MixingBassExpansion = m_MixingBassExpansion
End Property

Public Property Let MixingBassExpansion(ByVal New_MixingBassExpansion As Boolean)
    m_MixingBassExpansion = New_MixingBassExpansion
    PropertyChanged "MixingBassExpansion"
    If (m_MixingBassExpansion) Then
        ModPluginSetSettings ModPluginGetSettings Or MPMIX_BASS_EXPANSION
    Else
        ModPluginSetSettings ModPluginGetSettings And (Not MPMIX_BASS_EXPANSION)
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MixingDolbySurround() As Boolean
Attribute MixingDolbySurround.VB_Description = "Enables/Disables dolby surround sound mixing."
Attribute MixingDolbySurround.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    MixingDolbySurround = m_MixingDolbySurround
End Property

Public Property Let MixingDolbySurround(ByVal New_MixingDolbySurround As Boolean)
    m_MixingDolbySurround = New_MixingDolbySurround
    PropertyChanged "MixingDolbySurround"
    If (m_MixingDolbySurround) Then
        ModPluginSetSettings ModPluginGetSettings Or MPMIX_DOLBY_SURROUND
    Else
        ModPluginSetSettings ModPluginGetSettings And (Not MPMIX_DOLBY_SURROUND)
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get AutoPlay() As Boolean
Attribute AutoPlay.VB_Description = "Starts to play automatically when the music file is loaded."
Attribute AutoPlay.VB_ProcData.VB_Invoke_Property = "pagMPGeneral"
    AutoPlay = m_AutoPlay
End Property

Public Property Let AutoPlay(ByVal New_AutoPlay As Boolean)
    m_AutoPlay = New_AutoPlay
    PropertyChanged "AutoPlay"
    If (m_AutoPlay) Then
        ModPluginSetSettings ModPluginGetSettings And (Not MPMIX_DISABLE_AUTOPLAY)
    Else
        ModPluginSetSettings ModPluginGetSettings Or MPMIX_DISABLE_AUTOPLAY
    End If
End Property

