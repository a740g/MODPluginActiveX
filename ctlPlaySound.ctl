VERSION 5.00
Begin VB.UserControl PlaySound 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipBehavior    =   0  'None
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   Picture         =   "ctlPlaySound.ctx":0000
   PropertyPages   =   "ctlPlaySound.ctx":030A
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "ctlPlaySound.ctx":031D
   Windowless      =   -1  'True
End
Attribute VB_Name = "PlaySound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' PlaySound ActiveX Control
' Copyright (c) Samuel Gomes (Blade), 2001-2005
' mailto: v_2samg@hotmail.com

Option Explicit

'Default Property Values:
Const m_def_Repeat = 0
Const m_def_Wait = 0
Const m_def_Location = 0

'Enums:
Public Enum Locations
    FileName = 0
    Resource = 1
End Enum

'Property Variables:
Dim m_Repeat As Boolean
Dim m_Wait As Boolean
Dim m_Location As Byte

'Event Declarations:
Event LoadResource(ByVal sName As String, ByRef cBuffer() As Byte)
Attribute LoadResource.VB_Description = "Fired when the wave data from the resource is required. Argument sName is the name of the wave data and cBuffer is the sound data passed as an array of bytes."


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Play(ByVal sName As String) As Boolean
Attribute Play.VB_Description = "Plays the wave data. Argument sName is the name of the wave data in Resource or File."
    Static cSound() As Byte
    Dim lFlags As Long
    
    lFlags = IIf((m_Location = FileName), 0, SND_MEMORY)
    lFlags = lFlags Or IIf(m_Repeat, SND_LOOP, 0)
    lFlags = lFlags Or IIf(m_Wait, SND_SYNC, SND_ASYNC)
    
    ' Terminate any playing sounds
    Play = sndPlaySoundMemory(0, 0)
    
    ' Play from memory or file
    If (m_Location = FileName) Then
        'Debug.Print "Playing sound from file..."
        ' Play from file or sound event
        Play = sndPlaySoundFile(sName, lFlags)
    Else
        'Debug.Print "Playing sound from memory..."
        ' Trigger the LoadResource event
        RaiseEvent LoadResource(sName, cSound)
        ' Play from memory
        Play = sndPlaySoundMemory(VarPtr(cSound(0)), lFlags)
    End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Terminate() As Boolean
Attribute Terminate.VB_Description = "Stop sound playback."
    'Debug.Print "Stopping sound playback..."
    Terminate = sndPlaySoundMemory(0, 0)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Repeat() As Boolean
Attribute Repeat.VB_Description = "Loop sound playback?"
Attribute Repeat.VB_ProcData.VB_Invoke_Property = "pagPSGeneral"
    Repeat = m_Repeat
End Property

Public Property Let Repeat(ByVal New_Repeat As Boolean)
    m_Repeat = New_Repeat
    PropertyChanged "Repeat"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Wait() As Boolean
Attribute Wait.VB_Description = "Specifies if the wave data is to be played synchronously or asynchronously. Can be True or False."
Attribute Wait.VB_ProcData.VB_Invoke_Property = "pagPSGeneral"
    Wait = m_Wait
End Property

Public Property Let Wait(ByVal New_Wait As Boolean)
    m_Wait = New_Wait
    PropertyChanged "Wait"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get Location() As Byte
Attribute Location.VB_Description = "Specifies the location to load the wave data from. Can be Resource or File."
Attribute Location.VB_ProcData.VB_Invoke_Property = "pagPSGeneral"
    Location = m_Location
End Property

Public Property Let Location(ByVal New_Location As Byte)
    m_Location = New_Location
    PropertyChanged "Location"
End Property

Private Sub UserControl_Initialize()
    Terminate
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Repeat = m_def_Repeat
    m_Wait = m_def_Wait
    m_Location = m_def_Location
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Repeat = PropBag.ReadProperty("Repeat", m_def_Repeat)
    m_Wait = PropBag.ReadProperty("Wait", m_def_Wait)
    m_Location = PropBag.ReadProperty("Location", m_def_Location)
End Sub

Private Sub UserControl_Terminate()
    Terminate
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Repeat", m_Repeat, m_def_Repeat)
    Call PropBag.WriteProperty("Wait", m_Wait, m_def_Wait)
    Call PropBag.WriteProperty("Location", m_Location, m_def_Location)
End Sub

