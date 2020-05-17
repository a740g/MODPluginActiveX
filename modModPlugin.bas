Attribute VB_Name = "modModPlugin"
' ModPlugin 1.91+ Interface for Visual Basic
' Copyright (c) Samuel Gomes, 2002-2005
' mailto: v_2samg@hotmail.com

Option Explicit
  
' This function is identical to ModPlug_Create, but you can simply
' use a string, instead of an array of strings. Each keyword must
' be separated by a '|' character (It works also with carriage return).
' For example, the command <loop="true" vucolor="#ff00ff"> should
' be sent as "loop|true|vucolor|#ff00ff|".
'The other options you may want to add are the following:
'- volume="x": set the volume of the song to x.
'  (Default="100", range="1" to "100")
'
'- loop="true": Loop the song. (Default:"false")
'
'- autostart="true": Start playing the song when loaded.
'  (Default="false"). (Also "autoplay" on versions 1.31+)
'
'- autonext="true": Jump to next mod when finished playing.
'  (Default="false") This is useful only if you have more
'  than one song in your page. (You should use loop="true"
'  instead to make the same song loop)
'
'- controls="none"/"stereo":
'  * "none": don't display anything - use with care
'   because there is no way the user will be able to stop the
'   mod (besides exiting the page). (same as "hidden="true")
'  * "stereo"(1.65+): the horizontal spectrum analyzer will
'   be split for right/left, on each side of the plugin.
'   In this case the number of bands will be (width - 184) / 16.
'  NOTE (1.75+): The option controls="smallconsole" has been
'  removed , it 's now automatic if height < 20.
'
'- hidden="true": same as controls="none".
'
'- title="song_title": Displays the text "song_title" when
'  the song is not yet loaded. (Default: displays "Loading...")
'
'- shuffle="true" (1.57+): goes along with autonext="true".
'  when a song finishes playing, the next song will be randomly
'  chosen. The shuffle option is a global flag and should appear
'  on the first EMBED section in the page. This flags sets
'  the autonext="true" flag automatically.
'
'- bgcolor="black"(default),"gray" (1.60+): Select a background color
'  for the plugin.
'  v1.75+: you can specify a color like bgcolor="#RRGGBB".
'
'- spcolor="red"(default),"green","blue" (1.61+): Select the color of
'  the spectrum analyzer.
'  v1.75+: you can specify a color like spcolor="#RRGGBB".
'
'- vucolor="color1", vucolorhi="color2" (1.75+): you can now change
'  the color of the VU-meter. The default for vucolor is green, and
'  is red for vucolorhi.
'
'- spcolorhi: defines the top-color of the spectrum analyzer (1.75+).
'  The default color is red. The middle color will be the mix between
'  spcolor and spcolorhi.
'
'- autoload="false" (1.77+): if autoload is "false", the file will not
'  start loading unless the user tries to play the mod. The default is
'  autoload="true".
'
'NOTE: (v1.43+)
'--------------
'- If you want to enable the VU-Meter, the width parameter
'  should be set to "168". The VU-Meter cannot be used with
'  the controls="smallconsole" option.
'- If the plugin is hidden (with controls="none" or hidden commands),
'  autostart will be set to "true".
'
'Spectrum Analyzer (1.60+)
'-------------------------
'- You can enable the spectrum analyzer by setting a height
'  of 96 (bottom spectrum) or a width of 336 (right spectrum).
'- v1.61+: You can have from 3 to 80 bands in your right spectrum
'  analyzer: set the width parameter to 176+(numbands*8).
'  The frequency range is between 86Hz (left) and 11KHz (right).
'- v1.65+: With the stereo spectrum, the number of bands on each
'  side will be (width - 184) / 16.
Public Declare Function ModPlug_Create Lib "npmod32" Alias "ModPlug_CreateEx" (ByVal lpszParams As String) As Long

' This function destroys a plugin created by ModPlug_Create.
Public Declare Function ModPlug_Destroy Lib "npmod32" (ByVal plugin As Long) As Long

' Returns the current playing position. (Between 0 and ModPlug_GetMaxPosition)
Public Declare Function ModPlug_GetCurrentPosition Lib "npmod32" (ByVal plugin As Long) As Long

' Returns the maximum position of the song.
Public Declare Function ModPlug_GetMaxPosition Lib "npmod32" (ByVal plugin As Long) As Long

' Gets the module length in milliseconds
Public Declare Function ModPlug_GetSongLength Lib "npmod32" (ByVal plugin As Long) As Long

' This function returns the current version of the plugin:
' ie: for version 1.75, it will return 0x175.
Public Declare Function ModPlug_GetVersion Lib "npmod32" () As Long

' Gets the playback volume.
Public Declare Function ModPlug_GetVolume Lib "npmod32" (ByVal plugin As Long) As Long

' Returns TRUE is the plugin is currently playing a song.
Public Declare Function ModPlug_IsPlaying Lib "npmod32" (ByVal plugin As Long) As Long

' Returns TRUE is a song is correctly loaded.
Public Declare Function ModPlug_IsReady Lib "npmod32" (ByVal plugin As Long) As Long

' This function loads a module.
Public Declare Function ModPlug_Load Lib "npmod32" (ByVal plugin As Long, ByVal lpszFileName As String) As Long

' This function starts playing the module loaded.
Public Declare Function ModPlug_Play Lib "npmod32" (ByVal plugin As Long) As Long

' Sets the current playing position.
Public Declare Function ModPlug_SetCurrentPosition Lib "npmod32" (ByVal plugin As Long, ByVal nPos As Long) As Long

' Sets the playback volume.
Public Declare Function ModPlug_SetVolume Lib "npmod32" (ByVal plugin As Long, ByVal nVol As Long) As Long

' This function should be called right after ModPlug_Create,
' with the window handle of the window where the plugin will
' draw itself. You are responsible for creating this child window.
Public Declare Function ModPlug_SetWindow Lib "npmod32" (ByVal plugin As Long, ByVal hWnd As Long) As Long

' This function stops playing.
Public Declare Function ModPlug_Stop Lib "npmod32" (ByVal plugin As Long) As Long

' Note: This declaration does not match that of WIN32API.TXT
' This is because I have used this funtion to play
' sound from memory and hence some modification was needed.
Public Declare Function sndPlaySoundFile Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function sndPlaySoundMemory Lib "winmm" Alias "sndPlaySoundA" (ByVal lSoundPtr As Long, ByVal uFlags As Long) As Long

Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'---------------------------------------------------------------
'-Registry API Declarations...
'---------------------------------------------------------------
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
' The following def was hacked to work for 32-bit LONGs
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long

'---------------------------------------------------------------
'- Registry Api Constants...
'---------------------------------------------------------------
' Reg Data Types...
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number

' Reg Create Type Values...
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

' Reg Key Security Options...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Public Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' Return Value...
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0

' Constants
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const SND_NOSTOP = &H10
Public Const SND_SYNC = &H0
Public Const MAX_PATH = 260

' Version of DLL in our resource (update if resource changed)
Public Const VER_NPMOD32_DLL_RES = 1.91
' This procedure will load resource strings associated with controls on a
' form based on the Resource ID stored in the Tag property of  a control.
' This module reads and writes registry keys.  Unlike the
' internal registry access methods of VB, it can read and
' write any registry keys with string values.

' ModPlugin mixer flags
Public Const MPMIX_STEREO = &H1&
Public Const MPMIX_16BIT = &H2&
Public Const MPMIX_DISABLE_OVERSAMPLING = &H10&
Public Const MPMIX_BASS_EXPANSION = &H20&
Public Const MPMIX_DOLBY_SURROUND = &H8&
Public Const MPMIX_DISABLE_AUTOPLAY = &H40&

'---------------------------------------------------------------
'- Registry Security Attributes TYPE...
'---------------------------------------------------------------
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

'-------------------------------------------------------------------------------------------------
' sample usage - Debug.Print UpodateKey(HKEY_CLASSES_ROOT, "keyname", "newvalue")
' HACK! HACK! HACK!
'-------------------------------------------------------------------------------------------------
Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As Long) As Boolean
    Dim rc As Long                                      ' Return Code
    Dim hKey As Long                                    ' Handle To A Registry Key
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' Registry Security Type
    
    lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- Create/Open Registry Key...
    '------------------------------------------------------------
    ' Create/Open //KeyRoot//KeyName
    rc = RegCreateKeyEx(KeyRoot, KeyName, 0, REG_DWORD, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, hKey, hDepth)
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Errors...
    
    '------------------------------------------------------------
    '- Create/Modify Key Value...
    '------------------------------------------------------------
    ' Create/Modify Key Value
    rc = RegSetValueEx(hKey, SubKeyName, 0, REG_DWORD, SubKeyValue, Len(SubKeyValue))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Error
    '------------------------------------------------------------
    '- Close Registry Key...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' Close Key
    
    UpdateKey = True                                    ' Return Success
    
    Exit Function                                       ' Exit
    
CreateKeyError:
    UpdateKey = False                                   ' Set Error Return Code
    rc = RegCloseKey(hKey)                              ' Attempt To Close Key
End Function

'-------------------------------------------------------------------------------------------------
' sample usage - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
      
    tmpVal = Left(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' Search Data Types...
    Case REG_SZ, REG_EXPAND_SZ                              ' String Registry Key Data Type
        sKeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        sKeyVal = Format("&h" + sKeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = sKeyVal                                   ' Return Value
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' Set Return Val To Empty String
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

' Converts a null-terminated string to a VB string
Public Function CStrToBStr(ByVal lpszString As String) As String
    lpszString = lpszString & vbNullChar
    CStrToBStr = Left(lpszString, InStr(lpszString, vbNullChar) - 1)
End Function

' Returns the expanded version of the pathname
Public Function FullPathName(ByVal sFileName As String) As String
    Dim myBuffer As String
    Dim junk As Long

    myBuffer = String(MAX_PATH + 1, vbNullChar)
    GetFullPathName sFileName, MAX_PATH, myBuffer, junk
    
    FullPathName = CStrToBStr(myBuffer)
End Function

' Fixes any problems in a pathname
Public Function FixPath(ByVal sPath As String) As String
    sPath = FullPathName(sPath)
    FixPath = sPath & IIf((Right(sPath, 1) = "\"), "", "\")
End Function

' Returns the system directory
Public Function SystemDirectory() As String
    Dim myBuffer As String
    Dim spaceRequired As Long
    
    ' Calculate space needed
    myBuffer = Space(3)
    spaceRequired = GetSystemDirectory(myBuffer, 2)

    myBuffer = String(spaceRequired + 1, vbNullChar)
    GetSystemDirectory myBuffer, spaceRequired
    
    SystemDirectory = FixPath(myBuffer)
End Function

Public Function RGBRed(ByVal lRGB As Long) As Byte
    RGBRed = CByte(lRGB And &HFF&)
End Function

Public Function RGBGreen(ByVal lRGB As Long) As Byte
    RGBGreen = CByte((lRGB \ 256&) And &HFF&)
End Function

Public Function RGBBlue(ByVal lRGB As Long) As Byte
    RGBBlue = CByte(lRGB \ 65536 And &HFF&)
End Function

' Warning: Call this function with CARE!!!
Public Sub SaveResource(ByVal sItem As String, ByVal sGroup As String, ByVal sDestin As String)
    Dim cData() As Byte
    Dim hFile As Integer
    
    'Debug.Print "Saving "; sItem; " from "; sGroup; " to "; sDestin; "..."
    
    ' Load data into memory
    cData = LoadResData(sItem, sGroup)
    
    ' Dump it where it belongs
    hFile = FreeFile
    Open sDestin For Output Access Write As hFile
    Close hFile
    Open sDestin For Binary Access Write As hFile
    Put hFile, , cData
    Close hFile
End Sub

' Used multiple times; so I write this function :)
Public Function ModPluginGetVersion() As Single
    Dim v As Long
    
    v = ModPlug_GetVersion
    ModPluginGetVersion = Val(Val(Hex(v \ 256&)) & "." & Val(Hex(v And &HFF&)))
End Function

' The following two keys contains the ModPlugin settings
' HKEY_CURRENT_USER\Software\Olivier Lapicque\MOD Plugin\Quality
' HKEY_CURRENT_USER\Software\Olivier Lapicque\MOD Plugin\Mixing_Rate

Public Sub ModPluginSetSoundRate(ByVal lSoundRate As Long)
    On Error Resume Next
    UpdateKey HKEY_CURRENT_USER, "Software\Olivier Lapicque\MOD Plugin", "Mixing_Rate", lSoundRate
End Sub

' Gets/Sets ModPlug settings
Public Function ModPluginGetSettings() As Long
    On Error Resume Next
    ModPluginGetSettings = Val(GetKeyValue(HKEY_CURRENT_USER, "Software\Olivier Lapicque\MOD Plugin", "Quality"))
End Function

Public Sub ModPluginSetSettings(ByVal lSettings As Long)
    On Error Resume Next
    UpdateKey HKEY_CURRENT_USER, "Software\Olivier Lapicque\MOD Plugin", "Quality", lSettings
End Sub

' ActiveX Entry Point
Public Sub Main()
    Dim sNPModDLL As String
    
    ' Normalize full DLL path
    sNPModDLL = SystemDirectory & "npmod32.dll"
    
    ' Save a copy of npmod32.dll from resource to the
    ' system directory if it's not there.
    If (Dir(sNPModDLL) = "") Then
        On Error GoTo errSetupFail
        SaveResource "NPMOD32_DLL", "BINARY", sNPModDLL
        On Error GoTo 0
    End If

    ' Check the DLL version
    If (ModPluginGetVersion < VER_NPMOD32_DLL_RES) Then
        On Error GoTo errSetupFail
        SaveResource "NPMOD32_DLL", "BINARY", sNPModDLL
        On Error GoTo 0
    End If

    Exit Sub
    
errSetupFail:
    MsgBox "Failed to setup ModPlugin ActiveX Control! Need administrator rights." & vbCrLf & "Ask your administrator to copy the latest version of 'npmod32.dll' to the Windows 'System/System32' folder. Also delete any local copies of 'npmod32.dll'.", vbCritical
End Sub

