Attribute VB_Name = "Nebular"

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_SRC_MIDIVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_SRC_I25InVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_SRC_TADVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
                             (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_src_AUXVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
Public Const MIXERLINE_COMPONENTTYPE_SRC_PSPKVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MBOOST = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINEVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_SRC_CDVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20
' Mixer control types
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Declare Function RegisterDLL Lib "Regist10.dll" Alias "REGISTERDLL" _
(ByVal DllPath As String, bRegister As Boolean) As Boolean

Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal X As Long, ByVal uFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Declare Function mmioReadFormat Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByRef pch As waveFormat, ByVal cch As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
Declare Function mixerGetControlDetails Lib "winmm.dll" _
               Alias "mixerGetControlDetailsA" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
            
Declare Function mixerGetLineControls Lib "winmm.dll" _
               Alias "mixerGetLineControlsA" _
               (ByVal hmxobj As Long, _
               pmxlc As MIXERLINECONTROLS, _
               ByVal fdwControls As Long) As Long
               
Declare Function mixerGetLineInfo Lib "winmm.dll" _
               Alias "mixerGetLineInfoA" _
               (ByVal hmxobj As Long, _
               pmxl As MIXERLINE, _
               ByVal fdwInfo As Long) As Long
               
Declare Function mixerOpen Lib "winmm.dll" _
               (phmx As Long, _
               ByVal uMxId As Long, _
               ByVal dwCallback As Long, _
               ByVal dwInstance As Long, _
               ByVal fdwOpen As Long) As Long
               
Declare Function mixerSetControlDetails Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
               
Declare Sub CopyStructFromPtr Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (struct As Any, _
               ByVal ptr As Long, ByVal cb As Long)
               
Declare Sub CopyPtrFromStruct Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (ByVal ptr As Long, _
               struct As Any, _
               ByVal cb As Long)
               
Declare Function GlobalAlloc Lib "kernel32" _
               (ByVal wFlags As Long, _
               ByVal dwBytes As Long) As Long
               
Declare Function GlobalLock Lib "kernel32" _
               (ByVal hmem As Long) As Long
               
Declare Function GlobalFree Lib "kernel32" _
               (ByVal hmem As Long) As Long

Dim rc As Long



Public formatA As waveFormat
Dim hmmioOut As Long
Dim mmckinfoParentIn As MMCKINFO
Dim mmckinfoSubchunkIn As MMCKINFO
Dim bufferIn As Long
Dim hmem As Long
Public numSamples As Long
Public drawFrom As Long
Public drawTo As Long
Public fFileLoaded As Boolean

Type waveFormat
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Type MIXERCONTROL
   cbStruct As Long
   dwControlID As Long
   dwControlType As Long
   fdwControl As Long
   cMultipleItems As Long
   szShortName As String * MIXER_SHORT_NAME_CHARS
   szName As String * MIXER_LONG_NAME_CHARS
   lMinimum As Long
   lMaximum As Long
   Reserved(10) As Long
   End Type

Type MIXERCONTROLDETAILS
   cbStruct As Long
   dwControlID As Long
   cChannels As Long
   Item As Long
   cbDetails As Long
   paDetails As Long
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
   dwValue As Long
End Type

Type MIXERLINE
   cbStruct As Long
   dwDestination As Long
   dwSource As Long
   dwLineID As Long
   fdwLine As Long
   dwUser As Long
   dwComponentType As Long
   cChannels As Long
   cConnections As Long
   cControls As Long
   szShortName As String * MIXER_SHORT_NAME_CHARS
   szName As String * MIXER_LONG_NAME_CHARS
   dwType As Long
   dwDeviceID As Long
   wMid  As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
End Type

Type MIXERLINECONTROLS
   cbStruct As Long
   dwLineID As Long
                          
   dwControl As Long
   cControls As Long
   cbmxctrl As Long
   pamxctrl As Long
End Type

Type mmioinfo
   dwFlags As Long
   fccIOProc As Long
   pIOProc As Long
   wErrorRet As Long
   htask As Long
   cchBuffer As Long
   pchBuffer As String
   pchNext As String
   pchEndRead As String
   pchEndWrite As String
   lBufOffset As Long
   lDiskOffset As Long
   adwInfo(4) As Long
   dwReserved1 As Long
   dwReserved2 As Long
   hmmio As Long
End Type

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
    End Type

Function GetMixerControl(ByVal hmixer As Long, _
                        ByVal componentType As Long, _
                        ByVal ctrlType As Long, _
                        ByRef mxc As MIXERCONTROL) As Boolean
                        

   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType

   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = ctrlType
       mxlc.cControls = 1
       mxlc.cbmxctrl = Len(mxc)
       
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       
       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
       If (MMSYSERR_NOERROR = rc) Then
           GetMixerControl = True
           
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       Else
           GetMixerControl = False
       End If
       GlobalFree (hmem)
       Exit Function
   End If
   GetMixerControl = False
End Function

Function SetVolumeControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volume As Long) As Boolean
  Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.Item = 0
   mxcd.cbDetails = Len(vol)
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   vol.dwValue = volume
   
   CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetVolumeControl = True
   Else
       SetVolumeControl = False
   End If
End Function

Function SetPANControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volL As Long, ByVal volR As Long) As Boolean

                        
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol(1) As MIXERCONTROLDETAILS_UNSIGNED

   mxcd.Item = mxc.cMultipleItems
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cbStruct = Len(mxcd)
   mxcd.cbDetails = Len(vol(1))
   
   
   
   mxcd.cChannels = 2
   
   hmem = GlobalAlloc(&H40, Len(vol(1)))
   mxcd.paDetails = GlobalLock(hmem)
   
   vol(1).dwValue = volR
   vol(0).dwValue = volL
  
   
   CopyPtrFromStruct mxcd.paDetails, vol(1).dwValue, Len(vol(0)) * mxcd.cChannels
   CopyPtrFromStruct mxcd.paDetails, vol(0).dwValue, Len(vol(1)) * mxcd.cChannels
   
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetPANControl = True
   Else
       SetPANControl = False
   End If
   
End Function

Function unSetMuteControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal unmute As Long) As Boolean
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.Item = 0
   mxcd.cbDetails = Len(vol)
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   vol.dwValue = unmute
   CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       unSetMuteControl = True
   Else
       unSetMuteControl = False
   End If
End Function


Function SetMuteControl(ByVal hmixer As Long, mxc As MIXERCONTROL, mute As Boolean) As Boolean
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.Item = 0
   mxcd.cbDetails = Len(vol)
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   vol.dwValue = volume
   CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetMuteControl = True
   Else
       SetMuteControl = False
   End If
End Function

Function GetVolumeControlValue(ByVal hmixer As Long, mxc As MIXERCONTROL) As Long

    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    mxcd.cbStruct = Len(mxcd)
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cChannels = 1
    mxcd.Item = 0
    mxcd.cbDetails = Len(vol)
    mxcd.paDetails = 0
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr vol, mxcd.paDetails, Len(vol)
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
       GetVolumeControlValue = vol.dwValue
    Else
        GetVolumeControlValue = -1
    End If
End Function

Sub LoadFile(inFile As String)

   Dim hmmioIn As Long
   Dim mmioinf As mmioinfo
   fFileLoaded = False
   If (inFile = "") Then
       GlobalFree (hmem)
       Exit Sub
   End If
   
   hmmioIn = mmioOpen(inFile, mmioinf, MMIO_READ)
   If hmmioIn = 0 Then
       MsgBox "Error opening input file, rc = " & mmioinf.wErrorRet
       Exit Sub
   End If
   
   
   mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
   rc = mmioDescendParent(hmmioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
   If (rc <> 0) Then
       rc = mmioClose(hmmioOut, 0)
       MsgBox "Not a wave file"
       Exit Sub
   End If
   
   
   mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("fmt", 0)
   rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
   If (rc <> 0) Then
       rc = mmioClose(hmmioOut, 0)
       MsgBox "Couldn't get format chunk"
       Exit Sub
   End If
   rc = mmioReadFormat(hmmioIn, formatA, Len(formatA))
   If (rc = -1) Then
      rc = mmioClose(hmmioOut, 0)
      MsgBox "Error reading format"
      Exit Sub
   End If
   rc = mmioAscend(hmmioIn, mmckinfoSubchunkIn, 0)
   
   
   mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
   rc = mmioDescend(hmmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
   If (rc <> 0) Then
      rc = mmioClose(hmmioOut, 0)
      MsgBox "Couldn't get data chunk"
      Exit Sub
   End If
   
   
   GlobalFree hmem
   hmem = GlobalAlloc(&H40, mmckinfoSubchunkIn.ckSize)
   bufferIn = GlobalLock(hmem)
   rc = mmioRead(hmmioIn, bufferIn, mmckinfoSubchunkIn.ckSize)
   
   numSamples = mmckinfoSubchunkIn.ckSize / formatA.nBlockAlign
   
   
   rc = mmioClose(hmmioOut, 0)
   
   fFileLoaded = True
    
End Sub

Sub GetStereo16Sample(ByVal sample As Long, ByRef LeftVol As Double, ByRef rightvol As Double)

   Dim sample16 As Integer
   Dim ptr As Long
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample16, ptr, 2
   LeftVol = sample16 / 32768
   CopyStructFromPtr sample16, ptr + 2, 2
   rightvol = sample16 / 32768

End Sub

Sub GetStereo8Sample(ByVal sample As Long, ByRef LeftVol As Double, ByRef rightvol As Double)

   Dim sample8 As Byte
   Dim ptr As Long
   
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample8, ptr, 1
   LeftVol = (sample8 - 128) / 128
   CopyStructFromPtr sample8, ptr + 1, 1
   rightvol = (sample8 - 128) / 128

End Sub

Sub GetMono16Sample(ByVal sample As Long, ByRef LeftVol As Double)

   Dim sample16 As Integer
   Dim ptr As Long
   
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample16, ptr, 2
   LeftVol = sample16 / 32768

End Sub

Sub GetMono8Sample(ByVal sample As Long, ByRef LeftVol As Double)

   Dim sample8 As Byte
   Dim ptr As Long
   
   ptr = sample * formatA.nBlockAlign + bufferIn
   CopyStructFromPtr sample8, ptr, 1
   LeftVol = (sample8 - 128) / 128

End Sub

Public Function CheckFileVersion(FilenameAndPath As Variant) As Variant
    On Error GoTo HandelCheckFileVersionError
    Dim lDummy As Long, lsize As Long, rc As Long
    Dim lVerbufferLen As Long, lVerPointer As Long
    Dim sBuffer() As Byte
    Dim udtVerBuffer As VS_FIXEDFILEINFO
    Dim ProdVer As String
    lsize = GetFileVersionInfoSize(FilenameAndPath, lDummy)
    If lsize < 1 Then Exit Function
    ReDim sBuffer(lsize)
    rc = GetFileVersionInfo(FilenameAndPath, 0&, lsize, sBuffer(0))
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    
    
    ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl)
    CheckFileVersion = ProdVer
    Exit Function
HandelCheckFileVersionError:
    CheckFileVersion = "N/A"
    Exit Function

End Function



