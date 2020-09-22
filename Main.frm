VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form Main 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Millenium CD Player"
   ClientHeight    =   4605
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6735
   DrawWidth       =   4
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame help 
      Caption         =   "Opcional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   46
      Top             =   2640
      Width           =   2870
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
         Height          =   2175
         Left            =   0
         TabIndex        =   47
         Top             =   -240
         Width           =   2895
         _cx             =   4199410
         _cy             =   4198140
         Movie           =   "d:\vb\Millenium Player Pro\Millenium Player Pro2\walk.swf"
         Src             =   "d:\vb\Millenium Player Pro\Millenium Player Pro2\walk.swf"
         WMode           =   "Window"
         Play            =   0   'False
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "MIXER"
      Height          =   210
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame6 
      Caption         =   "Microfone"
      Height          =   2325
      Left            =   4680
      TabIndex        =   15
      Top             =   0
      Width           =   975
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   795
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   20
         Top             =   1800
         Width           =   210
      End
      Begin VB.PictureBox Picture4 
         Height          =   330
         Left            =   375
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   17
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option8 
            BackColor       =   &H00808080&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00400000&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   1230
         Left            =   0
         TabIndex        =   23
         Top             =   525
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cd Vol"
      Height          =   2325
      Left            =   3960
      TabIndex        =   8
      Top             =   0
      Width           =   735
      Begin VB.PictureBox Picture7 
         Height          =   330
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   12
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option14 
            BackColor       =   &H00808080&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton Option13 
            BackColor       =   &H00400000&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   1800
         Width           =   210
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   1230
         Left            =   30
         TabIndex        =   10
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Volume"
      Height          =   2325
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   825
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   7
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Height          =   330
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   4
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton Option1 
            BackColor       =   &H00808080&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00400000&
            Height          =   150
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox txtMasterVolume 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   675
      End
      Begin ComctlLib.Slider sliderMasterVolume 
         Height          =   1230
         Left            =   30
         TabIndex        =   3
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider Slider19 
         Height          =   225
         Left            =   60
         TabIndex        =   34
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -32767
         Max             =   32767
         TickStyle       =   3
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6240
      Top             =   2880
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H8000000A&
      Caption         =   "Control"
      Height          =   2325
      Left            =   5760
      TabIndex        =   16
      Top             =   0
      Width           =   855
      Begin VB.CheckBox Check11 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   21
         ToolTipText     =   "Links all selected faders to SBM fader"
         Top             =   2040
         Width           =   225
      End
      Begin ComctlLib.Slider Slider6 
         Height          =   1740
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3069
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   1
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Selec"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   2040
         Width           =   435
      End
   End
   Begin VB.Frame Frame14 
      Height          =   885
      Left            =   0
      TabIndex        =   26
      Top             =   840
      Width           =   2870
      Begin VB.TextBox timeWindow 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   190
         Width           =   2745
      End
   End
   Begin VB.Frame Frame16 
      Height          =   800
      Left            =   0
      TabIndex        =   28
      Top             =   1800
      Width           =   2865
      Begin VB.TextBox CD 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Millenium CD Player"
         Top             =   480
         Width           =   2745
      End
      Begin VB.Label tracktime 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   60
         TabIndex        =   29
         Top             =   150
         Width           =   2745
      End
   End
   Begin VB.Frame Frame17 
      Height          =   810
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   2865
      Begin VB.Label totalplay 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   60
         TabIndex        =   32
         Top             =   150
         Width           =   2745
      End
      Begin VB.Label trtime 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   60
         TabIndex        =   31
         Top             =   480
         Width           =   2745
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Painel de Controle"
      Height          =   2160
      Left            =   2880
      TabIndex        =   35
      Top             =   2530
      Width           =   3840
      Begin Threed.SSCommand Command1 
         Height          =   375
         Left            =   2520
         TabIndex        =   45
         Top             =   1560
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Sair"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   4
      End
      Begin Threed.SSCommand eject1 
         Height          =   375
         Left            =   1320
         TabIndex        =   44
         Top             =   1560
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Fechar CD"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   4
      End
      Begin Threed.SSCommand eject0 
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1560
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Abrir CD"
         ForeColor       =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   4
      End
      Begin Threed.SSCommand ff 
         Height          =   375
         Left            =   2040
         TabIndex        =   42
         ToolTipText     =   "Avançar dentro da Faixa"
         Top             =   960
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   4
         Picture         =   "Main.frx":08CA
      End
      Begin Threed.SSCommand rew 
         Height          =   375
         Left            =   3000
         TabIndex        =   41
         ToolTipText     =   "Recuar dentro da Faixa"
         Top             =   960
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   4
         Picture         =   "Main.frx":104C
      End
      Begin Threed.SSCommand btrack 
         Height          =   375
         Left            =   1080
         TabIndex        =   40
         ToolTipText     =   "Recuar Faixa"
         Top             =   960
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   4
         Picture         =   "Main.frx":17C7
      End
      Begin Threed.SSCommand ftrack 
         Height          =   375
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "Avançar Faixa"
         Top             =   960
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   4
         Picture         =   "Main.frx":1EED
      End
      Begin Threed.SSCommand play 
         Height          =   375
         Left            =   360
         TabIndex        =   36
         ToolTipText     =   "Play"
         Top             =   360
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   8388608
         BevelWidth      =   4
         Picture         =   "Main.frx":2614
      End
      Begin Threed.SSCommand pause 
         Height          =   375
         Left            =   1440
         TabIndex        =   37
         ToolTipText     =   "Pause"
         Top             =   360
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   78
         ForeColor       =   -2147483633
         BevelWidth      =   4
         Picture         =   "Main.frx":2A04
      End
      Begin Threed.SSCommand stopbtn 
         Height          =   375
         Left            =   2640
         TabIndex        =   38
         ToolTipText     =   "Stop"
         Top             =   360
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   78
         BevelWidth      =   4
         Picture         =   "Main.frx":316E
      End
   End
   Begin VB.Menu mnusobre 
      Caption         =   "&Menu"
      Begin VB.Menu mnusobremillenium 
         Caption         =   "Sobre"
      End
      Begin VB.Menu mnusair 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Millenium CD Player
'OFPG - Marketing e Publicidade, Lda - Brain Design Solutions
'http://www.braindesignsolutions.com
'softwaredirector@braindesignsolutions.com
'Agosto 2001
'Made in Portugal (UE)
'-------------------------------------------
Option Explicit
Dim volR As Long
Dim volL As Long
Dim volume As Long
Dim mute As MIXERCONTROL
Dim unmute As MIXERCONTROL
Dim hmixer As Long
Dim VolCtrl As MIXERCONTROL
Dim WavCtrl As MIXERCONTROL
Dim CDVol As MIXERCONTROL
Dim LineVol As MIXERCONTROL
Dim MBOOST As MIXERCONTROL
Dim PSPKVol As MIXERCONTROL

Dim AUXVol As MIXERCONTROL
Dim TADVol As MIXERCONTROL

Dim MIDIVol As MIXERCONTROL

Dim I25InVol As MIXERCONTROL
Dim Treble As MIXERCONTROL
Dim Bass As MIXERCONTROL

Dim rc As Long
Dim ok As Boolean




Dim fastForwardSpeed As Long
Dim fPlaying As Boolean
Dim fCDLoaded As Boolean
Dim numTracks As Integer
Dim trackLength() As String
Dim track As Integer
Dim min As Integer
Dim SEC As Integer
Dim cmd As String


Private Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200

rc = mciSendString(cmd, 0, 0, hWnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    SendMCIString "close all", False
    cmd = "close all"
    SendMCIString cmd, True
    Unload Main
End If
SendMCIString = (rc = 0)
End Function

Private Sub Check11_Click()
If Check11.Value = 1 Then
Check1.Value = 1
Check2.Value = 1
Check4.Value = 1
End If
If Check11.Value = 0 Then
Check1.Value = 0
Check2.Value = 0
Check4.Value = 0
End If
End Sub



Private Sub Command1_Click()
SendMCIString "close all", False
cmd = "close all"
SendMCIString cmd, True
'Index.Enabled = True
Unload Main
End Sub

Private Sub Command2_Click()
    'Open the mixer with deviceID 0.
    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Não é possível abrir a mesa de mistura. Certifique que está correctamente instalada."
        Exit Sub
    End If

    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, VolCtrl)
    If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, VolCtrl)
        If volume <> -1 Then
            txtMasterVolume.Text = volume \ 6553
            sliderMasterVolume.Value = 65535 - volume
        End If
    End If
   
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, WavCtrl)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, WavCtrl)
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MBOOST, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, MBOOST)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, MBOOST)
        If volume <> -1 Then
            Text2.Text = volume \ 6553
            Slider2.Value = 65535 - volume
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, CDVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, CDVol)
        If volume <> -1 Then
            Text1.Text = volume \ 6553
            Slider1.Value = 65535 - volume
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_src_AUXVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, AUXVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, AUXVol)
        If volume <> -1 Then
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TADVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, TADVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, TADVol)
        If volume <> -1 Then
         
        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, MIDIVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, MIDIVol)
        If volume <> -1 Then

        End If
    End If

        ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PSPKVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, PSPKVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, PSPKVol)
        If volume <> -1 Then

        End If
    End If

    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, I25InVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, I25InVol)
        If volume <> -1 Then

        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, LineVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, LineVol)
        If volume <> -1 Then

        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_BASS, Bass)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, Bass)
        If volume <> -1 Then

        End If
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_TREBLE, Treble)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, Treble)
        If volume <> -1 Then

        End If
    End If
End Sub

Private Sub Form_Load()
ShockwaveFlash1.Movie = App.Path & "\walk.swf"

If (App.PrevInstance = True) Then
    End
End If

Timer1.Enabled = False
fastForwardSpeed = 5
fCDLoaded = False


If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
    timeWindow.Text = "Cd in use"
    End
End If
SendMCIString "set cd time format tmsf wait", True
Timer1.Enabled = True
Command2_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)

SendMCIString "close all", False
End Sub


Private Sub mnusair_Click()
End
End Sub

Private Sub mnusobremillenium_Click()
frmAbout.Show
End Sub

Private Sub Option1_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option10_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option17_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PSPKVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option18_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PSPKVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option19_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option20_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option9_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option11_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TADVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option12_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TADVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option13_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option14_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option15_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_src_AUXVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option16_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_src_AUXVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option2_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option3_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option4_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
unSetMuteControl hmixer, mute, 1
End Sub

Private Sub Option5_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Private Sub Option6_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option7_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MBOOST, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Option8_Click()
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MBOOST, _
                                  MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub

Function Errora()
    MsgBox "Your sound card does not support a bass control"
End Function

Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)
    Select Case command
    
    End Select
End Sub



Private Sub Slider19_Scroll()
    volL = CLng(32767.5 - Slider19)
    volR = CLng(32767.5 + Slider19)
    SetPANControl hmixer, VolCtrl, volL, volR
End Sub

Private Sub trebleslider_Scroll()

      ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_TREBLE, Treble)
    If ok = False Then
    Errora
    Exit Sub
    End If
    SetVolumeControl hmixer, Treble, volume
End Sub

Private Sub BassSlider_Scroll()
    SetVolumeControl hmixer, Bass, volume
End Sub



Private Sub Play_Click()
SendMCIString "play cd", True
fPlaying = True
CD.Text = "A Reproduzir"
End Sub

Private Sub Slider6_scroll()
Dim link As Long
link = 65535 - CLng(Slider6.Value)
If Check2.Value = 1 Then
Slider1.Value = Slider6.Value
Text1.Text = link \ 6553
SetVolumeControl hmixer, CDVol, link
End If
If Check4.Value = 1 Then
Slider2.Value = Slider6.Value
Text2.Text = link \ 6553
SetVolumeControl hmixer, MBOOST, link
End If
If Check1.Value = 1 Then
sliderMasterVolume.Value = Slider6.Value
txtMasterVolume.Text = link \ 6553
SetVolumeControl hmixer, VolCtrl, link
End If

End Sub

Private Sub SSCommand1_Click()

End Sub


Private Sub stopbtn_Click()
SendMCIString "stop cd wait", True
cmd = "seek cd to " & track
SendMCIString cmd, True
fPlaying = False
CD.Text = "CD Parado"
Update
End Sub

Private Sub pause_Click()
SendMCIString "pause cd", True
fPlaying = False
CD.Text = "Cd em Pausa"
Update
End Sub


Private Sub Eject0_Click()
SendMCIString "set cd door open", True
CD.Text = "Insira CD"
eject1.Visible = True
eject0.Visible = False
Update
End Sub
Private Sub Eject1_Click()
CD.Text = "Aguarde Por Favor"
SendMCIString "set cd door closed", True
eject0.Visible = True
eject1.Visible = False
Update
End Sub

Private Sub ff_Click()
Dim s As String * 40

SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", s, Len(s), 0
If (fPlaying) Then
    cmd = "play cd from " & CStr(CLng(s) + fastForwardSpeed * 1000)
Else
    cmd = "seek cd to " & CStr(CLng(s) + fastForwardSpeed * 1000)
End If
mciSendString cmd, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub

Private Sub rew_Click()
Dim s As String * 40

SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", s, Len(s), 0
If (fPlaying) Then
    cmd = "play cd from " & CStr(CLng(s) - fastForwardSpeed * 1000)
Else
    cmd = "seek cd to " & CStr(CLng(s) - fastForwardSpeed * 1000)
End If
mciSendString cmd, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub

Private Sub ftrack_Click()
If (track < numTracks) Then
    If (fPlaying) Then
        cmd = "play cd from " & track + 1
        SendMCIString cmd, True
    Else
        cmd = "seek cd to " & track + 1
        SendMCIString cmd, True
    End If
Else
    SendMCIString "seek cd to 1", True
End If
Update
End Sub

Private Sub btrack_Click()
Dim from As String
If (min = 0 And SEC = 0) Then
    If (track > 1) Then
        from = CStr(track - 1)
    Else
        from = CStr(numTracks)
    End If
Else
    from = CStr(track)
End If
If (fPlaying) Then
    cmd = "play cd from " & from
    SendMCIString cmd, True
Else
    cmd = "seek cd to " & from
    SendMCIString cmd, True
End If
Update
End Sub

Private Sub Update()
Static s As String * 30


mciSendString "status cd media present", s, Len(s), 0
If (CBool(s)) Then

    If (fCDLoaded = False) Then
        mciSendString "status cd number of tracks wait", s, Len(s), 0
        numTracks = CInt(Mid$(s, 1, 2))
        eject0.Visible = True
        eject1.Visible = False
        CD.Text = "CD Pronto"

        If (numTracks = 1) Then
            CD.Text = "Formato Inválido"
            Exit Sub
        End If
        
        mciSendString "status cd length wait", s, Len(s), 0
        totalplay.Caption = "Faixas: " & numTracks
        trtime.Caption = "Tempo Total: " & s
        ReDim trackLength(1 To numTracks)
        Dim i As Integer
        For i = 1 To numTracks
            cmd = "status cd length track " & i
            mciSendString cmd, s, Len(s), 0
            trackLength(i) = s
        Next
        timeWindow.FontSize = 18
        play.Enabled = True
        pause.Enabled = True
        ff.Enabled = True
        rew.Enabled = True
        ftrack.Enabled = True
        btrack.Enabled = True
        stopbtn.Enabled = True
        fCDLoaded = True
        SendMCIString "seek cd to 1", True
    End If

    mciSendString "status cd position", s, Len(s), 0
    track = CInt(Mid$(s, 1, 2))
    min = CInt(Mid$(s, 4, 2))
    SEC = CInt(Mid$(s, 7, 2))
    timeWindow.Text = "[" & Format(track, "00") & "] " & Format(min, "00") _
            & ":" & Format(SEC, "00")
    tracktime.Caption = "Tempo Faixa: " & trackLength(track)

    mciSendString "status cd mode", s, Len(s), 0
    fPlaying = (Mid$(s, 1, 7) = "playing")
Else

    If (fCDLoaded = True) Then
        play.Enabled = False
        pause.Enabled = False
        ff.Enabled = False
        rew.Enabled = False
        ftrack.Enabled = False
        btrack.Enabled = False
        stopbtn.Enabled = False
        fCDLoaded = False
        fPlaying = False
        totalplay.Caption = ""
        tracktime.Caption = ""
        CD.Text = "Não existe CD no Leitor"
    End If
End If
End Sub

Private Sub Timer1_Timer()
Update
End Sub
Private Sub Slider1_Scroll()
    volume = 65535 - CLng(Slider1.Value)
    Text1.Text = volume \ 6553
    SetVolumeControl hmixer, CDVol, volume
End Sub
Private Sub Slider2_Scroll()
    volume = 65535 - CLng(Slider2.Value)
    Text2.Text = volume \ 6553
    SetVolumeControl hmixer, MBOOST, volume
End Sub
Private Sub Slider3_Scroll()
    SetVolumeControl hmixer, AUXVol, volume
End Sub
Private Sub Slider4_Scroll()
    SetVolumeControl hmixer, TADVol, volume
End Sub
Private Sub Slider5_Scroll()
    SetVolumeControl hmixer, MIDIVol, volume
End Sub
Private Sub Slider7_Scroll()
    SetVolumeControl hmixer, PSPKVol, volume
End Sub
Private Sub Slider8_Scroll()
    SetVolumeControl hmixer, I25InVol, volume
End Sub
Private Sub Slider9_Scroll()
    SetVolumeControl hmixer, LineVol, volume
End Sub
Private Sub sliderMasterVolume_Scroll()
    volume = 65535 - CLng(sliderMasterVolume.Value)
    txtMasterVolume.Text = volume \ 6553
    SetVolumeControl hmixer, VolCtrl, volume
End Sub
Private Sub sliderWaveOutVolume_Scroll()
    SetVolumeControl hmixer, WavCtrl, volume
End Sub



