VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form frmAbout 
   Caption         =   "Sobre Millenium CD Player"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _cx             =   4202585
      _cy             =   4200045
      Movie           =   "d:\VB\Millenium Player Pro\Millenium Player Pro2\walk.swf"
      Src             =   "d:\VB\Millenium Player Pro\Millenium Player Pro2\walk.swf"
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
Attribute VB_Name = "frmAbout"
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
Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)
    Select Case command
    
    End Select
End Sub
Private Sub Form_Load()
ShockwaveFlash1.Movie = App.Path & "\walk.swf"
End Sub
