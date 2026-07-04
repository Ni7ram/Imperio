VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMusica 
   Caption         =   "Form1"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1125
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin WMPLibCtl.WindowsMediaPlayer Musica 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      URL             =   "C:\Documents and Settings\martin\Escritorio\(N!7RaM)\Música\Ska-p\Ska-P - 2002 - ¡¡ Que corra la voz !!\05. El Niño Soldado.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2778
      _cy             =   2143
   End
End
Attribute VB_Name = "frmMusica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Musica.URL = (App.Path & "\Musica\Musica.wpl")
End Sub

