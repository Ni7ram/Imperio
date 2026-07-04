VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Principal 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   Picture         =   "Principal.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   14370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Height          =   1095
      Left            =   5520
      Picture         =   "Principal.frx":1A3142
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   3135
   End
   Begin VB.CommandButton cmdOpciones 
      Height          =   1095
      Left            =   5520
      Picture         =   "Principal.frx":1B04D4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CommandButton Opciones 
      Height          =   1095
      Left            =   5520
      Picture         =   "Principal.frx":1BF612
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   3135
   End
   Begin WMPLibCtl.WindowsMediaPlayer Menu 
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
      URL             =   ""
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
      _cx             =   2355
      _cy             =   1296
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cardos Games - 2006"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12600
      TabIndex        =   3
      Top             =   8520
      Width           =   1575
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Sub cmdOpciones_Click()
    MenuOpciones.Show vbModal
End Sub

Private Sub cmdOpciones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdOpciones.Picture = LoadPicture(App.Path & "\Menu\Opciones2.GRAF")
End Sub

Private Sub cmdSalir_Click()
    End
End Sub

Private Sub cmdSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdSalir.Picture = LoadPicture(App.Path & "\Menu\Salir2.GRAF")
End Sub

Private Sub Form_Load()
    Menu.URL = (App.Path & "\Musica\Musica.wpl")
    Menu.Controls.Next
    Menu.Controls.Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdSalir.Picture = LoadPicture(App.Path & "\Menu\Salir.GRAF")
    cmdOpciones.Picture = LoadPicture(App.Path & "\Menu\Opciones.GRAF")
    Opciones.Picture = LoadPicture(App.Path & "\Menu\Jugar.GRAF")
End Sub

Private Sub Opciones_Click()
    Menu.URL = ""
    frmJuego.Show
    Unload MenuOpciones
    Principal.Hide
End Sub

Private Sub Opciones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Opciones.Picture = LoadPicture(App.Path & "\Menu\Jugar2.GRAF")
End Sub
