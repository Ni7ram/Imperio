VERSION 5.00
Begin VB.Form MenuOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Recursos"
      Height          =   3015
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox Arboles 
         Height          =   315
         ItemData        =   "MenuOpciones.frx":0000
         Left            =   2880
         List            =   "MenuOpciones.frx":0013
         TabIndex        =   13
         Text            =   "Normal"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cantMetal 
         Height          =   315
         ItemData        =   "MenuOpciones.frx":0044
         Left            =   2880
         List            =   "MenuOpciones.frx":0066
         TabIndex        =   11
         Text            =   "6"
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox cantPiedra 
         Height          =   315
         ItemData        =   "MenuOpciones.frx":0089
         Left            =   2880
         List            =   "MenuOpciones.frx":00AB
         TabIndex        =   9
         Text            =   "6"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Arboles"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad de canteras de metal"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad de canteras de piedra"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.CheckBox Musica 
      Caption         =   "Musica"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   480
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox Ambiente 
      Caption         =   "Sonidos de ambiente"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Graficos"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
      Begin VB.CheckBox FondoOpt 
         Caption         =   "Desactivar fondo (para PCs lentas)"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Audio"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "MenuOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Musica.Value = 0 Then
        AudioMusica = True
    Else
       AudioMusica = False
    End If
    If Ambiente.Value = 0 Then
        AudioSonidos = True
    Else
        AudioSonidos = False
    End If
    If FondoOpt.Value = 1 Then
        Fondo = True
    Else
        Fondo = False
    End If
    Select Case Arboles
        Case "Pocos"
            cantArboles = 30
        Case "Normal"
            cantArboles = 50
        Case "Muchos"
            cantArboles = 100
        Case "Exagerado"
            cantArboles = 250
        Case "Imposible"
            cantArboles = 1000
    End Select
    maxCanteras = cantPiedra
    maxMetales = cantMetal
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Fondo_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Hide
End Sub
