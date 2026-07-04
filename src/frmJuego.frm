VERSION 5.00
Begin VB.Form frmJuego 
   BackColor       =   &H00004000&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "frmJuego.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmJuego.frx":030A
   MousePointer    =   99  'Custom
   Picture         =   "frmJuego.frx":0BD4
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "Pausa"
      Height          =   495
      Left            =   14280
      TabIndex        =   51
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdObjetivos 
      Height          =   495
      Left            =   11160
      Picture         =   "frmJuego.frx":240C16
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Son1 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   10440
      Picture         =   "frmJuego.frx":242260
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Son2 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   10440
      Picture         =   "frmJuego.frx":242B92
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton GuardarTrigo 
      Caption         =   "Almacenar 10 gavillas de trigo"
      Height          =   615
      Left            =   9840
      TabIndex        =   40
      Top             =   10800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton consMetalero 
      Height          =   495
      Left            =   2640
      Picture         =   "frmJuego.frx":2434C4
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   10920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000007&
      Height          =   855
      Left            =   13200
      Picture         =   "frmJuego.frx":253012
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton consCorazas 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   840
      Picture         =   "frmJuego.frx":25735C
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consBarracas 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1440
      Picture         =   "frmJuego.frx":26779E
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consArqueria 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2040
      Picture         =   "frmJuego.frx":27E340
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consArmeria 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   240
      Picture         =   "frmJuego.frx":296316
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000007&
      Height          =   855
      Left            =   13200
      Picture         =   "frmJuego.frx":2A6B64
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton consSerreria 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   840
      Picture         =   "frmJuego.frx":2A87BE
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consRefineria 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1440
      Picture         =   "frmJuego.frx":2B7C08
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   10920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consGranero 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2040
      Picture         =   "frmJuego.frx":2C917E
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consMercado 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   240
      Picture         =   "frmJuego.frx":2F013C
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   10920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000007&
      Height          =   855
      Left            =   14280
      Picture         =   "frmJuego.frx":310BC2
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000007&
      Height          =   855
      Left            =   13200
      Picture         =   "frmJuego.frx":3130C4
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Height          =   855
      Left            =   14280
      Picture         =   "frmJuego.frx":314FDE
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton consFuente 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   3240
      Picture         =   "frmJuego.frx":3173E0
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consCasa 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   240
      Picture         =   "frmJuego.frx":3188EA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consCamino 
      Height          =   495
      Left            =   2640
      Picture         =   "frmJuego.frx":31AB4C
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consGranja 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2040
      Picture         =   "frmJuego.frx":31BFC2
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAtras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   3240
      Picture         =   "frmJuego.frx":32E68C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   10920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consPiedra 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1440
      Picture         =   "frmJuego.frx":32F0DE
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton consLeña 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   840
      Picture         =   "frmJuego.frx":3400A0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdOpciones 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   14280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmJuego.frx":34DE0A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdHistorial 
      BackColor       =   &H80000007&
      Height          =   855
      Left            =   13200
      Picture         =   "frmJuego.frx":350014
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H80000007&
      Height          =   855
      Left            =   14280
      Picture         =   "frmJuego.frx":351A56
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdCrear 
      Height          =   855
      Left            =   240
      Picture         =   "frmJuego.frx":3538F8
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   10440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame fraRecursos 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10335
      Begin VB.Label lblOro 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6600
         TabIndex        =   38
         Top             =   120
         Width           =   735
      End
      Begin VB.Label LimiteCiudadanos 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   9720
         TabIndex        =   11
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblCiudadanos 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   9000
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblRecursos 
         BackStyle       =   0  'Transparent
         Caption         =   "      Trigo:                    Madera:                    Piedra:                   Oro:                  Cuidadanos:          /"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   9735
      End
      Begin VB.Label lblComida 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblMadera 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblPiedra 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5040
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CommandButton Flechas 
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   13920
      Picture         =   "frmJuego.frx":355BF6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10080
      Width           =   495
   End
   Begin VB.CommandButton Flechas 
      Height          =   495
      Index           =   1
      Left            =   14640
      Picture         =   "frmJuego.frx":355DB8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10440
      Width           =   495
   End
   Begin VB.CommandButton Flechas 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   13200
      Picture         =   "frmJuego.frx":355F7A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10440
      Width           =   495
   End
   Begin VB.CommandButton Flechas 
      Height          =   495
      Index           =   2
      Left            =   13920
      Picture         =   "frmJuego.frx":35613C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10800
      Width           =   495
   End
   Begin VB.Frame frmInfo 
      BackColor       =   &H80000007&
      Caption         =   $"frmJuego.frx":3562FE
      ForeColor       =   &H8000000E&
      Height          =   1455
      Left            =   0
      TabIndex        =   41
      Top             =   10080
      Width           =   11535
      Begin VB.Frame Division 
         BackColor       =   &H80000007&
         Height          =   2055
         Left            =   3960
         TabIndex        =   44
         Top             =   0
         Width           =   135
      End
      Begin VB.CommandButton Construir 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   360
         Picture         =   "frmJuego.frx":3563E9
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdMilitar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   1560
         Picture         =   "frmJuego.frx":357DC3
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblVidaFijo 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   5280
         TabIndex        =   47
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Obj 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   5280
         TabIndex        =   46
         Top             =   240
         Width           =   6135
      End
      Begin VB.Image ima 
         Height          =   855
         Left            =   4200
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblVida 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6120
         TabIndex        =   45
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Image Metalero 
      Height          =   2025
      Index           =   0
      Left            =   5040
      Picture         =   "frmJuego.frx":35B435
      Top             =   -3000
      Width           =   2265
   End
   Begin VB.Image Serreria 
      Height          =   1995
      Index           =   0
      Left            =   5520
      Picture         =   "frmJuego.frx":36AF83
      Top             =   -3000
      Width           =   2265
   End
   Begin VB.Image Refineria 
      Height          =   2355
      Index           =   0
      Left            =   7080
      Picture         =   "frmJuego.frx":37A719
      Top             =   -3000
      Width           =   2250
   End
   Begin VB.Image Mercado 
      Height          =   2865
      Index           =   0
      Left            =   240
      Picture         =   "frmJuego.frx":38C8DB
      Top             =   -3000
      Width           =   3495
   End
   Begin VB.Image Granero 
      Height          =   3450
      Index           =   0
      Left            =   2040
      Picture         =   "frmJuego.frx":3AEB49
      Top             =   -3500
      Width           =   2295
   End
   Begin VB.Image Corazas 
      Height          =   2040
      Index           =   0
      Left            =   6960
      Picture         =   "frmJuego.frx":3C9AD3
      Top             =   -3000
      Width           =   2625
   End
   Begin VB.Image Barracas 
      Height          =   2280
      Index           =   0
      Left            =   1440
      Picture         =   "frmJuego.frx":3DC05D
      Top             =   -3000
      Width           =   3045
   End
   Begin VB.Image Arqueria 
      Height          =   2505
      Index           =   0
      Left            =   4200
      Picture         =   "frmJuego.frx":3E4E67
      Top             =   -3000
      Width           =   2940
   End
   Begin VB.Image Armeria 
      Height          =   2040
      Index           =   0
      Left            =   4560
      Picture         =   "frmJuego.frx":3FE089
      Top             =   -3000
      Width           =   2625
   End
   Begin VB.Image CanteraMetal 
      Height          =   1200
      Index           =   0
      Left            =   9360
      Picture         =   "frmJuego.frx":410613
      Top             =   -3000
      Width           =   1995
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   9720
      Width           =   11295
   End
   Begin VB.Label Horizontal 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Vertical 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image PicaPiedra 
      Height          =   2190
      Index           =   0
      Left            =   2160
      Picture         =   "frmJuego.frx":41899D
      Top             =   -3000
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Image Granja 
      Height          =   1860
      Index           =   0
      Left            =   4200
      Picture         =   "frmJuego.frx":42995F
      Tag             =   "100"
      Top             =   -3000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image Leñador 
      Height          =   1830
      Index           =   0
      Left            =   2400
      Picture         =   "frmJuego.frx":43C029
      Top             =   -3000
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image CentroCiudad 
      Height          =   1875
      Left            =   6120
      Picture         =   "frmJuego.frx":449D93
      Top             =   4800
      Width           =   2235
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   0
      Top             =   11040
      Width           =   9855
   End
   Begin VB.Image Arbol 
      Height          =   2265
      Index           =   0
      Left            =   4920
      Picture         =   "frmJuego.frx":458261
      Tag             =   "100"
      Top             =   -3000
      Width           =   1185
   End
   Begin VB.Image Fuente 
      Height          =   1005
      Index           =   0
      Left            =   4440
      Picture         =   "frmJuego.frx":45BCEF
      Top             =   -3000
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image Casa 
      Height          =   1665
      Index           =   0
      Left            =   4200
      Picture         =   "frmJuego.frx":45D1F9
      Top             =   -3000
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image CanteraPiedra 
      Height          =   1185
      Index           =   0
      Left            =   11640
      Picture         =   "frmJuego.frx":46607B
      Top             =   -3000
      Width           =   1590
   End
   Begin VB.Image Camino 
      Height          =   675
      Index           =   0
      Left            =   5280
      Picture         =   "frmJuego.frx":46C875
      Top             =   -3000
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmJuego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------
'Proyecto: Juego
'
'Nombre: Imperio ( a cambiar )       (13/Mar/06)
'Tipo: Estrategia
'
'Revisado/mejorado: el 23/Mar/06

'©Martin Cardozo '*(N!7RaM)*'
'----------------------------------------------

Option Explicit
    Dim CostoPiedra#
    Dim booleanCasa As Boolean
    Dim Creacion As String 'Para poner en el historial que se creo y sumar a la plantilla
    Dim limitePobl#
    Dim mercadoBoolean As Boolean

Private Sub Arbol_Click(Index As Integer)
    Obj = "Arbol: puedes talarlos para extraer madera."
    lblVidaFijo = "Materia prima: "
    lblVidaFijo.Visible = True
    lblVida = "100"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Arbol.GRAF")
End Sub

Private Sub Arbol_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Pino"
End Sub

Private Sub Armeria_Click(Index As Integer)
    Obj = "Armeria: aquí el metal y los tablones se transforman el armas necesarias para la guerra"
    lblVidaFijo = "Vida: "
    lblVida = "300"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Armeria.GRAF")
End Sub

Private Sub Armeria_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Armería"
End Sub

Private Sub Arqueria_Click(Index As Integer)
    Obj = "Arqueria: aqui tus arqueros afinan su puntería para ganar las batallas"
    lblVidaFijo = "Vida: "
    lblVida = "300"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Arqueria.GRAF")
End Sub

Private Sub Arqueria_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Arqueria"
End Sub

Private Sub Barracas_Click(Index As Integer)
    Obj = "Barracas: aqui tus soldados entrenan para ganar las batallas"
    lblVidaFijo = "Vida: "
    lblVida = "300"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Barracas.GRAF")
End Sub

Private Sub Barracas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Barracas"
End Sub

Private Sub Camino_Click(Index As Integer)
    lblVidaFijo = ""
    lblVida = ""
    ima = Nothing
    Obj = "Camino: dan atractivo a tu cuidad y a tus aldeanos les gusta. Cuanto más camino pongas alrededor de un lugar de trabajo más producirá"
End Sub

Private Sub CanteraMetal_Click(Index As Integer)
    Obj = "Cantera de Metal: puedes extraer el metal para fabricar armas"
    lblVidaFijo = "Materia prima: "
    lblVida = "100"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Metal.GRAF")
End Sub

Private Sub CanteraMetal_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Cantera de metal"
End Sub

Private Sub CanteraPiedra_Click(Index As Integer)
    Obj = "Cantera de Piedra: de aquí puedes extraer la piedra."
    lblVidaFijo = "Materia prima: "
    lblVida = "100"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Piedra.GRAF")
End Sub

Private Sub CanteraPiedra_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Cantera de Piedra"
End Sub

Private Sub Casa_Click(Index As Integer)
    Obj = "Casa: en estas cómodas casas viven tus ciudadanos"
    lblVidaFijo = "Vida: "
    lblVidaFijo.Visible = True
    lblVida = "50"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Casa.GRAF")
    Sonidos.Sonidos.URL = App.Path & "\Sonidos\Casa.WAV"
End Sub

Private Sub Casa_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Casa"
End Sub

Private Sub cmdAtras_Click()
    Construir.Visible = True
    consFuente.Visible = False
    consCamino.Visible = False
    consCasa.Visible = False
    consGranja.Visible = False
    consPiedra.Visible = False
    consLeña.Visible = False
    cmdAtras.Visible = False
    consSerreria.Visible = False
    consMercado.Visible = False
    consRefineria.Visible = False
    consGranero.Visible = False
    If Objetivos2 = True Then
        cmdMilitar.Visible = True
    End If
    consBarracas.Visible = False
    consArqueria.Visible = False
    consArmeria.Visible = False
    consCorazas.Visible = False
    consMetalero.Visible = False
End Sub

Private Sub cmdAtras_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Volver"
End Sub

Private Sub cmdCrear_Click()
    If lblComida >= 50 Then
        If lblCiudadanos < limitePobl Then
            lblCiudadanos = lblCiudadanos + 1
            If lblCiudadanos >= 50 Then
                frmObjetivos.Check1.Value = Checked
                frmObjetivos.Check1.Enabled = False
            End If
            lblComida = lblComida - 50
            frmHistorial.Historial.AddItem ("Creaste un ciudadano")
            frmPlantilla.lblTraDis = frmPlantilla.lblTraDis + 1
        Else
            frmHistorial.Historial.AddItem ("Alcanzaste el limite de población")
        End If
    End If
End Sub

Private Sub cmdCrear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblCiudadanos >= limitePobl Then
        lblInfo = "Haz alcanzado el limite de poblacion"
    Else
        lblInfo = "Crear cuidadano: 50 comida"
    End If
End Sub

Private Sub cmdHistorial_Click()
    If frmHistorial.Visible = False Then
        frmHistorial.Visible = True
    Else
        frmHistorial.Visible = False
    End If
End Sub

Private Sub cmdHistorial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Ver Historial"
End Sub

Private Sub cmdMilitar_Click()
    Obj = ""
    lblVidaFijo = ""
    lblVida = ""
    ima.Picture = Nothing
    Construir.Visible = False
    consFuente.Visible = False
    consCamino.Visible = False
    consCasa.Visible = False
    consGranja.Visible = False
    consPiedra.Visible = False
    consLeña.Visible = False
    cmdAtras.Visible = True
    consSerreria.Visible = False
    consMercado.Visible = False
    consRefineria.Visible = False
    consGranero.Visible = False
    
    cmdMilitar.Visible = False
    consBarracas.Visible = True
    consArqueria.Visible = True
    consArmeria.Visible = True
    consCorazas.Visible = True
    GuardarTrigo.Visible = False
End Sub

Private Sub cmdMilitar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Construir edificios militares"
End Sub

Private Sub cmdObjetivos_Click()
    If frmObjetivos.Visible = False Then
        frmObjetivos.Visible = True
    Else
        frmObjetivos.Visible = False
    End If
End Sub

Private Sub cmdObjetivos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Ver objetivos"
End Sub

Private Sub cmdOpciones_Click()
    If frmPlantilla.Visible = False Then 'La plantilla es como el Zeus
        frmPlantilla.Show
    Else
        frmPlantilla.Visible = False
    End If
End Sub

Private Sub cmdReplantar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Cuando la cosecha se acaba puedes replantarla: 50 piedra, 50 madera"
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("Realmente desea salir?", vbYesNo, "Salir") = vbYes Then
        End
    End If
End Sub

Private Sub cmdSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Salir"
End Sub

Private Sub Command1_Click()
    If frmArmeria.Visible = False Then
        frmArmeria.Show vbModal
    Else
        frmArmeria.Visible = False
    End If
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Ventana armeria"
End Sub

Private Sub Command2_Click()
    If frmIndustria.Visible = False Then
        frmIndustria.Visible = True
    Else
        frmIndustria.Visible = False
    End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Ventana Industria"
End Sub

Private Sub Command3_Click()
    If frmEjercito.Visible = False Then
        frmEjercito.Show vbModal
    Else
        frmEjercito.Visible = False
    End If
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Ventana Ejercito"
End Sub

Private Sub Command4_Click()
    If frmAlmacen.Visible = False Then
        frmAlmacen.Visible = True
    Else
        frmAlmacen.Visible = False
    End If
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Ventana Almacen"
End Sub

Private Sub Command5_Click()
    If frmMercado.Visible = False And mercadoBoolean = True Then
        frmMercado.Show vbModal
    Else
        frmMercado.Visible = False
    End If
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mercadoBoolean = True Then
        lblInfo = "Ventana Mercado"
    Else
        lblInfo = "No tienes un mercado"
    End If
End Sub

Private Sub Command6_Click()
    MsgBox "Juego Pausado", , "Pausa"
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Pausa"
End Sub

Private Sub consArmeria_Click()

    '-----CONSTRUIR ARMERIA-------------

    If 3 <= frmAlmacen.lblMadera And 3 <= frmAlmacen.lblBloque Then
        Creacion = "armeria"
        Load Armeria(Armeria.Count)
        With Armeria(Armeria.Count - 1)
            .Move consArmeria.Left + consArmeria.Width, consArmeria.Top + consArmeria.Height, Armeria(0).Width, Armeria(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consArmeria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Armeria: 3 bloques de piedra y 3 tablones de madera"
End Sub

Private Sub consArqueria_Click()

    '-----CONSTRUIR ARQUERIA-------------

    If 4 <= frmAlmacen.lblMadera And 4 <= frmAlmacen.lblBloque Then
        Creacion = "arqueria"
        Load Arqueria(Arqueria.Count)
        With Arqueria(Arqueria.Count - 1)
            .Move consArqueria.Left + consArqueria.Width, consArqueria.Top + consArqueria.Height, Arqueria(0).Width, Arqueria(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consArqueria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Arqueria: 4 bloques de piedra y 4 tablones de madera"
End Sub

Private Sub consBarracas_Click()

    '-----CONSTRUIR BARRACAS-------------

    If 4 <= frmAlmacen.lblMadera And 4 <= frmAlmacen.lblBloque Then
        Creacion = "barracas"
        Load Barracas(Barracas.Count)
        With Barracas(Barracas.Count - 1)
            .Move consBarracas.Left + consBarracas.Width, consBarracas.Top + consBarracas.Height, Barracas(0).Width, Barracas(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consBarracas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Barracas: 4 bloques de piedra y 4 tablones de madera"
End Sub

Private Sub consCamino_Click()

    '-----CONSTRUIR CAMINO-------------

    lblInfo = ""
    If CostoPiedra <= lblPiedra Then
        Creacion = "camino"
        Load Camino(Camino.Count)
        With Camino(Camino.Count - 1)
            .Move consCamino.Left, consCamino.Top, Camino(0).Width, Camino(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
        End With
    End If
End Sub

Private Sub consCamino_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Camino: 5 piedra"
End Sub

Private Sub consCasa_Click()

    '-----CONSTRUIR CASA-------------

    If 3 <= frmAlmacen.lblMadera And 2 <= frmAlmacen.lblBloque Then
        booleanCasa = True
        Creacion = "casa"
        Load Casa(Casa.Count)
        With Casa(Casa.Count - 1)
            .Move consCasa.Left + consCasa.Width, consCasa.Top + consCasa.Height, 1545, 1665
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
        frmHistorial.Historial.AddItem (Now & ":     Extendiste el limite de población 5 ciudadanos mas")
    End If
End Sub

Private Sub consCasa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Casa: 3 tablones de madera y 2 bloques de piedra"
End Sub

Private Sub consCorazas_Click()

    '-----CONSTRUIR CORAZAS-------------

    If 3 <= frmAlmacen.lblMadera And 3 <= frmAlmacen.lblBloque Then
        Creacion = "corazas"
        Load Corazas(Corazas.Count)
        With Corazas(Corazas.Count - 1)
            .Move consCorazas.Left + consCorazas.Width, consCorazas.Top + consCorazas.Height, Corazas(0).Width, Corazas(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consCorazas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Herreria: 3 bloques de piedra y 3 tablones de madera"
End Sub

Private Sub consFuente_Click()

    '-----CONSTRUIR FUENTE-------------

    lblInfo = ""
    If CostoPiedra <= lblPiedra Then
        Creacion = "fuente"
        Load Fuente(Fuente.Count)
        With Fuente(Fuente.Count - 1)
            .Move consFuente.Left, consFuente.Top, Fuente(0).Width, Fuente(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consFuente_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Fuente: 50 piedra"
End Sub

Private Sub consGranero_Click()

    '-----CONSTRUIR GRANERO-------------

    If 3 <= frmAlmacen.lblMadera And 4 <= frmAlmacen.lblBloque Then
        Creacion = "Granero"
        Load Granero(Granero.Count)
        With Granero(Granero.Count - 1)
            .Move consGranero.Left + consGranero.Width, consGranero.Top + consGranero.Height, Granero(0).Width, Granero(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consGranero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Granero: 4 bloques de piedra y 3 tablones de madera"
End Sub

Private Sub consGranja_Click()
    
    '-----CONSTRUIR GRANJA-------------

    If 5 <= frmAlmacen.lblMadera And 5 <= frmAlmacen.lblBloque Then
        Creacion = "granja de trigo"
        Load Granja(Granja.Count)
        With Granja(Granja.Count - 1)
            .Move consGranja.Left + consGranja.Width, consGranja.Top + consGranja.Height, Granja(0).Width, Granja(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consGranja_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Granja de Trigo: 5 tablones de madera y 5 bloques de piedra"
End Sub

Private Sub consLeña_Click()

    '-----CONSTRUIR LEÑADOR-------------
    
    If 2 <= frmAlmacen.lblMadera And 2 <= frmAlmacen.lblBloque Then
        Creacion = "casa de leñador"
        Load Leñador(Leñador.Count)
        With Leñador(Leñador.Count - 1)
            .Move consLeña.Left + consLeña.Width, consLeña.Top + consLeña.Height, 2205, 1830
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consLeña_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Casa de Leñador: 2 tablones de madera y 2 bloques de piedra"
End Sub

Private Sub consMercado_Click()

    '-----CONSTRUIR MERCADO-------------

    If 10 <= frmAlmacen.lblMadera And 10 <= frmAlmacen.lblBloque Then
        Creacion = "Mercado"
        Load Mercado(Mercado.Count)
        With Mercado(Mercado.Count - 1)
            .Move consMercado.Left + consMercado.Width, consMercado.Top + consMercado.Height, Mercado(0).Width, Mercado(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consMercado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Mercado: 10 bloques de piedra y 10 tablones de madera"
End Sub

Private Sub consMetalero_Click()

    '-----CONSTRUIR METALERO-------------

    If cantMetaleros < maxMetales Then
    If 3 <= frmAlmacen.lblMadera And 3 <= frmAlmacen.lblBloque Then
        Creacion = "casa de metalúrgico"
        Load Metalero(Metalero.Count)
        With Metalero(Metalero.Count - 1)
            .Move consMetalero.Left + consMetalero.Width, consMetalero.Top + consMetalero.Height, Metalero(0).Width, Metalero(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
    End If
End Sub

Private Sub consMetalero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cantMetaleros < maxMetales Then
        lblInfo = "Casa de Metalúrgico: 3 tablones de madera y 3 bloques de piedra"
    Else
        lblInfo = "Ya has ocupado todas las canteras de metal"
    End If
End Sub

Private Sub consPiedra_Click()
    
    '-----CONSTRUIR PICADOR-------------
    
    If cantPicadores < maxCanteras Then
    If 3 <= frmAlmacen.lblMadera And 2 <= frmAlmacen.lblBloque Then
        Creacion = "casa de Picador"
        Load PicaPiedra(PicaPiedra.Count)
        With PicaPiedra(PicaPiedra.Count - 1)
            .Move consPiedra.Left + consPiedra.Width, consPiedra.Top + consPiedra.Height, 2265, 2190
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
    End If
End Sub

Private Sub consPiedra_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cantPicadores < maxCanteras Then
        lblInfo = "Picador de Piedra: 3 tablones de madera y 2 bloques de piedra"
    Else
        lblInfo = "Ya has ocupado todas las canteras de piedra"
    End If
End Sub

Private Sub consRefineria_Click()

    '-----CONSTRUIR REFINERIA-------------

    If 3 <= frmAlmacen.lblMadera And 3 <= frmAlmacen.lblBloque Then
        Creacion = "refineria"
        Load Refineria(Refineria.Count)
        With Refineria(Refineria.Count - 1)
            .Move consRefineria.Left + consRefineria.Width, consRefineria.Top + consRefineria.Height, Refineria(0).Width, Refineria(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consRefineria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Refineria: 3 bloques de piedra y 3 tablones de madera"
End Sub

Private Sub consSerreria_Click()

    '-----CONSTRUIR SERRERIA-------------

    If 2 <= frmAlmacen.lblMadera And 3 <= frmAlmacen.lblBloque Then
        Creacion = "serreria"
        Load Serreria(Serreria.Count)
        With Serreria(Serreria.Count - 1)
            .Move consSerreria.Left + consSerreria.Width, consSerreria.Top + consSerreria.Height, Serreria(0).Width, Serreria(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .Drag
            .ZOrder
        End With
    End If
End Sub

Private Sub consSerreria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Serreria: 3 bloques de piedra y 2 tablones de madera"
End Sub

Private Sub Construir_Click()
    Obj = ""
    lblVidaFijo = ""
    lblVida = ""
    ima.Picture = Nothing
    Construir.Visible = False
    consFuente.Visible = True
    consCamino.Visible = True
    consCasa.Visible = True
    consGranja.Visible = True
    consPiedra.Visible = True
    consLeña.Visible = True
    cmdAtras.Visible = True
    consSerreria.Visible = True
    consMercado.Visible = True
    consRefineria.Visible = True
    consGranero.Visible = True
    consMetalero.Visible = True
    
    cmdMilitar.Visible = False
    consBarracas.Visible = False
    consArqueria.Visible = False
    consArmeria.Visible = False
    consCorazas.Visible = False
    GuardarTrigo.Visible = False
End Sub

Private Sub Construir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Construir"
End Sub

Private Sub CrearCuidadano_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Crear ciudadano"
End Sub

Private Sub Corazas_Click(Index As Integer)
    Obj = "Herreria: algunos soldados necesitan más que solo una espada, aqui puedes hacer escudos, cascos y armaduras"
    lblVidaFijo = "Vida: "
    lblVida = "300"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Corazas.GRAF")
End Sub

Private Sub Corazas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Corazas"
End Sub

Private Sub Flechas_Click(Index As Integer)

    If frmHistorial.Visible = True Then
        frmHistorial.ZOrder
    End If
    If frmPlantilla.Visible = True Then
       frmPlantilla.ZOrder
    End If
    If frmIndustria.Visible = True Then
       frmIndustria.ZOrder
    End If
    If frmEjercito.Visible = True Then
       frmEjercito.ZOrder
    End If
    If frmArmeria.Visible = True Then
       frmArmeria.ZOrder
    End If
    If frmMercado.Visible = True Then
       frmMercado.ZOrder
    End If
    If frmAlmacen.Visible = True Then
       frmAlmacen.ZOrder
    End If
    If frmObjetivos.Visible = True Then
       frmObjetivos.ZOrder
    End If
    
    
    If Horizontal <= 600 Then 'Controla que  no se vaya para el limite izquierdo
        Flechas(0).Enabled = False
    Else
        Flechas(0).Enabled = True
    End If
    
    If Horizontal > 29400 Then 'Controla que  no se vaya para el limite derecho
        Flechas(1).Enabled = False
    Else
        Flechas(1).Enabled = True
    End If
    
    If Vertical <= 600 Then
        Flechas(3).Enabled = False
    Else
        Flechas(3).Enabled = True
    End If
    
    If Vertical > 30000 Then
        Flechas(2).Enabled = False
    Else
        Flechas(2).Enabled = True
    End If
    
    Select Case Index
    Case 0
        Dim i#
        Horizontal = Horizontal - 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Left = Casa(i).Left + 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Left = Arbol(i).Left + 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Left = CanteraPiedra(i).Left + 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Left = Armeria(i).Left + 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Left = CanteraMetal(i).Left + 600
        Next
        CentroCiudad.Left = CentroCiudad.Left + 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Left = Barracas(i).Left + 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Left = Arqueria(i).Left + 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Left = Camino(i).Left + 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Left = Fuente(i).Left + 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Left = Mercado(i).Left + 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Left = Granja(i).Left + 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Left = Corazas(i).Left + 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Left = Granero(i).Left + 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Left = Leñador(i).Left + 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Left = Metalero(i).Left + 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Left = PicaPiedra(i).Left + 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Left = Serreria(i).Left + 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Left = Refineria(i).Left + 600
        Next
    Case 1
        Horizontal = Horizontal + 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Left = Casa(i).Left - 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Left = Arbol(i).Left - 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Left = CanteraPiedra(i).Left - 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Left = Armeria(i).Left - 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Left = CanteraMetal(i).Left - 600
        Next
        CentroCiudad.Left = CentroCiudad.Left - 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Left = Barracas(i).Left - 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Left = Arqueria(i).Left - 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Left = Camino(i).Left - 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Left = Fuente(i).Left - 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Left = Mercado(i).Left - 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Left = Granja(i).Left - 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Left = Corazas(i).Left - 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Left = Granero(i).Left - 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Left = Leñador(i).Left - 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Left = Metalero(i).Left - 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Left = PicaPiedra(i).Left - 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Left = Serreria(i).Left - 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Left = Refineria(i).Left - 600
        Next
    Case 2
        Vertical = Vertical + 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Top = Casa(i).Top - 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Top = Arbol(i).Top - 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Top = CanteraPiedra(i).Top - 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Top = Armeria(i).Top - 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Top = CanteraMetal(i).Top - 600
        Next
        CentroCiudad.Top = CentroCiudad.Top - 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Top = Barracas(i).Top - 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Top = Arqueria(i).Top - 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Top = Camino(i).Top - 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Top = Fuente(i).Top - 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Top = Mercado(i).Top - 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Top = Granja(i).Top - 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Top = Corazas(i).Top - 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Top = Granero(i).Top - 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Top = Leñador(i).Top - 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Top = Metalero(i).Top - 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Top = PicaPiedra(i).Top - 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Top = Serreria(i).Top - 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Top = Refineria(i).Top - 600
        Next
    Case 3
        Vertical = Vertical - 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Top = Casa(i).Top + 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Top = Arbol(i).Top + 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Top = CanteraPiedra(i).Top + 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Top = Armeria(i).Top + 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Top = CanteraMetal(i).Top + 600
        Next
        CentroCiudad.Top = CentroCiudad.Top + 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Top = Barracas(i).Top + 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Top = Arqueria(i).Top + 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Top = Camino(i).Top + 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Top = Fuente(i).Top + 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Top = Mercado(i).Top + 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Top = Granja(i).Top + 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Top = Corazas(i).Top + 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Top = Granero(i).Top + 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Top = Leñador(i).Top + 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Top = Metalero(i).Top + 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Top = PicaPiedra(i).Top + 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Top = Serreria(i).Top + 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Top = Refineria(i).Top + 600
        Next
   End Select
End Sub


Private Sub Flechas_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If frmHistorial.Visible = True Then
        frmHistorial.ZOrder
    End If
    If frmPlantilla.Visible = True Then
       frmPlantilla.ZOrder
    End If
    If frmIndustria.Visible = True Then
       frmIndustria.ZOrder
    End If
    If frmEjercito.Visible = True Then
       frmEjercito.ZOrder
    End If
    If frmArmeria.Visible = True Then
       frmArmeria.ZOrder
    End If
    If frmMercado.Visible = True Then
       frmMercado.ZOrder
    End If
    If frmAlmacen.Visible = True Then
       frmAlmacen.ZOrder
    End If
    If frmObjetivos.Visible = True Then
       frmObjetivos.ZOrder
    End If
    
    If Horizontal <= 600 Then 'Controla que  no se vaya para el limite izquierdo
        Flechas(0).Enabled = False
    Else
        Flechas(0).Enabled = True
    End If
    
    If Horizontal > 29400 Then 'Controla que  no se vaya para el limite derecho
        Flechas(1).Enabled = False
    Else
        Flechas(1).Enabled = True
    End If
    
    If Vertical <= 600 Then
        Flechas(3).Enabled = False
    Else
        Flechas(3).Enabled = True
    End If
    
    If Vertical > 30000 Then
        Flechas(2).Enabled = False
    Else
        Flechas(2).Enabled = True
    End If
    
    Select Case Index
    Case 0
        Dim i#
        Horizontal = Horizontal - 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Left = Casa(i).Left + 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Left = Arbol(i).Left + 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Left = CanteraPiedra(i).Left + 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Left = Armeria(i).Left + 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Left = CanteraMetal(i).Left + 600
        Next
        CentroCiudad.Left = CentroCiudad.Left + 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Left = Barracas(i).Left + 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Left = Arqueria(i).Left + 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Left = Camino(i).Left + 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Left = Fuente(i).Left + 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Left = Mercado(i).Left + 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Left = Granja(i).Left + 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Left = Corazas(i).Left + 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Left = Granero(i).Left + 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Left = Leñador(i).Left + 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Left = Metalero(i).Left + 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Left = PicaPiedra(i).Left + 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Left = Serreria(i).Left + 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Left = Refineria(i).Left + 600
        Next
    Case 1
        Horizontal = Horizontal + 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Left = Casa(i).Left - 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Left = Arbol(i).Left - 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Left = CanteraPiedra(i).Left - 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Left = Armeria(i).Left - 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Left = CanteraMetal(i).Left - 600
        Next
        CentroCiudad.Left = CentroCiudad.Left - 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Left = Barracas(i).Left - 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Left = Arqueria(i).Left - 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Left = Camino(i).Left - 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Left = Fuente(i).Left - 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Left = Mercado(i).Left - 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Left = Granja(i).Left - 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Left = Corazas(i).Left - 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Left = Granero(i).Left - 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Left = Leñador(i).Left - 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Left = Metalero(i).Left - 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Left = PicaPiedra(i).Left - 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Left = Serreria(i).Left - 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Left = Refineria(i).Left - 600
        Next
    Case 2
        Vertical = Vertical + 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Top = Casa(i).Top - 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Top = Arbol(i).Top - 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Top = CanteraPiedra(i).Top - 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Top = Armeria(i).Top - 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Top = CanteraMetal(i).Top - 600
        Next
        CentroCiudad.Top = CentroCiudad.Top - 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Top = Barracas(i).Top - 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Top = Arqueria(i).Top - 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Top = Camino(i).Top - 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Top = Fuente(i).Top - 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Top = Mercado(i).Top - 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Top = Granja(i).Top - 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Top = Corazas(i).Top - 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Top = Granero(i).Top - 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Top = Leñador(i).Top - 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Top = Metalero(i).Top - 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Top = PicaPiedra(i).Top - 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Top = Serreria(i).Top - 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Top = Refineria(i).Top - 600
        Next
    Case 3
        Vertical = Vertical - 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Top = Casa(i).Top + 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Top = Arbol(i).Top + 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Top = CanteraPiedra(i).Top + 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Top = Armeria(i).Top + 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Top = CanteraMetal(i).Top + 600
        Next
        CentroCiudad.Top = CentroCiudad.Top + 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Top = Barracas(i).Top + 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Top = Arqueria(i).Top + 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Top = Camino(i).Top + 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Top = Fuente(i).Top + 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Top = Mercado(i).Top + 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Top = Granja(i).Top + 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Top = Corazas(i).Top + 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Top = Granero(i).Top + 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Top = Leñador(i).Top + 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Top = Metalero(i).Top + 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Top = PicaPiedra(i).Top + 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Top = Serreria(i).Top + 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Top = Refineria(i).Top + 600
        Next
   End Select
End Sub

Private Sub Flechas_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If frmHistorial.Visible = True Then
        frmHistorial.ZOrder
    End If
    If frmPlantilla.Visible = True Then
       frmPlantilla.ZOrder
    End If
    If frmIndustria.Visible = True Then
       frmIndustria.ZOrder
    End If
    If frmEjercito.Visible = True Then
       frmEjercito.ZOrder
    End If
    If frmArmeria.Visible = True Then
       frmArmeria.ZOrder
    End If
    If frmMercado.Visible = True Then
       frmMercado.ZOrder
    End If
    If frmAlmacen.Visible = True Then
       frmAlmacen.ZOrder
    End If
    If frmObjetivos.Visible = True Then
       frmObjetivos.ZOrder
    End If
    
    If Horizontal <= 600 Then 'Controla que  no se vaya para el limite izquierdo
        Flechas(0).Enabled = False
    Else
        Flechas(0).Enabled = True
    End If
    
    If Horizontal > 29400 Then 'Controla que  no se vaya para el limite derecho
        Flechas(1).Enabled = False
    Else
        Flechas(1).Enabled = True
    End If
    
    If Vertical <= 600 Then
        Flechas(3).Enabled = False
    Else
        Flechas(3).Enabled = True
    End If
    
    If Vertical > 30000 Then
        Flechas(2).Enabled = False
    Else
        Flechas(2).Enabled = True
    End If
    
    Select Case Index
    Case 0
        Dim i#
        Horizontal = Horizontal - 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Left = Casa(i).Left + 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Left = Arbol(i).Left + 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Left = CanteraPiedra(i).Left + 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Left = Armeria(i).Left + 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Left = CanteraMetal(i).Left + 600
        Next
        CentroCiudad.Left = CentroCiudad.Left + 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Left = Barracas(i).Left + 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Left = Arqueria(i).Left + 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Left = Camino(i).Left + 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Left = Fuente(i).Left + 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Left = Mercado(i).Left + 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Left = Granja(i).Left + 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Left = Corazas(i).Left + 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Left = Granero(i).Left + 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Left = Leñador(i).Left + 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Left = Metalero(i).Left + 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Left = PicaPiedra(i).Left + 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Left = Serreria(i).Left + 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Left = Refineria(i).Left + 600
        Next
    Case 1
        Horizontal = Horizontal + 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Left = Casa(i).Left - 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Left = Arbol(i).Left - 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Left = CanteraPiedra(i).Left - 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Left = Armeria(i).Left - 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Left = CanteraMetal(i).Left - 600
        Next
        CentroCiudad.Left = CentroCiudad.Left - 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Left = Barracas(i).Left - 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Left = Arqueria(i).Left - 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Left = Camino(i).Left - 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Left = Fuente(i).Left - 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Left = Mercado(i).Left - 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Left = Granja(i).Left - 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Left = Corazas(i).Left - 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Left = Granero(i).Left - 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Left = Leñador(i).Left - 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Left = Metalero(i).Left - 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Left = PicaPiedra(i).Left - 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Left = Serreria(i).Left - 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Left = Refineria(i).Left - 600
        Next
    Case 2
        Vertical = Vertical + 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Top = Casa(i).Top - 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Top = Arbol(i).Top - 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Top = CanteraPiedra(i).Top - 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Top = Armeria(i).Top - 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Top = CanteraMetal(i).Top - 600
        Next
        CentroCiudad.Top = CentroCiudad.Top - 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Top = Barracas(i).Top - 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Top = Arqueria(i).Top - 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Top = Camino(i).Top - 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Top = Fuente(i).Top - 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Top = Mercado(i).Top - 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Top = Granja(i).Top - 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Top = Corazas(i).Top - 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Top = Granero(i).Top - 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Top = Leñador(i).Top - 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Top = Metalero(i).Top - 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Top = PicaPiedra(i).Top - 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Top = Serreria(i).Top - 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Top = Refineria(i).Top - 600
        Next
    Case 3
        Vertical = Vertical - 600
        For i = 0 To (Casa.Count - 1)
            Casa(i).Top = Casa(i).Top + 600
        Next
        For i = 0 To (Arbol.Count - 1)
            Arbol(i).Top = Arbol(i).Top + 600
        Next
        For i = 0 To (CanteraPiedra.Count - 1)
            CanteraPiedra(i).Top = CanteraPiedra(i).Top + 600
        Next
        For i = 0 To (Armeria.Count - 1)
            Armeria(i).Top = Armeria(i).Top + 600
        Next
        For i = 0 To (CanteraMetal.Count - 1)
            CanteraMetal(i).Top = CanteraMetal(i).Top + 600
        Next
        CentroCiudad.Top = CentroCiudad.Top + 600
        For i = 0 To (Barracas.Count - 1)
            Barracas(i).Top = Barracas(i).Top + 600
        Next
        For i = 0 To (Arqueria.Count - 1)
            Arqueria(i).Top = Arqueria(i).Top + 600
        Next
        For i = 0 To (Camino.Count - 1)
            Camino(i).Top = Camino(i).Top + 600
        Next
        For i = 0 To (Fuente.Count - 1)
            Fuente(i).Top = Fuente(i).Top + 600
        Next
        For i = 0 To (Mercado.Count - 1)
            Mercado(i).Top = Mercado(i).Top + 600
        Next
        For i = 0 To (Granja.Count - 1)
            Granja(i).Top = Granja(i).Top + 600
        Next
        For i = 0 To (Corazas.Count - 1)
            Corazas(i).Top = Corazas(i).Top + 600
        Next
        For i = 0 To (Granero.Count - 1)
            Granero(i).Top = Granero(i).Top + 600
        Next
        For i = 0 To (Leñador.Count - 1)
            Leñador(i).Top = Leñador(i).Top + 600
        Next
        For i = 0 To (Metalero.Count - 1)
            Metalero(i).Top = Metalero(i).Top + 600
        Next
        For i = 0 To (PicaPiedra.Count - 1)
            PicaPiedra(i).Top = PicaPiedra(i).Top + 600
        Next
        For i = 0 To (Serreria.Count - 1)
            Serreria(i).Top = Serreria(i).Top + 600
        Next
        For i = 0 To (Refineria.Count - 1)
            Refineria(i).Top = Refineria(i).Top + 600
        Next
   End Select
End Sub

Private Sub Form_Click()
    Obj = ""
    lblVidaFijo = ""
    lblVida = ""
    ima.Picture = Nothing
    cmdCrear.Visible = False
    Construir.Visible = True
    consFuente.Visible = False
    consCamino.Visible = False
    consCasa.Visible = False
    consLeña.Visible = False
    consPiedra.Visible = False
    consGranja.Visible = False
    cmdAtras.Visible = False
    consSerreria.Visible = False
    consMercado.Visible = False
    consRefineria.Visible = False
    consGranero.Visible = False
    consBarracas.Visible = False
    consArqueria.Visible = False
    consArmeria.Visible = False
    consCorazas.Visible = False
    If Objetivos2 = True Then
        cmdMilitar.Visible = True
    End If
    consMetalero.Visible = False
    GuardarTrigo.Visible = False
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
If Creacion <> "trigo" Then

    Select Case Creacion
    
    Case "casa de metalúrgico"
        cantMetaleros = cantMetaleros + 1
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 3
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 3
        
    Case "casa de leñador"
        cantLeñadores = cantLeñadores + 1
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 2
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 2
        
    Case "casa de Picador"
        cantPicadores = cantPicadores + 1
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 3
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 2
        
    Case "granja de trigo"
        cantGranjeros = cantGranjeros + 1
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 5
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 5
        
    Case "serreria"
        cantSerrerias = cantSerrerias + 1
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 2
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 3
        
    Case "refineria"
        cantRefinerias = cantRefinerias + 1
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 3
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 3
        
    Case "armeria"
        cantArmerias = cantArmerias + 1
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 3
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 3
        
    Case "corazas"
        cantCorazas = cantCorazas + 1
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 3
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 3
        
    Case "Mercado"
        mercadoBoolean = True
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 10
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 10
        frmObjetivos.Check2.Value = 1
        frmObjetivos.Check2.Enabled = False
        
    Case "barracas"
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 4
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 4
        cantBarracas = cantBarracas + 1
        maxSoldados = maxSoldados + 5
        
    Case "arqueria"
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 4
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 4
        cantArquerias = cantArquerias + 1
        maxArqueros = maxArqueros + 5
        
    Case "Granero"
        frmAlmacen.lblmaxTrigo = frmAlmacen.lblmaxTrigo + 500
        
    Case "fuente"
        lblPiedra = lblPiedra - 50
            
    Case "camino"
        lblPiedra = lblPiedra - 5
    
    Case "casa"
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 3
        frmAlmacen.lblBloque = frmAlmacen.lblBloque - 2
    
    End Select
    If booleanCasa = True Then 'Si lo que se contruye es una casa se aumenta el limite de poblacion en 5
        Dim TempInt#
        TempInt = LimiteCiudadanos.Caption
        TempInt = TempInt + 5
        limitePobl = limitePobl + 5
        LimiteCiudadanos.Caption = TempInt
    End If
    booleanCasa = False
    Source.Move X - (Source.Width / 2), Y - (Source.Height / 2)
    Source.Visible = True
    With lblInfo
        .Caption = ""
        .ZOrder
    End With
    frmHistorial.Historial.AddItem (Now & ":     Se creo un/a " & Creacion)
    End If
End Sub

Private Sub Form_Load()

    If Fondo = True Then
        frmJuego.Picture = Nothing
    End If
    
    If AudioMusica = False Then
        Load frmMusica
        frmMusica.Hide
    End If
    
    If AudioSonidos = False Then
        Load Sonidos
        Sonidos.Hide
    End If
        
    Load frmAlmacen
    Load frmMercado
    frmMercado.Hide
    TotMetal = maxMetales * 100
    TotPiedra = maxCanteras * 100
    TotMadera = cantArboles * 100
    
    frmJuego.MouseIcon = LoadPicture(App.Path & "\Punteros\frmJuego.ico")
    
    '--------------------------------------------------------------------------
    'TIRA AL BOLEO 50 ARBOLES, 6 CANTERAS DE PIEDRA Y 5 CANTERAS DE METAL------
    '--------------------------------------------------------------------------
    
    Dim Arboles#
    For Arboles = 0 To cantArboles
        Load Arbol(Arbol.Count)
        With Arbol(Arbol.Count - 1)
            Randomize
            .Move Rnd * 40000, Rnd * 40000, Arbol(0).Width, Arbol(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .ZOrder
            .Visible = True
        End With
    Next
    Dim cantRoca#
    For cantRoca = 0 To 5
        Load CanteraPiedra(CanteraPiedra.Count)
        With CanteraPiedra(CanteraPiedra.Count - 1)
            Randomize
            .Move Rnd * 40000, Rnd * 40000, CanteraPiedra(0).Width, CanteraPiedra(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .ZOrder
            .Visible = True
        End With
    Next
    Dim cantMetal#
    For cantMetal = 0 To 4
        Load CanteraMetal(CanteraMetal.Count)
        With CanteraMetal(CanteraMetal.Count - 1)
            Randomize
            .Move Rnd * 40000, Rnd * 40000, CanteraMetal(0).Width, CanteraMetal(0).Height
            .Appearance = 0
            .BorderStyle = 0
            .ZOrder
            .Visible = True
        End With
    Next
    lblInfo.ZOrder
    Load frmHistorial
    Load frmPlantilla
    frmHistorial.Hide
    frmPlantilla.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = ""
End Sub

Private Sub frmInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = ""
End Sub

Private Sub Fuente_Click(Index As Integer)
    Obj = "Fuente: tus ciudadanos necesitan beber agua potable."
    lblVidaFijo = "Vida: "
    lblVida = "50"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Fuente.GRAF")
End Sub

Private Sub Fuente_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Fuente"
End Sub

Private Sub Granero_Click(Index As Integer)
    Obj = "Granero: ademas de guardar el trigo, el granero lo convierte en pan para que tus ciudadanos coman"
    lblVidaFijo = "Vida: "
    lblVida = "500"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Granero.GRAF")
End Sub

Private Sub Granero_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Creacion = "trigo" Then
        frmAlmacen.lblTrigo = frmAlmacen.lblTrigo + 100
        lblComida = lblComida - 100
        If frmAlmacen.lblTrigo >= 1500 And Objetivos1 = True Then
            frmObjetivos.Check4.Value = 1
        End If
        frmHistorial.Historial.AddItem (Now & ": Guardaste 10 gavillas de trigo en un Granero")
    End If
End Sub

Private Sub Granero_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Granero"
End Sub

Private Sub Granja_Click(Index As Integer)
    Obj = "Granja: es el unico medio de extración de comida. Necesitas la comida para crear más ciudadanos y para alimentarlos"
    lblVidaFijo = "Vida: "
    lblVida = "200"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Granja.GRAF")
    GuardarTrigo.Visible = True
End Sub

Private Sub Granja_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Graja de trigo"
End Sub

Private Sub CentroCiudad_Click()
    Obj = "Centro de Ciudad"
    lblVidaFijo = "Vida "
    lblVida = 2000
    ima = LoadPicture(App.Path & "\Graficos\Chicos\CentroCiudad.GRAF")
    cmdCrear.Visible = True
    Construir.Visible = False
    consFuente.Visible = False
    consCamino.Visible = False
    consCasa.Visible = False
    consLeña.Visible = False
    consPiedra.Visible = False
    consGranja.Visible = False
    cmdAtras.Visible = False
    consSerreria.Visible = False
    consMercado.Visible = False
    consRefineria.Visible = False
    consGranero.Visible = False
    cmdMilitar.Visible = False
    consMetalero.Visible = False
    GuardarTrigo.Visible = False
    consArmeria.Visible = False
    consCorazas.Visible = False
    consBarracas.Visible = False
    consArqueria.Visible = False
End Sub

Private Sub CentroCiudad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Centro de Cuidad"
End Sub

Private Sub GuardarTrigo_Click()

    '-----DRAGEAR TRIGO-------------

    If lblComida >= 100 Then
        If frmAlmacen.lblTrigo + 100 < frmAlmacen.lblmaxTrigo Then
        Load Fuente(Fuente.Count)
        Creacion = "trigo"
        With Fuente(Fuente.Count - 1)
            .Move GuardarTrigo.Left + GuardarTrigo.Width, GuardarTrigo.Top + GuardarTrigo.Height, Fuente(0).Width, Fuente(0).Height
            .Drag
        End With
        End If
    End If
End Sub

Private Sub GuardarTrigo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Clickea y arrastra el trigo hacia un granero"
End Sub

Private Sub Leñador_Click(Index As Integer)
    Obj = "Casa de leñador: Aqui los leñadores cortan los pinos que luego se transforman en tablones de madera para construir"
    lblVidaFijo = "Vida: "
    lblVida = "150"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\CasaLeñador.GRAF")
    Sonidos.Sonidos.URL = App.Path & "\Sonidos\Madera.WAV"
End Sub

Private Sub Leñador_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Casa de Leñador"
End Sub

Private Sub Mercado_Click(Index As Integer)
    Obj = "Mercado: si tienes materia prima o productos de sobra los puedes vender. También puedes comprar"
    lblVidaFijo = "Vida "
    lblVida = "1000"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Mercado.GRAF")
End Sub

Private Sub Mercado_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Mercado"
End Sub

Private Sub Metalero_Click(Index As Integer)
    Obj = "Casa de Metalero: el metalero extrae el metal de las canteras y luego lo pone en calderos para fabricar armas o comerciar"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Metalero.GRAF")
    lblVida = "200"
    lblVidaFijo = "Vida: "
End Sub

Private Sub PicaPiedra_Click(Index As Integer)
    Obj = "Casa de Picador: Aqui los picadores cortan la piedra que luego se transforma el bloques manejables para construir"
    lblVidaFijo = "Vida: "
    lblVida = "150"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\CasaPiedra.GRAF")
End Sub

Private Sub PicaPiedra_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Casa de Picador de Piedra"
End Sub

Private Sub Refineria_Click(Index As Integer)
    Obj = "Refineria: aquí los pedazos de piedra obtenidos de las canteras se transforman en bloques para construir"
    lblVidaFijo = "Vida: "
    lblVida = "300"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Refineria.GRAF")
End Sub

Private Sub Refineria_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Refinería"
End Sub

Private Sub Serreria_Click(Index As Integer)
    Obj = "Serrería: aquí la madera extraída de los árboles se convierte en tablones de madera para construir"
    lblVidaFijo = "Vida: "
    lblVida = "300"
    ima = LoadPicture(App.Path & "\Graficos\Chicos\Serreria.GRAF")
    Sonidos.Sonidos.URL = App.Path & "\Sonidos\Leñador.WAV"
End Sub

Private Sub Serreria_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Serrería"
End Sub

Private Sub Son1_Click()
    Son1.Visible = False
    Son2.Visible = True
    Unload frmMusica
End Sub

Private Sub Son1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Desactivar Música"
End Sub

Private Sub Son2_Click()
    Son2.Visible = False
    Son1.Visible = True
    Load frmMusica
End Sub

Private Sub Son2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo = "Activar Música"
End Sub
