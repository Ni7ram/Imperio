VERSION 5.00
Begin VB.Form frmEjercito 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejercito"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10890
   Icon            =   "frmEjercito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   6180
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   9840
      Top             =   3240
   End
   Begin VB.ComboBox ponGra 
      Height          =   315
      Left            =   9000
      TabIndex        =   24
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000008&
      Height          =   1815
      Left            =   4680
      Picture         =   "frmEjercito.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      Picture         =   "frmEjercito.frx":A4B8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000008&
      Height          =   1815
      Left            =   4680
      Picture         =   "frmEjercito.frx":136A2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      Picture         =   "frmEjercito.frx":1DCD4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar Reclutas"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9000
      TabIndex        =   25
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image7 
      Height          =   840
      Left            =   9120
      Picture         =   "frmEjercito.frx":2805A
      Top             =   4920
      Width           =   825
   End
   Begin VB.Label cantEspadas 
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
      Left            =   10200
      TabIndex        =   23
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Guerreros 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6480
      TabIndex        =   22
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Lanzas 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6480
      TabIndex        =   21
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Arqueros 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Espadas 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   19
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "GUERRERO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lba 
      BackStyle       =   0  'Transparent
      Caption         =   "SOLDADO CON LANZA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   6480
      TabIndex        =   17
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "SOLDADO CON ESPADA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2040
      TabIndex        =   16
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "ARQUERO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label cantReclutas 
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
      Left            =   1080
      TabIndex        =   14
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   360
      Picture         =   "frmEjercito.frx":2A55C
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 1 recluta, 1 espada, 1 escudo y 1 armadura."
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 1 recluta y 1 lanza."
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 1 recluta y 1 espada."
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 1 recluta y 1 arco."
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label cantLanzas 
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
      Left            =   8400
      TabIndex        =   9
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label cantArmaduras 
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
      Left            =   7080
      TabIndex        =   8
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label cantEscudos 
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
      Left            =   5520
      TabIndex        =   7
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label cantCascos 
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
      Left            =   3840
      TabIndex        =   6
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label cantArcos 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   750
      Left            =   3120
      Picture         =   "frmEjercito.frx":2C75E
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image Image6 
      Height          =   960
      Left            =   1800
      Picture         =   "frmEjercito.frx":2E488
      Top             =   4920
      Width           =   510
   End
   Begin VB.Image Image4 
      Height          =   1020
      Left            =   4560
      Picture         =   "frmEjercito.frx":2FECA
      Top             =   4920
      Width           =   1020
   End
   Begin VB.Image Image3 
      Height          =   1095
      Left            =   7800
      Picture         =   "frmEjercito.frx":3353C
      Top             =   4800
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   6240
      Picture         =   "frmEjercito.frx":359FE
      Top             =   5040
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Almacén:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   10560
      Y1              =   4320
      Y2              =   4320
   End
End
Attribute VB_Name = "frmEjercito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If cantArcos > 0 And cantReclutas > 0 And cantArqueros + 1 <= maxArqueros Then
        Arqueros = Arqueros + 1
        cantArcos = cantArcos - 1
        cantReclutas = cantReclutas - 1
        frmHistorial.Historial.AddItem (Now & ": Creaste un arquero")
        cantSoldados = cantSoldados + 1
        cantArqueros = cantArqueros + 1
    End If
End Sub

Private Sub Command5_Click()
    If cantEspadas > 0 And cantReclutas > 0 And cantArmaduras > 0 And cantCascos > 0 And cantSoldados + 1 <= maxSoldados Then
        Guerreros = Guerreros + 1
        cantEspadas = cantEspadas - 1
        cantArmaduras = cantArmaduras - 1
        cantCascos = cantCascos - 1
        cantReclutas = cantReclutas - 1
        frmHistorial.Historial.AddItem (Now & ": Creaste un Guerrero")
        cantSoldados = cantSoldados + 1
    End If
End Sub

Private Sub Command6_Click()
    If cantEspadas > 0 And cantReclutas > 0 And cantSoldados + 1 <= maxSoldados Then
        Espadas = Espadas + 1
        cantEspadas = cantEspadas - 1
        cantReclutas = cantReclutas - 1
        frmHistorial.Historial.AddItem (Now & ": Creaste un soldado con espada")
        cantSoldados = cantSoldados + 1
   End If
End Sub

Private Sub Command7_Click()
    If cantLanzas > 0 And cantReclutas > 0 And cantSoldados + 1 <= maxSoldados Then
        Lanzas = Lanzas + 1
        cantLanzas = cantLanzas - 1
        cantReclutas = cantReclutas - 1
        frmHistorial.Historial.AddItem (Now & ": Creaste un soldado con lanza")
        cantSoldados = cantSoldados + 1
    End If
End Sub

Private Sub Form_Load()
    Me.MouseIcon = LoadPicture(App.Path & "\Punteros\frmEjercito.ico")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    frmEjercito.Hide
End Sub

Private Sub ponGra_Click()
    Dim TempInt#
    TempInt = cantReclutas
    TempInt = TempInt + ponGra
    cantReclutas = TempInt
    frmPlantilla.lblTraDis = frmPlantilla.lblTraDis - ponGra
    ponGra = "0"
End Sub

Private Sub ponGra_DropDown()
    Dim i#
    ponGra.Clear
    For i = 0 To frmPlantilla.lblTraDis
        ponGra.AddItem (i)
    Next
End Sub

Private Sub Timer1_Timer()
    If Arqueros >= 3 And Espadas >= 3 And Lanzas >= 3 And Guerreros >= 3 Then
        frmObjetivos.Check4.Value = 1
        frmObjetivos.Check4.Enabled = False
    End If
End Sub
