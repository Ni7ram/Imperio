VERSION 5.00
Begin VB.Form frmIndustria 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Industria"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmIndustria.frx":0000
   LinkTopic       =   "Industria"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Comenzar producción"
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2160
      Top             =   3840
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000007&
      Height          =   1215
      Left            =   480
      Picture         =   "frmIndustria.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Height          =   1215
      Left            =   480
      Picture         =   "frmIndustria.frx":513C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblRefinerias 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblSerrerias 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblTablones 
      BackColor       =   &H00FF0000&
      Height          =   135
      Left            =   1800
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblBloques 
      BackColor       =   &H00FF0000&
      Height          =   135
      Left            =   1800
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Refinerias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Serrerias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   4440
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   3480
      X2              =   3480
      Y1              =   240
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   4440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label TablonesCola 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label BloquesCola 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "En cola:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 100 madera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 100 piedra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenar hacer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmIndustria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If frmJuego.lblPiedra > 99 And lblRefinerias > 0 And frmPlantilla.refT > 0 Then
        frmJuego.lblPiedra = frmJuego.lblPiedra - 100
        BloquesCola = BloquesCola + 1
    End If
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And BloquesCola > 0 Then
        BloquesCola = BloquesCola - 1
        frmJuego.lblPiedra = frmJuego.lblPiedra + 100
        With lblBloques
            .Visible = False
            .Width = 1500
        End With
    End If
End Sub

Private Sub Command2_Click()
    If Command2.Caption = "Comenzar producción" Then
        Timer1.Enabled = True
        Command2.Caption = "Detener producción"
    Else
        Command2.Caption = "Comenzar producción"
        Timer1.Enabled = False
    End If
End Sub

Private Sub Command3_Click()
    If frmJuego.lblMadera > 99 And lblSerrerias > 0 And frmPlantilla.carT > 0 Then
        frmJuego.lblMadera = frmJuego.lblMadera - 100
        TablonesCola = TablonesCola + 1
    End If
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And TablonesCola > 0 Then
        TablonesCola = TablonesCola - 1
        frmJuego.lblMadera = frmJuego.lblMadera + 100
        With lblTablones
            .Visible = False
            .Width = 1500
        End With
    End If
End Sub

Private Sub Form_Paint()
    lblSerrerias = cantSerrerias
    lblRefinerias = cantRefinerias
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Visible = False
End Sub

Private Sub Timer1_Timer()
    If BloquesCola > 0 Then
        lblBloques.Visible = True
        If lblBloques.Width > 30 Then
            lblBloques.Width = lblBloques.Width - (frmPlantilla.refT * 3)
        Else
            With lblBloques
                .Visible = False
                .Width = 1500
            End With
            BloquesCola = BloquesCola - 1
            frmAlmacen.lblBloque = frmAlmacen.lblBloque + 1
        End If
    End If

    If TablonesCola > 0 Then
        lblTablones.Visible = True
        If lblTablones.Width > 30 Then
            lblTablones.Width = lblTablones.Width - (frmPlantilla.carT * 3)
        Else
            With lblTablones
                .Visible = False
                .Width = 1500
            End With
            TablonesCola = TablonesCola - 1
            frmAlmacen.lblMadera = frmAlmacen.lblMadera + 1
        End If
    End If
    If TablonesCola = 0 And BloquesCola = 0 Then
        Timer1.Enabled = False
        Command2.Caption = "Comenzar producción"
    End If
End Sub
