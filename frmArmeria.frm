VERSION 5.00
Begin VB.Form frmArmeria 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Armeria"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmArmeria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmArmeria.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   7290
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Comenzar producción"
      Height          =   615
      Left            =   7560
      TabIndex        =   31
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6960
      Top             =   6480
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000007&
      Height          =   2055
      Left            =   240
      Picture         =   "frmArmeria.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000007&
      Height          =   1935
      Left            =   4800
      Picture         =   "frmArmeria.frx":A092
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000007&
      Height          =   1935
      Left            =   4800
      Picture         =   "frmArmeria.frx":112E4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000007&
      Height          =   1815
      Left            =   240
      Picture         =   "frmArmeria.frx":1EA4E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000007&
      Height          =   1815
      Left            =   4800
      Picture         =   "frmArmeria.frx":26144
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Height          =   1815
      Left            =   240
      Picture         =   "frmArmeria.frx":2C786
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblMadera 
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
      Height          =   255
      Left            =   6000
      TabIndex        =   30
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label lblMetal 
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
      Height          =   255
      Left            =   4440
      TabIndex        =   29
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label lblOro 
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
      Height          =   255
      Left            =   2760
      TabIndex        =   28
      Top             =   6600
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   525
      Left            =   5280
      Picture         =   "frmArmeria.frx":3662C
      Top             =   6480
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   3840
      Picture         =   "frmArmeria.frx":3787A
      Top             =   6480
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   2280
      Picture         =   "frmArmeria.frx":38930
      Top             =   6480
      Width           =   555
   End
   Begin VB.Label LanzasCola 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label EspadasCola 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label ArcosCola 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label EscudosCola 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   24
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label CascosCola 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label ArmadurasCola 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblEspadas 
      BackColor       =   &H00FF0000&
      Height          =   135
      Left            =   2160
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblArcos 
      BackColor       =   &H00FF0000&
      Height          =   135
      Left            =   6720
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblEscudos 
      BackColor       =   &H00FF0000&
      Height          =   135
      Left            =   6720
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblCascos 
      BackColor       =   &H00FF0000&
      Height          =   135
      Left            =   6720
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblLanzas 
      BackColor       =   &H00FF0000&
      Height          =   135
      Left            =   2160
      TabIndex        =   17
      Top             =   2760
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblArmaduras 
      BackColor       =   &H00FF0000&
      Height          =   135
      Left            =   2160
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   9000
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Armerias:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Corazas:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label11 
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
      Left            =   1320
      TabIndex        =   13
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label12 
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
      Left            =   1320
      TabIndex        =   12
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 5 calderos de metal y 20 de oro"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 5 calderos de metal y 20 de oro"
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   6720
      TabIndex        =   10
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 10 calderos de metal y 50 de oro"
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 5 calderos de metal y 3  tablones"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 1 caldero de metal y 1 tablón de madera."
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   2160
      TabIndex        =   7
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Costo: 3 tablones de madera."
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   6720
      TabIndex        =   6
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmArmeria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If lblMetal > 9 And lblOro > 49 And cantArmerias > 0 Then
        ArmadurasCola = ArmadurasCola + 1
        lblOro = lblOro - 50
        lblMetal = lblMetal - 10
        ArmadurasCola.Visible = True
    End If
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And ArmadurasCola > 0 Then
        ArmadurasCola = ArmadurasCola - 1
        If ArmadurasCola = 0 Then
            ArmadurasCola.Visible = False
        End If
        lblOro = lblOro + 50
        lblMetal = lblMetal + 10
        With lblArmaduras
            .Visible = False
            .Width = 1500
        End With
    End If
End Sub

Private Sub Command2_Click()
    If lblMadera > 2 And cantArmerias > 0 Then
        ArcosCola = ArcosCola + 1
        frmAlmacen.lblMadera = frmAlmacen.lblMadera - 3
        ArcosCola.Visible = True
    End If
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And ArcosCola > 0 Then
        ArcosCola = ArcosCola - 1
        lblMadera = lblMadera + 3
        If ArcosCola = 0 Then
            ArcosCola.Visible = False
        End If
        With lblArcos
            .Visible = False
            .Width = 1500
        End With
    End If
End Sub

Private Sub Command3_Click()
    If lblMetal > 4 And lblMadera > 2 And cantArmerias > 0 Then
        lblMetal = lblMetal - 5
        lblMadera = lblMadera - 3
        EspadasCola.Visible = True
        EspadasCola = EspadasCola + 1
    End If
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And EspadasCola > 0 Then
        EspadasCola = EspadasCola - 1
        If EspadasCola = 0 Then
            EspadasCola.Visible = False
        End If
        lblMetal = lblMetal + 5
        lblMadera = lblMadera + 3
        With lblEspadas
            .Visible = False
            .Width = 1500
        End With
    End If
End Sub

Private Sub Command4_Click()
    If lblOro > 19 And lblMetal > 4 And cantArmerias > 0 Then
        lblOro = lblOro - 20
        lblMetal = lblMetal - 5
        EscudosCola = EscudosCola + 1
        EscudosCola.Visible = True
    End If
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And EscudosCola > 0 Then
        EscudosCola = EscudosCola - 1
        If EscudosCola = 0 Then
            EscudosCola.Visible = False
        End If
        lblOro = lblOro + 20
        lblMetal = lblMetal + 5
        With lblEscudos
            .Visible = False
            .Width = 1500
        End With
    End If
End Sub

Private Sub Command5_Click()
    If lblOro > 19 And lblMetal > 4 And cantArmerias > 0 Then
        lblOro = lblOro - 20
        lblMetal = lblMetal - 5
        CascosCola = CascosCola + 1
        CascosCola.Visible = True
    End If
End Sub

Private Sub Command5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And CascosCola > 0 Then
        CascosCola = CascosCola - 1
        If CascosCola = 0 Then
            CascosCola.Visible = False
        End If
        lblOro = lblOro + 20
        lblMetal = lblMetal + 5
        With lblCascos
            .Visible = False
            .Width = 1500
        End With
    End If
End Sub

Private Sub Command6_Click()
    If lblMetal > 0 And lblMadera > 0 And cantArmerias > 0 Then
        LanzasCola = LanzasCola + 1
        lblMetal = lblMetal - 1
        LanzasCola.Visible = True
        lblMadera = lblMadera - 1
    End If
End Sub

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And LanzasCola > 0 Then
        LanzasCola = LanzasCola - 1
        If LanzasCola = 0 Then
            LanzasCola.Visible = False
        End If
        lblMetal = lblMetal + 1
        lblMadera = lblMadera + 1
        With lblLanzas
            .Visible = False
            .Width = 1500
        End With
    End If
End Sub

Private Sub Command7_Click()
    If Command7.Caption = "Comenzar producción" Then
        Timer1.Enabled = True
        Command7.Caption = "Detener producción"
    Else
        Command7.Caption = "Comenzar producción"
        Timer1.Enabled = False
    End If
End Sub

Private Sub Form_GotFocus()
    lblOro = frmJuego.lblOro
    lblMetal = frmAlmacen.lblMetal
    lblMadera = frmAlmacen.lblMadera
End Sub

Private Sub Form_Load()
    Me.MouseIcon = LoadPicture(App.Path & "\Punteros\frmEjercito.ico")
End Sub

Private Sub Form_Paint()
    lblOro = frmJuego.lblOro
    lblMetal = frmAlmacen.lblMetal
    lblMadera = frmAlmacen.lblMadera
End Sub

Private Sub Form_Resize()
    lblOro = frmJuego.lblOro
    lblMetal = frmAlmacen.lblMetal
    lblMadera = frmAlmacen.lblMadera
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    frmArmeria.Hide
End Sub

Private Sub Timer1_Timer()
    If ArmadurasCola > 0 Then
        lblArmaduras.Visible = True
        If lblArmaduras.Width > 30 Then
            lblArmaduras.Width = lblArmaduras.Width - 10
        Else
            With lblArmaduras
                .Visible = False
                .Width = 1500
            End With
            With ArmadurasCola
                .Caption = .Caption - 1
                If .Caption = 0 Then
                    .Visible = False
                End If
            End With
            frmEjercito.cantArmaduras = frmEjercito.cantArmaduras + 1
        End If
    End If
    
    If LanzasCola > 0 Then
        lblLanzas.Visible = True
        If lblLanzas.Width > 30 Then
            lblLanzas.Width = lblLanzas.Width - 10
        Else
            With lblLanzas
                .Visible = False
                .Width = 1500
            End With
            With LanzasCola
                .Caption = .Caption - 1
                If .Caption = 0 Then
                    .Visible = False
                End If
            End With
            frmEjercito.cantLanzas = frmEjercito.cantLanzas + 1
        End If
    End If
    
    If EspadasCola > 0 Then
        lblEspadas.Visible = True
        If lblEspadas.Width > 30 Then
            lblEspadas.Width = lblEspadas.Width - 10
        Else
            With lblEspadas
                .Visible = False
                .Width = 1500
            End With
            With EspadasCola
                .Caption = .Caption - 1
                If .Caption = 0 Then
                    .Visible = False
                End If
            End With
            frmEjercito.cantEspadas = frmEjercito.cantEspadas + 1
        End If
    End If
    
    If ArcosCola > 0 Then
        lblArcos.Visible = True
        If lblArcos.Width > 30 Then
            lblArcos.Width = lblArcos.Width - 10
        Else
            With lblArcos
                .Visible = False
                .Width = 1500
            End With
            With ArcosCola
                .Caption = .Caption - 1
                If .Caption = 0 Then
                    .Visible = False
                End If
            End With
            frmEjercito.cantArcos = frmEjercito.cantArcos + 1
        End If
    End If
    
    If EscudosCola > 0 Then
        lblEscudos.Visible = True
        If lblEscudos.Width > 30 Then
            lblEscudos.Width = lblEscudos.Width - 10
        Else
            With lblEscudos
                .Visible = False
                .Width = 1500
            End With
            With EscudosCola
                .Caption = .Caption - 1
                If .Caption = 0 Then
                    .Visible = False
                End If
            End With
            frmEjercito.cantEscudos = frmEjercito.cantEscudos + 1
        End If
    End If
        
    If CascosCola > 0 Then
        lblCascos.Visible = True
        If lblCascos.Width > 30 Then
            lblCascos.Width = lblCascos.Width - 10
        Else
            With lblCascos
                .Visible = False
                .Width = 1500
            End With
            With CascosCola
                .Caption = .Caption - 1
                If .Caption = 0 Then
                    .Visible = False
                End If
            End With
            frmEjercito.cantCascos = frmEjercito.cantCascos + 1
        End If
    End If
    If ArmadurasCola = 0 And LanzasCola = 0 And EspadasCola = 0 And CascosCola = 0 And EscudosCola = 0 And ArcosCola = 0 Then
        Timer1.Enabled = False
        Command7.Caption = "Comenzar producción"
    End If
End Sub
