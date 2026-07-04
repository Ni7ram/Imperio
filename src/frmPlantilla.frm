VERSION 5.00
Begin VB.Form frmPlantilla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plantilla"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   9960
      Top             =   2040
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SACAR"
      Height          =   375
      Left            =   7320
      TabIndex        =   32
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Parar producción"
      Height          =   375
      Left            =   5040
      TabIndex        =   31
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Parar producción"
      Height          =   375
      Left            =   5040
      TabIndex        =   30
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Parar producción"
      Height          =   375
      Left            =   5040
      TabIndex        =   29
      Top             =   600
      Width           =   1455
   End
   Begin VB.ComboBox sacPic 
      Height          =   315
      Left            =   7320
      TabIndex        =   28
      Text            =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox sacGra 
      Height          =   315
      Left            =   7320
      TabIndex        =   27
      Text            =   "0"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox ponLeñ 
      Height          =   315
      Left            =   9120
      TabIndex        =   26
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox ponPic 
      Height          =   315
      Left            =   9120
      TabIndex        =   25
      Text            =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox ponGra 
      Height          =   315
      Left            =   9120
      TabIndex        =   24
      Text            =   "0"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox sacLeñ 
      Height          =   315
      ItemData        =   "frmPlantilla.frx":0000
      Left            =   7320
      List            =   "frmPlantilla.frx":000D
      TabIndex        =   23
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   10335
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   10215
   End
   Begin VB.Line Line2 
      X1              =   2160
      X2              =   2160
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   6600
      X2              =   6600
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Label Label21 
      Caption         =   "Cambios:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   22
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label20 
      Caption         =   "Poner:"
      Height          =   255
      Left            =   8640
      TabIndex        =   21
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label19 
      Caption         =   "Sacar:"
      Height          =   255
      Left            =   6720
      TabIndex        =   20
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label18 
      Caption         =   "Poner:"
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Sacar:"
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "Poner:"
      Height          =   255
      Left            =   8640
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "Sacar:"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblTraDis 
      Caption         =   "15"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Trabajadores disponibles:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label graN 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label picN 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.Label leñN 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.Label graT 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label picT 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label leñT 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Oficio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Granjeros"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Picadores de Piedra"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Necesarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Trabajando"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Leñadores"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmPlantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim graRea As Boolean
    Dim picRea As Boolean
    Dim leñRea As Boolean

Private Sub cmdEjecutar_Click()
    Dim TempInt#
    TempInt = lblTraDis
    TempInt = TempInt - ponLeñ
    lblTraDis = TempInt
    TempInt = leñT
    TempInt = TempInt + ponLeñ
    leñT = TempInt
    
    TempInt = lblTraDis
    TempInt = TempInt - ponPic
    lblTraDis = TempInt
    TempInt = picT
    TempInt = TempInt + ponPic
    picT = TempInt
    
    TempInt = lblTraDis
    TempInt = TempInt - ponGra
    lblTraDis = TempInt
    TempInt = graT
    TempInt = TempInt + ponGra
    graT = TempInt
    
    ponLeñ = 0
    ponPic = 0
    ponGra = 0
End Sub

Private Sub Command1_Click()
    Dim TempInt#
    If leñRea = False Then
        leñRea = True
        Command1.Caption = "Reanudar"
        sacLeñ.Enabled = False
        ponLeñ.Enabled = False
        TempInt = lblTraDis
        TempInt = TempInt + leñT
        lblTraDis = TempInt
        leñN = 0
        leñT = 0
    Else
        leñRea = False
        Command1.Caption = "Parar producción"
        sacLeñ.Enabled = True
        ponLeñ.Enabled = True
        leñN = cantLeñadores * 5
    End If
End Sub

Private Sub Command2_Click()
    Dim TempInt#
    If picRea = False Then
        picRea = True
        Command2.Caption = "Reanudar"
        sacPic.Enabled = False
        ponPic.Enabled = False
        TempInt = lblTraDis
        TempInt = TempInt + picT
        lblTraDis = TempInt
        picT = 0
        picN = 0
    Else
        picRea = False
        Command2.Caption = "Parar producción"
        sacPic.Enabled = True
        ponPic.Enabled = True
        picN = cantPicadores * 5
    End If
End Sub

Private Sub Command3_Click()
    Dim TempInt#
    If graRea = False Then
        graRea = True
        Command3.Caption = "Reanudar"
        sacGra.Enabled = False
        ponGra.Enabled = False
        TempInt = lblTraDis
        TempInt = TempInt + picT
        lblTraDis = TempInt
        graT = 0
        graN = 0
    Else
        graRea = False
        Command3.Caption = "Parar producción"
        sacGra.Enabled = True
        ponGra.Enabled = True
        graN = cantGranjeros * 10
    End If
End Sub

Private Sub Form_Paint()
    leñN = (cantLeñadores * 5) - leñT
    picN = (cantPicadores * 5) - picT
    graN = (cantGranjeros * 10) - graT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Visible = False
End Sub

'*************************************************************
'-------------------------------------------------------------
'PONER - PONER - PONER - PONER - PONER - PONER - PONER - PONER
'-------------------------------------------------------------
'*************************************************************
Private Sub ponLeñ_DropDown()
    Dim i#
    ponLeñ.Clear
    For i = 0 To lblTraDis
        ponLeñ.AddItem (i)
    Next
End Sub

Private Sub ponPic_Click()
    Dim TempInt#
    If ponPic = "" Then
        ponPic = 0
    End If
    TempInt = ponPic
    If TempInt <= picN Then
        picN = picN - ponPic
        lblTraDis = lblTraDis - ponPic
        TempInt = picT
        TempInt = TempInt + ponPic
        picT = TempInt
        ponPic = "0"
    Else
        MsgBox "Sólo se requieren " + picN + " trabajadores más para las canteras de piedra"
        ponPic = "0"
        Exit Sub
    End If
End Sub

Private Sub ponpic_DropDown()
    Dim i#
    ponPic.Clear
    For i = 0 To lblTraDis
        ponPic.AddItem (i)
    Next
End Sub

Private Sub ponLeñ_Click()
    Dim TempInt#
    If ponLeñ = "" Then
        ponLeñ = 0
    End If
    TempInt = ponLeñ
    If TempInt <= leñN Then
        leñN = leñN - ponLeñ
        lblTraDis = lblTraDis - ponLeñ
        TempInt = leñT
        TempInt = TempInt + ponLeñ
        leñT = TempInt
        ponLeñ = "0"
    Else
        MsgBox "Sólo se requieren " + leñN + " trabajadores más para talar pinos"
        ponLeñ = "0"
        Exit Sub
    End If
End Sub

Private Sub pongra_DropDown()
    Dim i#
    ponGra.Clear
    For i = 0 To lblTraDis
        ponGra.AddItem (i)
    Next
End Sub

Private Sub ponGra_Click()
    Dim TempInt#
    If ponGra = "" Then
        ponGra = 0
    End If
    TempInt = ponGra
    If TempInt <= graN Then
        graN = graN - ponGra
        lblTraDis = lblTraDis - ponGra
        TempInt = graT
        TempInt = TempInt + ponGra
        graT = TempInt
        ponGra = "0"
    Else
        MsgBox "Sólo se requieren " + graN + " trabajadores más para las granjas"
        ponGra = "0"
        Exit Sub
    End If
End Sub
'*************************************************************
'-------------------------------------------------------------
'PONER - PONER - PONER - PONER - PONER - PONER - PONER - PONER
'-------------------------------------------------------------
'*************************************************************

'*************************************************************
'-------------------------------------------------------------
'SACAR - SACAR - SACAR - SACAR - SACAR - SACAR - SACAR - SACAR
'-------------------------------------------------------------
'*************************************************************
Private Sub Command4_Click()
    Dim TempInt#
    TempInt = lblTraDis
    TempInt = TempInt + sacLeñ
    leñT = leñT - sacLeñ
    TempInt = TempInt + sacPic
    picT = picT - sacPic
    TempInt = TempInt + sacGra
    graT = graT - sacGra
    lblTraDis = TempInt
    leñN = (cantLeñadores * 5) - leñT
    picN = (cantPicadores * 5) - picT
    graN = (cantGranjeros * 10) - graT
    sacGra = 0
    sacPic = 0
    sacLeñ = 0
End Sub

Private Sub sacgra_LostFocus()
    If sacGra = "" Then
        sacGra = 0
    End If
End Sub

Private Sub sacLeñ_DropDown()
    sacLeñ.Clear
    Dim i#
    For i = 0 To leñT
        sacLeñ.AddItem (i)
    Next
End Sub

Private Sub sacLeñ_LostFocus()
    If sacLeñ = "" Then
        sacLeñ = 0
    End If
End Sub

Private Sub sacPic_DropDown()
    sacPic.Clear
    Dim i#
    For i = 0 To picT
        sacPic.AddItem (i)
    Next
End Sub

Private Sub sacGra_DropDown()
    sacGra.Clear
    Dim i#
    For i = 0 To graT
        sacGra.AddItem (i)
    Next
End Sub

Private Sub sacPic_LostFocus()
    If sacPic = "" Then
        sacPic = 0
    End If
End Sub
'*************************************************************
'-------------------------------------------------------------
'SACAR - SACAR - SACAR - SACAR - SACAR - SACAR - SACAR - SACAR
'-------------------------------------------------------------
'*************************************************************


Private Sub Timer1_Timer()

    '-------------------BENEFICIOS------------------------------------
    

    '-------------------Para Granjas
    
    If cantGranjeros <> 0 Then
        frmJuego.lblComida = frmJuego.lblComida + (graT \ 5) 'Para que se vayan agotando las granjas
        Dim i#
        For i = 1 To (frmJuego.Granja.Count - 1)
            frmJuego.Granja(i).Tag = frmJuego.Granja(i).Tag - (graT \ 5)
        Next
    End If
    
    '-------------------Para Leñadores
    

    frmJuego.lblMadera = frmJuego.lblMadera + (leñT \ 5)

    
    '-------------------Para Picadores

    frmJuego.lblPiedra = frmJuego.lblPiedra + (picT \ 5)

End Sub
