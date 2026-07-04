VERSION 5.00
Begin VB.Form frmObjetivos 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Objetivos"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   Icon            =   "frmObjetivos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check4 
      BackColor       =   &H80000012&
      Caption         =   "Tener almacenados 1500 panes o mas"
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
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000012&
      Caption         =   "Tener en el tesoro 500 monedas de oro"
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
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   120
      Top             =   2880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listo"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000012&
      Caption         =   "Construir un mercado"
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
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000012&
      Caption         =   "Tener 50 ciudadanos"
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
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmObjetivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Check1.Value = 0
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Check1.Value = 0
End Sub

Private Sub Check2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Check2.Value = 0
End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Check2.Value = 0
End Sub

Private Sub Check3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Check3.Value = 0
End Sub

Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Check3.Value = 0
End Sub

Private Sub Check4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Check4.Value = 0
End Sub

Private Sub Check4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Check4.Value = 0
End Sub

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Hide
End Sub

Private Sub Timer1_Timer()
    If Check1.Value = 1 And Check2.Value = 1 And Objetivos1 = False Then
        MsgBox "Cumpliste tus objetivos y se te han asignado otros", vbInformation, "Objetivos Cumplidos"
        frmHistorial.Historial.AddItem ("***OBJETIVOS CUMPLIDOS***")
        Check1.Visible = False
        Check2.Visible = False
        Check3.Visible = True
        Check4.Visible = True
        Objetivos1 = True
    End If
    If Check1.Value = 3 And Check2.Value = 4 And Objetivos2 = False Then
        MsgBox "Cumpliste tus objetivos y se te han asignado otros", vbInformation, "Objetivos Cumplidos"
        frmHistorial.Historial.AddItem ("***OBJETIVOS CUMPLIDOS***")
        Check3.Visible = False
        Check4.Visible = False
        frmJuego.Command1.Visible = True
        frmJuego.Command3.Visible = True
        frmMercado.Command1.Enabled = True
        frmMercado.Command2.Enabled = True
        frmMercado.Command3.Enabled = True
        frmMercado.Command4.Enabled = True
        frmMercado.Command5.Enabled = True
        frmMercado.Command6.Enabled = True
        frmMercado.Command7.Enabled = True
        frmMercado.Command8.Enabled = True
        frmMercado.Command9.Enabled = True
        frmMercado.Command10.Enabled = True
        frmMercado.Command11.Enabled = True
        frmMercado.Command12.Enabled = True
        Objetivos2 = True
    End If
End Sub
