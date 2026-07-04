VERSION 5.00
Begin VB.Form frmHistorial 
   BackColor       =   &H80000012&
   Caption         =   "Historial"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   Icon            =   "frmHistorial.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar Historial"
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox Historial 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2205
      ItemData        =   "frmHistorial.frx":08CA
      Left            =   120
      List            =   "frmHistorial.frx":08D1
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Historial.Clear
End Sub

Private Sub Form_Resize()
    Historial.Left = 60
    Historial.Top = 60
    If ScaleWidth > Command1.Width + 50 Then
        Command1.Left = ScaleWidth - Command1.Width - 120
        If Command1.Left > 140 Then
            Historial.Width = Command1.Left - 120
        End If
    End If
    If ScaleHeight > 120 Then
        Historial.Height = ScaleHeight - 120
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    frmHistorial.Visible = False
End Sub
