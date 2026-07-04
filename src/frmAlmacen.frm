VERSION 5.00
Begin VB.Form frmAlmacen 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacen"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "frmAlmacen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   3240
      X2              =   3240
      Y1              =   240
      Y2              =   3600
   End
   Begin VB.Label lblmaxTrigo 
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
      Left            =   6600
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
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
      Left            =   6360
      TabIndex        =   5
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label lblTrigo 
      BackStyle       =   0  'Transparent
      Caption         =   "500"
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
      Left            =   5280
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   1185
      Left            =   3600
      Picture         =   "frmAlmacen.frx":08CA
      Top             =   2280
      Width           =   1560
   End
   Begin VB.Label lblBloque 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
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
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblMadera 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "En almacen:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.Image Image3 
      Height          =   1155
      Left            =   3600
      Picture         =   "frmAlmacen.frx":6954
      Top             =   720
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   1050
      Left            =   360
      Picture         =   "frmAlmacen.frx":AA8E
      Top             =   960
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   480
      Picture         =   "frmAlmacen.frx":F300
      Top             =   2160
      Width           =   1290
   End
End
Attribute VB_Name = "frmAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Visible = False
End Sub
