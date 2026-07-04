VERSION 5.00
Begin VB.Form frmMercado 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mercado"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   Icon            =   "frmMercado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   9630
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6480
      Top             =   8400
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Comprar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   73
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Vender"
      Height          =   375
      Left            =   5400
      TabIndex        =   31
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Comprar"
      Height          =   375
      Left            =   5400
      TabIndex        =   30
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Vender"
      Height          =   375
      Left            =   5400
      TabIndex        =   29
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Comprar"
      Height          =   375
      Left            =   5400
      TabIndex        =   28
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Comprar"
      Height          =   375
      Left            =   5400
      TabIndex        =   27
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Vender"
      Height          =   375
      Left            =   5400
      TabIndex        =   26
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Vender"
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Comprar"
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Comprar"
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Vender"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Vender"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   8760
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Comprar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Vender"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Comprar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Vender"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Comprar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Vender"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Comprar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Vender"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Vender"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comprar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Line Line34 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   3720
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line Line33 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   240
      Y1              =   9360
      Y2              =   8160
   End
   Begin VB.Line Line32 
      BorderColor     =   &H80000009&
      X1              =   3720
      X2              =   3720
      Y1              =   9360
      Y2              =   8160
   End
   Begin VB.Label Label57 
      BackStyle       =   0  'Transparent
      Caption         =   "Escudo"
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
      Left            =   360
      TabIndex        =   78
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label Label56 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   77
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Label55 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   76
      Top             =   8760
      Width           =   495
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Caption         =   "500"
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
      Left            =   3120
      TabIndex        =   75
      Top             =   8760
      Width           =   495
   End
   Begin VB.Label Label53 
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
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
      Left            =   3120
      TabIndex        =   74
      Top             =   8280
      Width           =   495
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5160
      TabIndex        =   72
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Image Image12 
      Height          =   1170
      Left            =   4080
      Picture         =   "frmMercado.frx":08CA
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   7320
      TabIndex        =   71
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   7320
      TabIndex        =   70
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   7320
      TabIndex        =   69
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   7320
      TabIndex        =   68
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   7320
      TabIndex        =   67
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
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
      Left            =   7320
      TabIndex        =   66
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   7320
      TabIndex        =   65
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
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
      Left            =   7320
      TabIndex        =   64
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "500"
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
      Left            =   7320
      TabIndex        =   63
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   7320
      TabIndex        =   62
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
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
      Left            =   3120
      TabIndex        =   61
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "500"
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
      Left            =   3120
      TabIndex        =   60
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
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
      Left            =   3120
      TabIndex        =   59
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "500"
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
      Left            =   3120
      TabIndex        =   58
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "600"
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
      Left            =   3120
      TabIndex        =   57
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "500"
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
      Left            =   3120
      TabIndex        =   56
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "400"
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
      Left            =   3120
      TabIndex        =   55
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "300"
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
      Left            =   3120
      TabIndex        =   54
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "500"
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
      Left            =   3120
      TabIndex        =   53
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "400"
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
      Left            =   3120
      TabIndex        =   52
      Top             =   1440
      Width           =   495
   End
   Begin VB.Line Line28 
      BorderColor     =   &H80000009&
      X1              =   3960
      X2              =   7800
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line25 
      BorderColor     =   &H80000009&
      X1              =   3960
      X2              =   7800
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line22 
      BorderColor     =   &H80000009&
      X1              =   3960
      X2              =   7800
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line19 
      BorderColor     =   &H80000009&
      X1              =   3960
      X2              =   7800
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   6840
      TabIndex        =   51
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   6840
      TabIndex        =   50
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   6840
      TabIndex        =   49
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   6840
      TabIndex        =   48
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   6840
      TabIndex        =   47
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   6840
      TabIndex        =   46
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   6840
      TabIndex        =   45
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   6840
      TabIndex        =   44
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   6840
      TabIndex        =   43
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   6840
      TabIndex        =   42
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   41
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   40
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   39
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   38
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   37
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   36
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   35
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   34
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   33
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Oro:"
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
      Left            =   2640
      TabIndex        =   32
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image11 
      Height          =   855
      Left            =   4080
      Picture         =   "frmMercado.frx":4C14
      Top             =   5400
      Width           =   1170
   End
   Begin VB.Image Image9 
      Height          =   1050
      Left            =   4080
      Picture         =   "frmMercado.frx":80E2
      Top             =   3960
      Width           =   1305
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tablones de Madera"
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
      Left            =   4080
      TabIndex        =   25
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Madera (100)"
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
      Left            =   4080
      TabIndex        =   24
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Piedra (100)"
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
      Left            =   4080
      TabIndex        =   23
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bloque de Piedra"
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
      Left            =   4200
      TabIndex        =   22
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Caldero de Metal"
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
      Left            =   4200
      TabIndex        =   21
      Top             =   720
      Width           =   2295
   End
   Begin VB.Image Image8 
      Height          =   1035
      Left            =   4080
      Picture         =   "frmMercado.frx":C954
      Top             =   2400
      Width           =   1290
   End
   Begin VB.Line Line31 
      BorderColor     =   &H80000009&
      X1              =   3960
      X2              =   7800
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line30 
      BorderColor     =   &H80000009&
      X1              =   3960
      X2              =   3960
      Y1              =   7920
      Y2              =   6720
   End
   Begin VB.Line Line29 
      BorderColor     =   &H80000009&
      X1              =   7800
      X2              =   7800
      Y1              =   7920
      Y2              =   6720
   End
   Begin VB.Line Line27 
      BorderColor     =   &H80000009&
      X1              =   3960
      X2              =   3960
      Y1              =   6480
      Y2              =   5280
   End
   Begin VB.Line Line26 
      BorderColor     =   &H80000009&
      X1              =   7800
      X2              =   7800
      Y1              =   6480
      Y2              =   5280
   End
   Begin VB.Line Line24 
      BorderColor     =   &H80000009&
      X1              =   3960
      X2              =   3960
      Y1              =   5040
      Y2              =   3840
   End
   Begin VB.Line Line23 
      BorderColor     =   &H80000009&
      X1              =   7800
      X2              =   7800
      Y1              =   5040
      Y2              =   3840
   End
   Begin VB.Line Line21 
      BorderColor     =   &H80000009&
      X1              =   3960
      X2              =   3960
      Y1              =   3600
      Y2              =   2280
   End
   Begin VB.Line Line20 
      BorderColor     =   &H80000009&
      X1              =   7800
      X2              =   7800
      Y1              =   3600
      Y2              =   2280
   End
   Begin VB.Line Line18 
      BorderColor     =   &H80000009&
      X1              =   3960
      X2              =   3960
      Y1              =   2040
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   7800
      X2              =   7800
      Y1              =   2040
      Y2              =   840
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Armadura"
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
      Left            =   360
      TabIndex        =   18
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Casco"
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
      Left            =   360
      TabIndex        =   17
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Espada"
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
      Left            =   360
      TabIndex        =   16
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Arco"
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
      Left            =   360
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lanza"
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
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   1455
   End
   Begin VB.Line Line17 
      BorderColor     =   &H80000009&
      X1              =   3720
      X2              =   3720
      Y1              =   7920
      Y2              =   6720
   End
   Begin VB.Line Line16 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   240
      Y1              =   7920
      Y2              =   6720
   End
   Begin VB.Line Line15 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   3720
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000009&
      X1              =   3720
      X2              =   3720
      Y1              =   6480
      Y2              =   5280
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   240
      Y1              =   6480
      Y2              =   5280
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   3720
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      X1              =   3720
      X2              =   3720
      Y1              =   5040
      Y2              =   3840
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   240
      Y1              =   5040
      Y2              =   3840
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   3720
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      X1              =   3720
      X2              =   3720
      Y1              =   3600
      Y2              =   2280
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   240
      Y1              =   3600
      Y2              =   2280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   3720
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      X1              =   3720
      X2              =   3720
      Y1              =   2040
      Y2              =   840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   240
      Y1              =   2040
      Y2              =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   3720
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image Image7 
      Height          =   1155
      Left            =   4080
      Picture         =   "frmMercado.frx":10FAA
      Top             =   960
      Width           =   1080
   End
   Begin VB.Image Image3 
      Height          =   960
      Left            =   480
      Picture         =   "frmMercado.frx":150E4
      Top             =   2400
      Width           =   510
   End
   Begin VB.Image Image6 
      Height          =   1020
      Left            =   240
      Picture         =   "frmMercado.frx":16B26
      Top             =   8280
      Width           =   1020
   End
   Begin VB.Image Image5 
      Height          =   750
      Left            =   360
      Picture         =   "frmMercado.frx":1A198
      Top             =   5520
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   825
      Left            =   240
      Picture         =   "frmMercado.frx":1BEC2
      Top             =   6840
      Width           =   930
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   360
      Picture         =   "frmMercado.frx":1E768
      Top             =   3960
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   360
      Picture         =   "frmMercado.frx":20C6A
      Top             =   960
      Width           =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   240
      X2              =   7800
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Productos"
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
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image Image10 
      Height          =   1095
      Left            =   4080
      Picture         =   "frmMercado.frx":2312C
      Top             =   6720
      Width           =   1320
   End
End
Attribute VB_Name = "frmMercado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Oro >= 500 Then
    frmEjercito.cantLanzas = frmEjercito.cantLanzas + 1
    Oro = Oro - 500
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command10_Click()
If Oro >= 2000 Then
    frmEjercito.cantArmaduras = frmEjercito.cantArmaduras + 1
    Oro = Oro - 2000
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command11_Click()
If frmEjercito.cantEscudos > 0 Then
    Oro = Oro + 500
    frmEjercito.cantEscudos = frmEjercito.cantEscudos - 1
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command12_Click()
If Oro >= 2000 Then
    frmEjercito.cantEscudos = frmEjercito.cantEscudos + 1
    Oro = Oro - 2000
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command13_Click()
If frmAlmacen.lblMetal > 0 Then
    frmAlmacen.lblMetal = frmAlmacen.lblMetal - 1
    Oro = Oro + 100
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command14_Click()
If Oro >= 500 Then
    frmAlmacen.lblMetal = frmAlmacen.lblMetal + 1
    Oro = Oro - 500
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command15_Click()
If Oro >= 100 Then
    frmAlmacen.lblBloque = frmAlmacen.lblBloque + 1
    Oro = Oro - 100
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command16_Click()
If frmAlmacen.lblBloque > 0 Then
    frmAlmacen.lblBloque = frmAlmacen.lblBloque - 1
    Oro = Oro + 10
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command17_Click()
If frmAlmacen.lblMadera > 0 Then
    frmAlmacen.lblMadera = frmAlmacen.lblMadera - 1
    Oro = Oro + 50
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command18_Click()
If Oro >= 100 Then
    frmAlmacen.lblMadera = frmAlmacen.lblMadera + 1
    Oro = Oro - 100
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command19_Click()
If Oro >= 100 Then
    frmJuego.lblMadera = frmJuego.lblMadera + 100
    Oro = Oro - 100
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command2_Click()
If frmEjercito.cantLanzas > 0 Then
    frmEjercito.cantLanzas = frmEjercito.cantLanzas - 1
    Oro = Oro + 400
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command20_Click()
If frmJuego.lblMadera >= 100 Then
    frmJuego.lblMadera = frmJuego.lblMadera - 100
    Oro = Oro + 5
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command21_Click()
If Oro >= 100 Then
    frmJuego.lblPiedra = frmJuego.lblPiedra + 100
    Oro = Oro - 100
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command22_Click()
If frmJuego.lblPiedra >= 100 Then
    Oro = Oro + 5
    frmJuego.lblPiedra = frmJuego.lblPiedra - 100
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command3_Click()
If frmEjercito.cantArcos > 0 Then
    frmEjercito.cantArcos = frmEjercito.cantArcos - 1
    Oro = Oro + 300
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command4_Click()
If Oro >= 400 Then
    frmEjercito.cantArcos = frmEjercito.cantArcos + 1
    Oro = Oro - 400
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command5_Click()
If frmEjercito.cantEspadas > 0 Then
    frmEjercito.cantEspadas = frmEjercito.cantEspadas - 1
    Oro = Oro + 500
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command6_Click()
If Oro >= 600 Then
    frmEjercito.cantEspadas = frmEjercito.cantEspadas + 1
    Oro = Oro - 600
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command7_Click()
If frmEjercito.cantCascos > 0 Then
    frmEjercito.cantCascos = frmEjercito.cantCascos - 1
    Oro = Oro + 500
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command8_Click()
If Oro >= 2000 Then
    frmEjercito.cantCascos = frmEjercito.cantCascos + 1
    Oro = Oro - 2000
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Command9_Click()
If frmEjercito.cantArmaduras > 0 Then
    frmEjercito.cantArmaduras = frmEjercito.cantArmaduras - 1
    Oro = Oro + 500
    If Sonidos.Sonidos.playState = wmppsStopped Then
        Sonidos.Sonidos.URL = (App.Path & "\Sonidos\Compra.WAV")
    Else
        If Sonidos.Sonidos2.playState = wmppsStopped Then
            Sonidos.Sonidos2.URL = (App.Path & "\Sonidos\Compra.WAV")
        Else
            Sonidos.Sonidos3.URL = (App.Path & "\Sonidos\Compra.WAV")
        End If
    End If
End If
End Sub

Private Sub Form_Load()
    Me.MouseIcon = LoadPicture(App.Path & "\Punteros\Pesos.ico")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 2
    Me.Visible = False
End Sub

Private Sub Timer1_Timer()
    frmJuego.lblOro = Oro
    If frmJuego.lblOro >= 500 And Objetivos1 = True Then
        frmObjetivos.Check3.Value = 1
        frmObjetivos.Check3.Enabled = False
    End If
End Sub
