VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Master Mind XP  -  by  0x34"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Neuropol"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":058A
   ScaleHeight     =   7245
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer ETim 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2640
      Top             =   1320
   End
   Begin VB.Timer MT 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   1320
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1680
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2640
      Picture         =   "Form1.frx":87436
      ScaleHeight     =   735
      ScaleWidth      =   3015
      TabIndex        =   21
      Top             =   120
      Width           =   3015
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   960
      Top             =   1320
   End
   Begin VB.CommandButton Enter 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2520
      Picture         =   "Form1.frx":90186
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton About 
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3900
      Picture         =   "Form1.frx":9BABE
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3900
      Picture         =   "Form1.frx":9E5EE
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton NewGame 
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3900
      Picture         =   "Form1.frx":A111E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Col4 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Col4 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Col4 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Col4 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Col3 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Col3 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Col3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Col3 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Col2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Col2 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Col2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Col2 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Col1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   375
   End
   Begin VB.CommandButton Col1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton Col1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton Col1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Multiples 
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3900
      Picture         =   "Form1.frx":A3C4E
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label LabLoser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOSER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   23
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label LabWin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WINNER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   22
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Shape MultiInd 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      Height          =   135
      Left            =   3712
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape Row4 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   6
      Left            =   2895
      Top             =   4200
      Width           =   135
   End
   Begin VB.Shape Row4 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   2895
      Top             =   3720
      Width           =   135
   End
   Begin VB.Shape Row4 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   2895
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Row4 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   2895
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape Row4 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   2895
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Row4 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   2895
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Row4 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   2895
      Top             =   1320
      Width           =   135
   End
   Begin VB.Shape Row3 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   6
      Left            =   2730
      Top             =   4200
      Width           =   135
   End
   Begin VB.Shape Row3 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   2730
      Top             =   3720
      Width           =   135
   End
   Begin VB.Shape Row3 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   2730
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Row3 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   2730
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape Row3 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   2730
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Row3 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   2730
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Row3 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   2730
      Top             =   1320
      Width           =   135
   End
   Begin VB.Shape Row2 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   6
      Left            =   2565
      Top             =   4200
      Width           =   135
   End
   Begin VB.Shape Row2 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   2565
      Top             =   3720
      Width           =   135
   End
   Begin VB.Shape Row2 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   2565
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Row2 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   2565
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape Row2 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   2565
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Row2 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   2565
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Row2 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   2565
      Top             =   1320
      Width           =   135
   End
   Begin VB.Shape Row1 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   6
      Left            =   2400
      Top             =   4200
      Width           =   135
   End
   Begin VB.Shape Row1 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   2400
      Top             =   3720
      Width           =   135
   End
   Begin VB.Shape Row1 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   2400
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape R4 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   6
      Left            =   1920
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape R4 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   5
      Left            =   1920
      Top             =   3600
      Width           =   375
   End
   Begin VB.Shape R4 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   4
      Left            =   1920
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape R3 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   6
      Left            =   1440
      Shape           =   1  'Square
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape R3 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   5
      Left            =   1440
      Shape           =   1  'Square
      Top             =   3600
      Width           =   375
   End
   Begin VB.Shape R3 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   4
      Left            =   1440
      Shape           =   1  'Square
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape R2 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   6
      Left            =   960
      Shape           =   1  'Square
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape R2 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   5
      Left            =   960
      Shape           =   1  'Square
      Top             =   3600
      Width           =   375
   End
   Begin VB.Shape R2 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   4
      Left            =   960
      Shape           =   1  'Square
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape R1 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   6
      Left            =   480
      Shape           =   1  'Square
      Top             =   4080
      Width           =   375
   End
   Begin VB.Shape R1 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   5
      Left            =   480
      Shape           =   1  'Square
      Top             =   3600
      Width           =   375
   End
   Begin VB.Shape R1 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   4
      Left            =   480
      Shape           =   1  'Square
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape Key 
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   3
      Left            =   2280
      Top             =   350
      Width           =   375
   End
   Begin VB.Shape Key 
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   1800
      Top             =   350
      Width           =   375
   End
   Begin VB.Shape Key 
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   1320
      Top             =   350
      Width           =   375
   End
   Begin VB.Shape Key 
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   840
      Top             =   350
      Width           =   375
   End
   Begin VB.Shape R4 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   3
      Left            =   1920
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape R4 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   1920
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape R4 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   1920
      Top             =   1680
      Width           =   375
   End
   Begin VB.Shape R4 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   1920
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape R3 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   3
      Left            =   1440
      Shape           =   1  'Square
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape R3 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   1440
      Shape           =   1  'Square
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape R3 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   1440
      Shape           =   1  'Square
      Top             =   1680
      Width           =   375
   End
   Begin VB.Shape R3 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   1440
      Shape           =   1  'Square
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape R2 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   3
      Left            =   960
      Shape           =   1  'Square
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape R2 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   960
      Shape           =   1  'Square
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape R2 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   960
      Shape           =   1  'Square
      Top             =   1680
      Width           =   375
   End
   Begin VB.Shape R2 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   960
      Shape           =   1  'Square
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape R1 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   3
      Left            =   480
      Shape           =   1  'Square
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape R1 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   480
      Shape           =   1  'Square
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape R1 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   480
      Shape           =   1  'Square
      Top             =   1680
      Width           =   375
   End
   Begin VB.Shape R1 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   480
      Shape           =   1  'Square
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Row1 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   2400
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape Row1 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   2400
      Top             =   2280
      Width           =   135
   End
   Begin VB.Shape Row1 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   2400
      Top             =   1800
      Width           =   135
   End
   Begin VB.Shape Row1 
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   2400
      Top             =   1320
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'                              Master Mind XP
'                              Coded by 0x34
'                         GraphX by Michael Slater
'                              (0x34's Son)


Dim ClearRow1 As Boolean
Dim ClearRow2 As Boolean
Dim ClearRow3 As Boolean
Dim ClearRow4 As Boolean
Dim ClearRow1a As Boolean
Dim ClearRow2a As Boolean
Dim ClearRow3a As Boolean
Dim ClearRow4a As Boolean
Dim ClearPos As Byte
Dim RPos As Integer
Dim DRTR As Boolean
Dim CRed As Integer
Dim CGreen As Integer
Dim CBlue As Integer
Dim CYellow As Integer
Dim Winner As Boolean
Dim Cover As Boolean
Dim Movedone As Boolean
Dim Bang As Boolean
Dim LIT As Boolean
Dim ErrCnt As Integer
Dim E As Integer
Dim SelClr As String

Dim P(4) As Byte               'Array to hold the Hidden Key
'                        0 = Red   Colors are always in this order
'                        1 = Green
'                        2 = Yellow
'                        3 = Blue
Dim Sel(4, 9) As Byte
Dim KC(4) As Integer

Dim Start As Boolean
Dim Multi As Boolean

Private Sub About_Click() ' Show the splash screen
    frmSplash.Show
End Sub

Public Sub PlySnd() ' Main CLICK sound for button press
    SPly = sndPlaySound("C:\Program Files\MasterMind\Sounds\Click.wav", 1)
End Sub

Private Sub Col1_Click(Index As Integer) ' Color Entry KeyPress - Coloum #1
Select Case Index
    Case 0
        R1(RPos).BackColor = vbRed
        SelClr = vbRed
    Case 1
        R1(RPos).BackColor = vbGreen
        SelClr = vbGreen
    Case 2
        R1(RPos).BackColor = vbYellow
        SelClr = vbYellow
    Case 3
        R1(RPos).BackColor = vbBlue
End Select
SelClr = R1(RPos).BackColor
E = NoDubs(RPos, "R1", SelClr)
Sel(0, RPos) = Index
PlySnd
End Sub

Private Sub Col2_Click(Index As Integer) ' Color Entry KeyPress - Coloum #2
Select Case Index
    Case 0
        R2(RPos).BackColor = vbRed
    Case 1
        R2(RPos).BackColor = vbGreen
    Case 2
        R2(RPos).BackColor = vbYellow
    Case 3
        R2(RPos).BackColor = vbBlue
End Select
SelClr = R2(RPos).BackColor
E = NoDubs(RPos, "R2", SelClr)
Sel(1, RPos) = Index
PlySnd
End Sub

Private Sub Col3_Click(Index As Integer) ' Color Entry KeyPress - Coloum #3
Select Case Index
    Case 0
        R3(RPos).BackColor = vbRed
    Case 1
        R3(RPos).BackColor = vbGreen
    Case 2
        R3(RPos).BackColor = vbYellow
    Case 3
        R3(RPos).BackColor = vbBlue
End Select
SelClr = R3(RPos).BackColor
E = NoDubs(RPos, "R3", SelClr)
Sel(2, RPos) = Index
PlySnd
End Sub

Private Sub Col4_Click(Index As Integer) ' Color Entry KeyPress - Coloum #4
Select Case Index
    Case 0
        R4(RPos).BackColor = vbRed
    Case 1
        R4(RPos).BackColor = vbGreen
    Case 2
        R4(RPos).BackColor = vbYellow
    Case 3
        R4(RPos).BackColor = vbBlue
End Select
SelClr = R4(RPos).BackColor
E = NoDubs(RPos, "R4", SelClr)
Sel(3, RPos) = Index
PlySnd
End Sub

Private Function NoDubs(R As Integer, C As String, ByVal cl As Long) ' Single colors ONLY!
If Multi = True Then Exit Function
    If C <> "R1" Then
        If R1(R).BackColor = cl Then
            R1(R).BackColor = vbBlack
        End If
    End If
    If C <> "R2" Then
        If R2(R).BackColor = cl Then
            R2(R).BackColor = vbBlack
        End If
    End If
    If C <> "R3" Then
        If R3(R).BackColor = cl Then
            R3(R).BackColor = vbBlack
        End If
    End If
    If C <> "R4" Then
        If R4(R).BackColor = cl Then
            R4(R).BackColor = vbBlack
        End If
    End If
End Function

Private Sub Command1_Click() ' Statistics Form
    If Form2.Visible Then
        Unload Form2
    Else
        Form2.Show
    End If
End Sub

Private Sub ETim_Timer() ' Flash empty space timer
 ErrCnt = ErrCnt + 1
    If LIT Then
        DouseEm
    Else
        LiteEm
    End If
If ErrCnt = 5 Or ErrCnt > 5 Then
    ETim.Enabled = False
    DouseEm
End If
 
End Sub

Private Sub LiteEm() '   Lite unfilled spaces
    If R1(RPos).BackColor = vbBlack Then
        R1(RPos).BackColor = vbWhite
    End If
    If R2(RPos).BackColor = vbBlack Then
        R2(RPos).BackColor = vbWhite
    End If
    If R3(RPos).BackColor = vbBlack Then
        R3(RPos).BackColor = vbWhite
    End If
    If R4(RPos).BackColor = vbBlack Then
        R4(RPos).BackColor = vbWhite
    End If
    LIT = True
End Sub

Private Sub DouseEm() '   Unlite unfilled spaces
    If R1(RPos).BackColor = vbWhite Then
        R1(RPos).BackColor = vbBlack
    End If
    If R2(RPos).BackColor = vbWhite Then
        R2(RPos).BackColor = vbBlack
    End If
    If R3(RPos).BackColor = vbWhite Then
        R3(RPos).BackColor = vbBlack
    End If
    If R4(RPos).BackColor = vbWhite Then
        R4(RPos).BackColor = vbBlack
    End If
    LIT = False
End Sub

Private Sub Enter_Click() ' Enter Key Press - It all happens here
Dim t As Integer
Dim CW As Integer
Dim CCy As Integer

If ETim.Enabled = True Then Exit Sub

If R1(RPos).BackColor = vbBlack Or R2(RPos).BackColor = vbBlack _
Or R3(RPos).BackColor = vbBlack Or R4(RPos).BackColor = vbBlack Then
    ETim.Enabled = True
    ErrCnt = 0
    LiteEm
    Exit Sub
End If


    SPly = sndPlaySound("C:\Program Files\MasterMind\Sounds\Enter.wav", 1)
    CCy = 0
    CW = 0
    CRed = 0
    CGreen = 0
    CYellow = 0
    CBlue = 0
    Timer1.Enabled = False              'Turn off wig wag position indicator
    Row1(RPos).BackColor = vbBlack
    Row2(RPos).BackColor = vbBlack
    Row3(RPos).BackColor = vbBlack
    Row4(RPos).BackColor = vbBlack

    For t = 0 To 3                      'Determine how many of each color have been entered
        Select Case Sel(t, RPos)
            Case Is = 0
                CRed = CRed + 1
            Case Is = 1
                CGreen = CGreen + 1
            Case Is = 2
                CYellow = CYellow + 1
            Case Is = 3
                CBlue = CBlue + 1
        End Select
    Next
                                        'KC(x) is 1 x each color in key
    If KC(0) > 0 Then                   'Increment CW for each color matching the key colors
        For t = 0 To KC(0) - 1
            If CRed > 0 Then
                CW = CW + 1
                CRed = CRed - 1
            End If
        Next
    End If
    If KC(1) > 0 Then
        For t = 0 To KC(1) - 1
            If CGreen > 0 Then
                CW = CW + 1
                CGreen = CGreen - 1
            End If
        Next
    End If
    If KC(2) > 0 Then
        For t = 0 To KC(2) - 1
            If CYellow > 0 Then
                CW = CW + 1
                CYellow = CYellow - 1
            End If
        Next
    End If
    If KC(3) > 0 Then
        For t = 0 To KC(3) - 1
            If CBlue > 0 Then
                CW = CW + 1
                CBlue = CBlue - 1
            End If
        Next
    End If
 
    If CW > 0 Then                      'Colorize boxes white, indicating correct colors
        Row1(RPos).BackColor = vbWhite
        CW = CW - 1
    End If
    If CW > 0 Then
        Row2(RPos).BackColor = vbWhite
        CW = CW - 1
    End If
    If CW > 0 Then
        Row3(RPos).BackColor = vbWhite
        CW = CW - 1
    End If
    If CW > 0 Then
        Row4(RPos).BackColor = vbWhite
        CW = CW - 1
    End If
    
    If R1(RPos).BackColor = Key(0).BackColor Then   'Increment CCY for each color in the right spot
        CCy = CCy + 1
    End If
    If R2(RPos).BackColor = Key(1).BackColor Then
        CCy = CCy + 1
    End If
    If R3(RPos).BackColor = Key(2).BackColor Then
        CCy = CCy + 1
    End If
    If R4(RPos).BackColor = Key(3).BackColor Then
        CCy = CCy + 1
    End If

    For t = 0 To CCy - 1                            'Make boxes blue for respective amount
        If Row1(RPos).BackColor <> &HFFFF80 Then    'of colors in the right spot.
            Row1(RPos).BackColor = &HFFFF80
        ElseIf Row2(RPos).BackColor <> &HFFFF80 Then
            Row2(RPos).BackColor = &HFFFF80
        ElseIf Row3(RPos).BackColor <> &HFFFF80 Then
            Row3(RPos).BackColor = &HFFFF80
        ElseIf Row4(RPos).BackColor <> &HFFFF80 Then
            Row4(RPos).BackColor = &HFFFF80
        End If
    Next
    Cover = False
    If CCy = 4 Then '   If all colors are correct (CCy = 4) then: WINNER
        SPly = sndPlaySound("C:\Program Files\MasterMind\Sounds\Win.wav", 1)
        Wins = Wins + 1             'Increment Wins Stat
        RPos = 0                    'Zero the Row Position
        Winner = True               'Set Winner reg
        DisableInput                'Disable the input buttons
        Timer3.Enabled = True       'Enable result timer to flash the result
        LabWin.Visible = True       'Display "Winner"
        MT.Enabled = True
        Exit Sub
    End If

    RPos = RPos + 1                 'Increment Row Position
    
    If RPos > 6 Then    '  If colors not correct, and last try exceeded, then: LOSER
        SPly = sndPlaySound("C:\Program Files\MasterMind\Sounds\Lose.wav", 1)
        Losses = Losses + 1         'Increment Losses Stat
        RPos = 0                    'Zero the Row Position
        Winner = False              'Clear Winner reg
        DisableInput                'Disable the input buttons
        Timer3.Enabled = True       'Enable result timer to flash the result
        LabLoser.Visible = True     'Display "Loser"
        MT.Enabled = True
        Exit Sub
    End If
    
    Row1(RPos).BackColor = &HFFFF80 'Set new wig wag start position
    Timer1.Enabled = True           'Restart wig wag place indicator
    
End Sub

Private Sub Form_Unload(Cancel As Integer) ' Exit program correctly
    If frmSplash.Visible Then Unload frmSplash
    If Form2.Visible Then Unload Form2
    Unload Me
    End
End Sub

Private Sub MT_Timer() ' Move key cover
If Cover Then
    If Picture1.Left <= 235 Then
        MT.Enabled = False
        NuGam
        Exit Sub
    End If
    Picture1.Left = Picture1.Left - 50
Else
    If Picture1.Left >= 2640 Then
        MT.Enabled = False
        Exit Sub
    End If
    Picture1.Left = Picture1.Left + 50
End If
    
End Sub

Private Sub Multiples_Click()   ' Toggle Multiples Option
If RPos > 0 Then
    E = MsgBox("Start new game?", vbYesNo, "Change Game Type")
    If E = 7 Then Exit Sub
End If
    SPly = sndPlaySound("C:\Program Files\MasterMind\Sounds\StatClick.wav", 1)
    If Multi = False Then
        MultiInd.BackColor = vbGreen
        Multi = True
    Else
        MultiInd.BackColor = vbBlack
        Multi = False
    End If
    NewGame_Click
End Sub

Private Sub Form_Load()
    LabWin.Visible = False
    LabLoser.Visible = False
    DisableInput
    frmSplash.Show
    Timer1.Interval = 2000  'Value in milliseconds to display splash screen
    Timer1.Enabled = True
    Form1.Visible = False
    Start = True
    Picture1.Left = 235
    Picture1.Top = 120
    DRTR = True
    Randomize
End Sub

Public Function DisableInput() ' Disable Coloum Keys and Enter Key
Dim t As Integer
    For t = 0 To 3
        Col1(t).Enabled = False
    Next
    For t = 0 To 3
        Col2(t).Enabled = False
    Next
        For t = 0 To 3
        Col3(t).Enabled = False
    Next
        For t = 0 To 3
        Col4(t).Enabled = False
    Next
    Enter.Enabled = False
End Function
Public Function EnableInput() ' Enable Coloum Keys and Enter Key
Dim t As Integer
    For t = 0 To 3
        Col1(t).Enabled = True
    Next
    For t = 0 To 3
        Col2(t).Enabled = True
    Next
        For t = 0 To 3
        Col3(t).Enabled = True
    Next
        For t = 0 To 3
        Col4(t).Enabled = True
    Next
    Enter.Enabled = True
End Function

Private Sub NewGame_Click() ' New Game
    If Picture1.Left > 235 Then
        SPly = sndPlaySound("C:\Program Files\MasterMind\Sounds\close.wav", 1)
        Bang = True
    End If
    Movedone = False
    Cover = True
    MT.Enabled = True
End Sub

Public Sub NuGam() ' Main New Game set-up routine
Dim i As Integer
Dim t As Integer
    If RPos > 0 Then    ' If a play has been made and the game is reset, it counts as a loss!
        Losses = Losses + 1
    End If
    If Bang Then
        SPly = sndPlaySound("C:\Program Files\MasterMind\Sounds\NewGameb.wav", 1)
    Else
        SPly = sndPlaySound("C:\Program Files\MasterMind\Sounds\NewGame.wav", 1)
    End If
    Bang = False
    Timer3.Enabled = False
    LabLoser.Visible = False
    LabWin.Visible = False
    Winner = False
    For t = 0 To 3
        KC(t) = 0
    Next
    Timer1.Enabled = False
    NewGame.Enabled = False
    Randomize
    ClearRow1 = False
    ClearRow2 = False
    ClearRow3 = False
    ClearRow4 = False
    ClearRow1a = False
    ClearRow2a = False
    ClearRow3a = False
    ClearRow4a = False
    DRTR = True
    ClearPos = 7
    Timer2.Enabled = True
    Picture1.Visible = True
    RPos = 0
    SelectColors
    For t = 0 To 8
        For i = 0 To 4
            Sel(i, t) = 5
        Next
    Next
End Sub

Public Sub SelectColors() ' Randomly Pick Key colors
Dim t As Byte
Randomize
    If Multi Then   ' Set for Multiples (Easier to code)
        For t = 0 To 3
            P(t) = Fix(Rnd * 4)
        Next
        SetColor
        Exit Sub
    End If
    For t = 0 To 3  'Set for Non-Multiples  (Harder to code)
        P(t) = 5
    Next
    P(0) = Fix(Rnd * 4)
MIX:  'Mix it up  -  Loop untill all colors are not duplicates
    If P(1) = P(0) Or P(1) = 5 Then
       P(1) = Fix(Rnd * 4)
        GoTo MIX
    End If
    If P(2) = P(0) Or P(2) = P(1) Or P(2) = 5 Then
        P(2) = Fix(Rnd * 4)
        GoTo MIX
    End If
    If P(3) = P(0) Or P(3) = P(1) Or P(3) = P(2) Or P(3) = 5 Then
        P(3) = Fix(Rnd * 4)
        GoTo MIX
    End If
    SetColor
End Sub

Public Sub SetColor() 'Place colores under hidden key
Dim t As Byte
    For t = 0 To 3
        If P(t) = 0 Then
            Key(t).BackColor = vbRed
            KC(0) = KC(0) + 1
        End If
        If P(t) = 1 Then
            Key(t).BackColor = vbGreen
            KC(1) = KC(1) + 1
        End If
        If P(t) = 2 Then
            Key(t).BackColor = vbYellow
            KC(2) = KC(2) + 1
        End If
        If P(t) = 3 Then
            Key(t).BackColor = vbBlue
            KC(3) = KC(3) + 1
        End If
    Next
End Sub

Private Sub Timer1_Timer()
    If Start = True Then    'Splash Screen Control
        Timer1.Interval = 75    'Wig Wag position indicator speed
        Timer1.Enabled = False
        Unload frmSplash
        Form1.Visible = True
        Start = False
        Exit Sub
    End If
If DRTR = True Then '                         Cool indicator for play position
    If Row1(RPos).BackColor = &HFFFF80 Then
        Row1(RPos).BackColor = vbBlack
        Row2(RPos).BackColor = &HFFFF80
        Exit Sub
    End If
    If Row2(RPos).BackColor = &HFFFF80 Then
        Row2(RPos).BackColor = vbBlack
       Row3(RPos).BackColor = &HFFFF80
        Exit Sub
    End If
        If Row3(RPos).BackColor = &HFFFF80 Then
        Row3(RPos).BackColor = vbBlack
        Row4(RPos).BackColor = &HFFFF80
        Exit Sub
    End If
        If Row4(RPos).BackColor = &HFFFF80 Then
        Row4(RPos).BackColor = vbBlack
        Row3(RPos).BackColor = &HFFFF80
        DRTR = False
        Exit Sub
    End If
Else
    If Row3(RPos).BackColor = &HFFFF80 Then
        Row3(RPos).BackColor = vbBlack
        Row2(RPos).BackColor = &HFFFF80
        Exit Sub
    End If
    If Row2(RPos).BackColor = &HFFFF80 Then
        Row2(RPos).BackColor = vbBlack
        Row1(RPos).BackColor = &HFFFF80
        Exit Sub
    End If
    If Row1(RPos).BackColor = &HFFFF80 Then
        Row1(RPos).BackColor = vbBlack
        Row2(RPos).BackColor = &HFFFF80
        DRTR = True
        Exit Sub
    End If
End If
End Sub

Private Sub Timer2_Timer()  ' Clear all for a new game
    If ClearRow1 = False Then
        ClearPos = ClearPos - 1
        R1(ClearPos).BackColor = vbWhite
        Row1(ClearPos).BackColor = vbWhite
        If ClearPos = 0 Then
            ClearPos = 7
            ClearRow1 = True
        End If
        Exit Sub
    End If
    If ClearRow2 = False Then
        ClearPos = ClearPos - 1
        R2(ClearPos).BackColor = vbWhite
        Row2(ClearPos).BackColor = vbWhite
        If ClearPos = 0 Then
            ClearPos = 7
            ClearRow2 = True
        End If
        Exit Sub
    End If
    If ClearRow3 = False Then
        ClearPos = ClearPos - 1
        R3(ClearPos).BackColor = vbWhite
        Row3(ClearPos).BackColor = vbWhite
        If ClearPos = 0 Then
            ClearPos = 7
            ClearRow3 = True
        End If
        Exit Sub
    End If
    If ClearRow4 = False Then
        ClearPos = ClearPos - 1
        R4(ClearPos).BackColor = vbWhite
        Row4(ClearPos).BackColor = vbWhite
        If ClearPos = 0 Then
            ClearPos = 7
            ClearRow4 = True
        End If
        Exit Sub
    End If
    '-------------------------------------(Black)----------
       If ClearRow1a = False Then
        ClearPos = ClearPos - 1
        R1(ClearPos).BackColor = vbBlack
        Row1(ClearPos).BackColor = vbBlack
        If ClearPos = 0 Then
            ClearPos = 7
            ClearRow1a = True
        End If
        Exit Sub
    End If
    If ClearRow2a = False Then
        ClearPos = ClearPos - 1
        R2(ClearPos).BackColor = vbBlack
        Row2(ClearPos).BackColor = vbBlack
        If ClearPos = 0 Then
            ClearPos = 7
            ClearRow2a = True
        End If
        Exit Sub
    End If
    If ClearRow3a = False Then
        ClearPos = ClearPos - 1
        R3(ClearPos).BackColor = vbBlack
        Row3(ClearPos).BackColor = vbBlack
        If ClearPos = 0 Then
            ClearPos = 7
            ClearRow3a = True
        End If
        Exit Sub
    End If
    If ClearRow4a = False Then
        ClearPos = ClearPos - 1
        R4(ClearPos).BackColor = vbBlack
        Row4(ClearPos).BackColor = vbBlack
        If ClearPos = 0 Then
            ClearPos = 7
            ClearRow4a = True
            Timer2.Enabled = False
            NewGame.Enabled = True
        End If
        Row1(RPos).BackColor = &HFFFF80
        Timer1.Enabled = True
        EnableInput
        Exit Sub
    End If
End Sub

Private Sub Timer3_Timer() 'Win/Lose Splash timer
    If Winner Then
        If LabWin.Visible Then
            LabWin.Visible = False
        Else
            LabWin.Visible = True
        End If
    Else
        If LabLoser.Visible Then
            LabLoser.Visible = False
        Else
            LabLoser.Visible = True
        End If
    End If
End Sub
