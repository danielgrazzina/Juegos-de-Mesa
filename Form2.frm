VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C000&
   Caption         =   "Sudoku"
   ClientHeight    =   8820
   ClientLeft      =   2565
   ClientTop       =   1440
   ClientWidth     =   16140
   FillColor       =   &H00C0C0C0&
   FontTransparent =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8820
   ScaleWidth      =   16140
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "Regresar al Menu"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   11160
      TabIndex        =   82
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   7560
      TabIndex        =   81
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   79
      Left            =   10560
      TabIndex        =   79
      Text            =   "0"
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   78
      Left            =   9960
      TabIndex        =   78
      Text            =   "0"
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   80
      Left            =   11160
      TabIndex        =   77
      Text            =   "0"
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   77
      Left            =   9360
      TabIndex        =   76
      Text            =   "0"
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   76
      Left            =   8760
      TabIndex        =   75
      Text            =   "0"
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   75
      Left            =   8160
      TabIndex        =   74
      Text            =   "0"
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   74
      Left            =   7560
      TabIndex        =   73
      Text            =   "0"
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   73
      Left            =   6960
      TabIndex        =   72
      Text            =   "0"
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   72
      Left            =   6360
      TabIndex        =   71
      Text            =   "0"
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   71
      Left            =   11160
      TabIndex        =   70
      Text            =   "0"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   70
      Left            =   10560
      TabIndex        =   69
      Text            =   "0"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   69
      Left            =   9960
      TabIndex        =   68
      Text            =   "0"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   68
      Left            =   9360
      TabIndex        =   67
      Text            =   "0"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   67
      Left            =   8760
      TabIndex        =   66
      Text            =   "0"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   66
      Left            =   8160
      TabIndex        =   65
      Text            =   "0"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   65
      Left            =   7560
      TabIndex        =   64
      Text            =   "0"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   64
      Left            =   6960
      TabIndex        =   63
      Text            =   "0"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   63
      Left            =   6360
      TabIndex        =   62
      Text            =   "0"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   62
      Left            =   11160
      TabIndex        =   61
      Text            =   "0"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   61
      Left            =   10560
      TabIndex        =   60
      Text            =   "0"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   60
      Left            =   9960
      TabIndex        =   59
      Text            =   "0"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   59
      Left            =   9360
      TabIndex        =   58
      Text            =   "0"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   58
      Left            =   8760
      TabIndex        =   57
      Text            =   "0"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   57
      Left            =   8160
      TabIndex        =   56
      Text            =   "0"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   56
      Left            =   7560
      TabIndex        =   55
      Text            =   "0"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   55
      Left            =   6960
      TabIndex        =   54
      Text            =   "0"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   54
      Left            =   6360
      TabIndex        =   53
      Text            =   "0"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   53
      Left            =   11160
      TabIndex        =   52
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   52
      Left            =   10560
      TabIndex        =   51
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   51
      Left            =   9960
      TabIndex        =   50
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   50
      Left            =   9360
      TabIndex        =   49
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   49
      Left            =   8760
      TabIndex        =   48
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   48
      Left            =   8160
      TabIndex        =   47
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   47
      Left            =   7560
      TabIndex        =   46
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   46
      Left            =   6960
      TabIndex        =   45
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   45
      Left            =   6360
      TabIndex        =   44
      Text            =   "0"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   44
      Left            =   11160
      TabIndex        =   43
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   43
      Left            =   10560
      TabIndex        =   42
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   42
      Left            =   9960
      TabIndex        =   41
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   41
      Left            =   9360
      TabIndex        =   40
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   40
      Left            =   8760
      TabIndex        =   39
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   39
      Left            =   8160
      TabIndex        =   38
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   38
      Left            =   7560
      TabIndex        =   37
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   37
      Left            =   6960
      TabIndex        =   36
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   36
      Left            =   6360
      TabIndex        =   35
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   35
      Left            =   11160
      TabIndex        =   34
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   34
      Left            =   10560
      TabIndex        =   33
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   33
      Left            =   9960
      TabIndex        =   32
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   32
      Left            =   9360
      TabIndex        =   31
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   31
      Left            =   8760
      TabIndex        =   30
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   30
      Left            =   8160
      TabIndex        =   29
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   29
      Left            =   7560
      TabIndex        =   28
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   28
      Left            =   6960
      TabIndex        =   27
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   27
      Left            =   6360
      TabIndex        =   26
      Text            =   "0"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   26
      Left            =   11160
      TabIndex        =   25
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   25
      Left            =   10560
      TabIndex        =   24
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   24
      Left            =   9960
      TabIndex        =   23
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   9360
      TabIndex        =   22
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   22
      Left            =   8760
      TabIndex        =   21
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   21
      Left            =   8160
      TabIndex        =   20
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   20
      Left            =   7560
      TabIndex        =   19
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   19
      Left            =   6960
      TabIndex        =   18
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   18
      Left            =   6360
      TabIndex        =   17
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   11160
      TabIndex        =   16
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   10560
      TabIndex        =   15
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   9960
      TabIndex        =   14
      Text            =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   9360
      TabIndex        =   13
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   8760
      TabIndex        =   12
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   8160
      TabIndex        =   11
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   7560
      TabIndex        =   10
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   6960
      TabIndex        =   9
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   6360
      TabIndex        =   8
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   10560
      TabIndex        =   7
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   9960
      TabIndex        =   6
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   9360
      TabIndex        =   5
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   8760
      TabIndex        =   4
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   8160
      TabIndex        =   3
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   6360
      TabIndex        =   2
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox celda 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   6960
      TabIndex        =   1
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "finalizar"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1200
      TabIndex        =   0
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "iniciar juego"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1200
      OLEDropMode     =   1  'Manual
      TabIndex        =   80
      Top             =   720
      Width           =   2295
   End
   Begin VB.Line Line4 
      BorderWidth     =   13
      X1              =   11880
      X2              =   11880
      Y1              =   1200
      Y2              =   6840
   End
   Begin VB.Line Line3 
      BorderWidth     =   13
      X1              =   6240
      X2              =   6240
      Y1              =   6840
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderWidth     =   13
      X1              =   11880
      X2              =   6240
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderWidth     =   13
      X1              =   6240
      X2              =   11880
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit 'esta opcion es para la declaracion explicita de las varibles
Dim matrix(8, 8) As Integer 'se declara una matriz de 9x9

Private Sub Command1_Click()
    Form2.Hide
    Load Form1
    Form1.Show
End Sub

Private Sub Command2_Click()
    Randomize 'se inicia la funcion para obtener un numero aleatorio
    Dim vector(9), x, y, r, a, z As Integer 'se declaran las variables a utilizar
    
    For x = 0 To 8
       vector(x) = x + 1 'se llena el vector del 1 al 9
    Next x
    
    x = x - 1 'se le resta 1 para que la varible quede en 8 y se inicializa y
    y = 0
    
    Do While x > -1 'se usa la variable x como contador del 8 al 0
       r = Int((x - 0 + 1) * Rnd + 0) 'se almacena un numero randon entre 0 y la x que es el contador que va decreciendo en la variable r
       matrix(0, y) = vector(r) 'rellena la fila 0 con un numero randon
       For a = r To x 'este for es para ir vaciado el vector y que no se repitan los numeros
          vector(a) = vector(a + 1)
       Next a
       x = x - 1
       y = y + 1
    Loop
    'se construye el resto de la matriz en base a la primera fila cambiando las celdas de las matrices pequeñas de 3x3
    'respetando el orden correcto en la matriz 9x9 para armar un sudoku de con solucion unica
    '1
     matrix(1, 0) = matrix(0, 3)
     matrix(1, 1) = matrix(0, 4)
     matrix(1, 2) = matrix(0, 5)
     '2
     matrix(2, 0) = matrix(0, 6)
     matrix(2, 1) = matrix(0, 7)
     matrix(2, 2) = matrix(0, 8)
     
     '3
     matrix(1, 3) = matrix(0, 6)
     matrix(1, 4) = matrix(0, 7)
     matrix(1, 5) = matrix(0, 8)
     '4
     matrix(2, 3) = matrix(0, 0)
     matrix(2, 4) = matrix(0, 1)
     matrix(2, 5) = matrix(0, 2)
     
     '5
     matrix(1, 6) = matrix(0, 0)
     matrix(1, 7) = matrix(0, 1)
     matrix(1, 8) = matrix(0, 2)
     '6
     matrix(2, 6) = matrix(0, 3)
     matrix(2, 7) = matrix(0, 4)
     matrix(2, 8) = matrix(0, 5)
     
     '7
     matrix(3, 0) = matrix(0, 1)
     matrix(3, 1) = matrix(0, 2)
     matrix(3, 2) = matrix(1, 0)
     '8
     matrix(4, 0) = matrix(1, 1)
     matrix(4, 1) = matrix(1, 2)
     matrix(4, 2) = matrix(2, 0)
     '9
     matrix(5, 0) = matrix(2, 1)
     matrix(5, 1) = matrix(2, 2)
     matrix(5, 2) = matrix(0, 0)
     
     '10
     matrix(3, 3) = matrix(4, 0)
     matrix(3, 4) = matrix(4, 1)
     matrix(3, 5) = matrix(4, 2)
     '11
     matrix(4, 3) = matrix(5, 0)
     matrix(4, 4) = matrix(5, 1)
     matrix(4, 5) = matrix(5, 2)
     '12
     matrix(5, 3) = matrix(3, 0)
     matrix(5, 4) = matrix(3, 1)
     matrix(5, 5) = matrix(3, 2)
     
     '13
     matrix(3, 6) = matrix(5, 0)
     matrix(3, 7) = matrix(5, 1)
     matrix(3, 8) = matrix(5, 2)
     '14
     matrix(4, 6) = matrix(3, 0)
     matrix(4, 7) = matrix(3, 1)
     matrix(4, 8) = matrix(3, 2)
     '15
     matrix(5, 6) = matrix(4, 0)
     matrix(5, 7) = matrix(4, 1)
     matrix(5, 8) = matrix(4, 2)
     
     '16
     matrix(6, 0) = matrix(3, 1)
     matrix(6, 1) = matrix(3, 2)
     matrix(6, 2) = matrix(4, 0)
     '17
     matrix(7, 0) = matrix(4, 1)
     matrix(7, 1) = matrix(4, 2)
     matrix(7, 2) = matrix(5, 0)
     '18
     matrix(8, 0) = matrix(5, 1)
     matrix(8, 1) = matrix(5, 2)
     matrix(8, 2) = matrix(3, 0)
     
     '19
     matrix(6, 3) = matrix(7, 0)
     matrix(6, 4) = matrix(7, 1)
     matrix(6, 5) = matrix(7, 2)
     '20
     matrix(7, 3) = matrix(8, 0)
     matrix(7, 4) = matrix(8, 1)
     matrix(7, 5) = matrix(8, 2)
     '21
     matrix(8, 3) = matrix(6, 0)
     matrix(8, 4) = matrix(6, 1)
     matrix(8, 5) = matrix(6, 2)
     
     '22
     matrix(6, 6) = matrix(8, 0)
     matrix(6, 7) = matrix(8, 1)
     matrix(6, 8) = matrix(8, 2)
     '23
     matrix(7, 6) = matrix(6, 0)
     matrix(7, 7) = matrix(6, 1)
     matrix(7, 8) = matrix(6, 2)
     '24
     matrix(8, 6) = matrix(7, 0)
     matrix(8, 7) = matrix(7, 1)
     matrix(8, 8) = matrix(7, 2)
     
    z = 0 'esta variable es para aumentar el indice del textbox
    For x = 0 To 8 'este for se usa para vaciar textbox ramdon
       For y = 0 To 8
          r = Int((1 - 0 + 1) * Rnd + 0) ' este comando es par obtener un numaro radom del 0 al 1
          If r = 0 Then 'si r es 0 se vacia el textbox
             celda(z) = ""
          Else
             celda(z) = matrix(x, y) 'si no se le coloca el valor respctivo de el sudoku generado
          End If
          z = z + 1
       Next y
    Next x
End Sub

Private Sub Command3_Click()
    Dim c, x, y, z As Integer
    'se usa un bucle para comprobar si el sudoku es correcto
    c = 0
    z = 0
    For x = 0 To 8
        For y = 0 To 8
         If celda(z) <> "" Then ' comprueba si el textbox esta vacio
            If celda(z) = matrix(x, y) Then 'se compara si los textbox coiciden con la matriz
               c = c + 1 'si coicide se le suma uno
            End If
         End If
          z = z + 1 'z se usa como contador para recorrer las celdas
        Next y
     Next x
    If c = 81 Then
       MsgBox "Sudoku Correctamente", , "Felicidades"
    Else
       MsgBox "Sudoku Incorrecto...", , ""
    End If
End Sub

Private Sub Form2_Load()
Dim x As Integer

For x = 0 To 80 'se uso este bucle para cambiar la propiedad del textbox
   celda(x) = ""
   celda(x).MaxLength = 1 'esta propiedad limita los caracteres del textbox a el numero deseado
Next x
End Sub

