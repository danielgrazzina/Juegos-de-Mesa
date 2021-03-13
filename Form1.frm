VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menu"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "MV Boli"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   503
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   911
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12480
      Top             =   2040
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000C0C0&
      Caption         =   "Pise"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C000&
      Caption         =   "Damas"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "Sudoku"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "salir del programa"
      Height          =   1575
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":48316
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Alumno"
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   120
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7800
      TabIndex        =   8
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fecha"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12120
      TabIndex        =   7
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "reloj"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12120
      TabIndex        =   6
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    End 'comando para salir del programa
End Sub

Private Sub Command2_Click()
    Form1.Hide 'ocultar formulario
    Load Form2 'cargar formulario
    Form2.Show 'mostrar formulario
End Sub

Private Sub Command3_Click()
    Form1.Hide
    Load Form3
    Form3.Show
End Sub

Private Sub Command4_Click()
    Form1.Hide
    Load Form4
    Form4.Show
End Sub

Private Sub Form_Load()
    'se soloca el texto del label 1 y el label 2
    Label1.Caption = "Nombre: Daniel Grazzina" & vbNewLine _
        & "Cedula: 26.254.611" & vbNewLine & "Seccion: V-0401" & vbNewLine
    Label4.Caption = "Elige el juego que desees jugar" & vbNewLine
End Sub

Private Sub Timer1_Timer()
    Label2.Caption = Date 'este comado coloca la fecha en el label
    Label3.Caption = Time 'este comando coloca la hora en el label
End Sub

