VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H0000C0C0&
   Caption         =   "Pise"
   ClientHeight    =   9405
   ClientLeft      =   1590
   ClientTop       =   945
   ClientWidth     =   14775
   LinkTopic       =   "Form4"
   ScaleHeight     =   9405
   ScaleWidth      =   14775
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
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   2295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form4.Hide
    Load Form1
    Form1.Show
End Sub
