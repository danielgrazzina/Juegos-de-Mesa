VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H0000C000&
   Caption         =   "Damas"
   ClientHeight    =   9675
   ClientLeft      =   1905
   ClientTop       =   945
   ClientWidth     =   15540
   LinkTopic       =   "Form3"
   ScaleHeight     =   9675
   ScaleWidth      =   15540
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4440
      Top             =   8280
   End
   Begin VB.CommandButton Command3 
      Caption         =   "confirmar movimiento"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      TabIndex        =   2
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4440
      Top             =   7680
   End
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
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "iniciar partida"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5160
      TabIndex        =   0
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   4
      Top             =   7920
      Width           =   4095
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   6015
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   62
      Left            =   11520
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   60
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   58
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   56
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   55
      Left            =   12120
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   53
      Left            =   10920
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   51
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   49
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   46
      Left            =   11520
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   44
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   42
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   40
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   39
      Left            =   12120
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   38
      Left            =   11520
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   37
      Left            =   10920
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   36
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   35
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   34
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   33
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   32
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   31
      Left            =   12120
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   30
      Left            =   11520
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   29
      Left            =   10920
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   28
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   27
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   26
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   25
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   24
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   23
      Left            =   12120
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   21
      Left            =   10920
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   19
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   17
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   14
      Left            =   11520
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   12
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   10
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   7
      Left            =   12120
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   5
      Left            =   10920
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   3
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   1
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   8
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   41
      Left            =   8520
      Picture         =   "Form3.frx":0F09
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   43
      Left            =   9720
      Picture         =   "Form3.frx":1C00
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   45
      Left            =   10920
      Picture         =   "Form3.frx":28F7
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   47
      Left            =   12120
      Picture         =   "Form3.frx":35EE
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   48
      Left            =   7920
      Picture         =   "Form3.frx":42E5
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   50
      Left            =   9120
      Picture         =   "Form3.frx":4FDC
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   52
      Left            =   10320
      Picture         =   "Form3.frx":5CD3
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   54
      Left            =   11520
      Picture         =   "Form3.frx":69CA
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   63
      Left            =   12120
      Picture         =   "Form3.frx":76C1
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   61
      Left            =   10920
      Picture         =   "Form3.frx":83B8
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   59
      Left            =   9720
      Picture         =   "Form3.frx":90AF
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   57
      Left            =   8520
      Picture         =   "Form3.frx":9DA6
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   22
      Left            =   11520
      Picture         =   "Form3.frx":AA9D
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   20
      Left            =   10320
      Picture         =   "Form3.frx":B9D5
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   18
      Left            =   9120
      Picture         =   "Form3.frx":C90D
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   16
      Left            =   7920
      Picture         =   "Form3.frx":D845
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   15
      Left            =   12120
      Picture         =   "Form3.frx":E77D
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   6
      Left            =   11520
      Picture         =   "Form3.frx":F6B5
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   13
      Left            =   10920
      Picture         =   "Form3.frx":105ED
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   4
      Left            =   10320
      Picture         =   "Form3.frx":11525
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   11
      Left            =   9720
      Picture         =   "Form3.frx":1245D
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   2
      Left            =   9120
      Picture         =   "Form3.frx":13395
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   9
      Left            =   8520
      Picture         =   "Form3.frx":142CD
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image celda 
      Height          =   615
      Index           =   0
      Left            =   7920
      Picture         =   "Form3.frx":15205
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image tablero 
      Enabled         =   0   'False
      Height          =   6105
      Left            =   7320
      Picture         =   "Form3.frx":1613D
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   6015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' este comando es para declarar explicitamente las variables
Dim matriz(7, 7) As Integer ' se declara una matriz de 8x8
Dim moverfi, comerfi, turno, moverde, movera, x, y, a, b, z, l, c, ca, cb, reinicio As Integer



Private Sub celda_Click(Index As Integer) 'este es el vector del objeto imagen para poder seleccionar el lugar par moverte
    z = Index ' le da el valor del indice del objeto imagen
    Select Case z
        Case 0
            If matriz(0, 0) = 4 Then ' segun la posicion en la matriz se le asigna un valor a las variables 4 son las fichas blancas
                x = 0 ' las variables x y y se usan para la posicion de las fichas
                y = 0
                moverde = 4 ' moverde en 4 es para mover una ficha blanca
                l = z ' para guardar el valor del indice que se usara para saber la posicion de la imagen a cambiar
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(0, 0) = 5 Then ' el 5 son las fichas negras
                   x = 0
                   y = 0
                   moverde = 5 ' el 5 es para mover ficha negra
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(0, 0) = 3 Then 'en el caso de que sea 3 es que la posicion esta vacia
                        a = 0 ' a y b son para las cordenadas a donde se movera
                        b = 0
                        movera = 3 'esta variable esta la que indica que el lugar a donde se mueve la ficha esta vacia
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 1 ' todas estas estructuras fucionan igual pero para cado posicion en concreto
            If matriz(0, 1) = 4 Then
                x = 0
                y = 1
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(0, 1) = 5 Then
                   x = 0
                   y = 1
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(0, 1) = 3 Then
                        a = 0
                        b = 1
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 2
            If matriz(0, 2) = 4 Then
                x = 0
                y = 2
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(0, 2) = 5 Then
                   x = 0
                   y = 2
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(0, 2) = 3 Then
                        a = 0
                        b = 2
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 3
            If matriz(0, 3) = 4 Then
                x = 0
                y = 3
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(0, 3) = 5 Then
                   x = 0
                   y = 3
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(0, 3) = 3 Then
                        a = 0
                        b = 3
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 4
            If matriz(0, 4) = 4 Then
                x = 0
                y = 4
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(0, 4) = 5 Then
                   x = 0
                   y = 4
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(0, 4) = 3 Then
                        a = 0
                        b = 4
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 5
            If matriz(0, 5) = 4 Then
                    x = 0
                    y = 5
                    moverde = 4
                    l = z
                    Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
                Else
                    If matriz(0, 5) = 5 Then
                       x = 0
                       y = 5
                       moverde = 5
                       l = z
                       Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                    Else
                        If matriz(0, 5) = 3 Then
                            a = 0
                            b = 5
                            movera = 3
                            Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                        End If
                    End If
                End If
        Case 6
            If matriz(0, 6) = 4 Then
                x = 0
                y = 6
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(0, 6) = 5 Then
                   x = 0
                   y = 6
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(0, 6) = 3 Then
                        a = 0
                        b = 6
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 7
            If matriz(0, 7) = 4 Then
                x = 0
                y = 7
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(0, 7) = 5 Then
                   x = 0
                   y = 7
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(0, 7) = 3 Then
                        a = 0
                        b = 7
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 8
            If matriz(1, 0) = 4 Then
                x = 1
                y = 0
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(1, 0) = 5 Then
                   x = 1
                   y = 0
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(1, 0) = 3 Then
                        a = 1
                        b = 0
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 9
            If matriz(1, 1) = 4 Then
                x = 1
                y = 1
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(1, 1) = 5 Then
                   x = 1
                   y = 1
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(1, 1) = 3 Then
                        a = 1
                        b = 1
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 10
            If matriz(1, 2) = 4 Then
                x = 1
                y = 2
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(1, 2) = 5 Then
                   x = 1
                   y = 2
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(1, 2) = 3 Then
                        a = 1
                        b = 2
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 11
            If matriz(1, 3) = 4 Then
                x = 1
                y = 3
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(1, 3) = 5 Then
                   x = 1
                   y = 3
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(1, 3) = 3 Then
                        a = 1
                        b = 3
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 12
            If matriz(1, 4) = 4 Then
                x = 1
                y = 4
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(1, 4) = 5 Then
                   x = 1
                   y = 4
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(1, 4) = 3 Then
                        a = 1
                        b = 4
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 13
            If matriz(1, 5) = 4 Then
                x = 1
                y = 5
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(1, 5) = 5 Then
                   x = 1
                   y = 5
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(1, 5) = 3 Then
                        a = 1
                        b = 5
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 14
            If matriz(1, 6) = 4 Then
                x = 1
                y = 6
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(1, 6) = 5 Then
                   x = 1
                   y = 6
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(1, 6) = 3 Then
                        a = 1
                        b = 6
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 15
            If matriz(1, 7) = 4 Then
                x = 1
                y = 7
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(1, 7) = 5 Then
                   x = 1
                   y = 7
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(1, 7) = 3 Then
                        a = 1
                        b = 7
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 16
            If matriz(2, 0) = 4 Then
                x = 2
                y = 0
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(2, 0) = 5 Then
                   x = 2
                   y = 0
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(2, 0) = 3 Then
                        a = 2
                        b = 0
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 17
            If matriz(2, 1) = 4 Then
                x = 2
                y = 1
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(2, 1) = 5 Then
                   x = 2
                   y = 1
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(2, 1) = 3 Then
                        a = 2
                        b = 1
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 18
            If matriz(2, 2) = 4 Then
                x = 2
                y = 2
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(2, 2) = 5 Then
                   x = 2
                   y = 2
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(2, 2) = 3 Then
                        a = 2
                        b = 2
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 19
            If matriz(2, 3) = 4 Then
                x = 2
                y = 3
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(2, 3) = 5 Then
                   x = 2
                   y = 3
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(2, 3) = 3 Then
                        a = 2
                        b = 3
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 20
            If matriz(2, 4) = 4 Then
                x = 2
                y = 4
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(2, 4) = 5 Then
                   x = 2
                   y = 4
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(2, 4) = 3 Then
                        a = 2
                        b = 4
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 21
            If matriz(2, 5) = 4 Then
                x = 2
                y = 5
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(2, 5) = 5 Then
                   x = 2
                   y = 5
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(2, 5) = 3 Then
                        a = 2
                        b = 5
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 22
            If matriz(2, 6) = 4 Then
                x = 2
                y = 6
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(2, 6) = 5 Then
                   x = 2
                   y = 6
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(2, 6) = 3 Then
                        a = 2
                        b = 6
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 23
            If matriz(2, 7) = 4 Then
                x = 2
                y = 7
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(2, 7) = 5 Then
                   x = 2
                   y = 7
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(2, 7) = 3 Then
                        a = 2
                        b = 7
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 24
            If matriz(3, 0) = 4 Then
                x = 3
                y = 0
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(3, 0) = 5 Then
                   x = 3
                   y = 0
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(3, 0) = 3 Then
                        a = 3
                        b = 0
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 25
            If matriz(3, 1) = 4 Then
                x = 3
                y = 1
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(3, 1) = 5 Then
                   x = 3
                   y = 1
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(3, 1) = 3 Then
                        a = 3
                        b = 1
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 26
            If matriz(3, 2) = 4 Then
                x = 3
                y = 2
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(3, 2) = 5 Then
                   x = 3
                   y = 2
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(3, 2) = 3 Then
                        a = 3
                        b = 2
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 27
            If matriz(3, 3) = 4 Then
                x = 3
                y = 3
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(3, 3) = 5 Then
                   x = 3
                   y = 3
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(3, 3) = 3 Then
                        a = 3
                        b = 3
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 28
            If matriz(3, 4) = 4 Then
                x = 3
                y = 4
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(3, 4) = 5 Then
                   x = 3
                   y = 4
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(3, 4) = 3 Then
                        a = 3
                        b = 4
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 29
            If matriz(3, 5) = 4 Then
                x = 3
                y = 5
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(3, 5) = 5 Then
                   x = 3
                   y = 5
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(3, 5) = 3 Then
                        a = 3
                        b = 5
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 30
            If matriz(3, 6) = 4 Then
                x = 3
                y = 6
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(3, 6) = 5 Then
                   x = 3
                   y = 6
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(3, 6) = 3 Then
                        a = 3
                        b = 6
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 31
            If matriz(3, 7) = 4 Then
                x = 3
                y = 7
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(3, 7) = 5 Then
                   x = 3
                   y = 7
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(3, 7) = 3 Then
                        a = 3
                        b = 7
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 32
            If matriz(4, 0) = 4 Then
                x = 4
                y = 0
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(4, 0) = 5 Then
                   x = 4
                   y = 0
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(4, 0) = 3 Then
                        a = 4
                        b = 0
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 33
            If matriz(4, 1) = 4 Then
                x = 4
                y = 1
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(4, 1) = 5 Then
                   x = 4
                   y = 1
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(4, 1) = 3 Then
                        a = 4
                        b = 1
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 34
            If matriz(4, 2) = 4 Then
                x = 4
                y = 2
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(4, 2) = 5 Then
                   x = 4
                   y = 2
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(4, 2) = 3 Then
                        a = 4
                        b = 2
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 35
            If matriz(4, 3) = 4 Then
                x = 4
                y = 3
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(4, 3) = 5 Then
                   x = 4
                   y = 3
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(4, 3) = 3 Then
                        a = 4
                        b = 3
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 36
            If matriz(4, 4) = 4 Then
                x = 4
                y = 4
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(4, 4) = 5 Then
                   x = 4
                   y = 4
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(4, 4) = 3 Then
                        a = 4
                        b = 4
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 37
            If matriz(4, 5) = 4 Then
                x = 4
                y = 5
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(4, 5) = 5 Then
                   x = 4
                   y = 5
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(4, 5) = 3 Then
                        a = 4
                        b = 5
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 38
            If matriz(4, 6) = 4 Then
                x = 4
                y = 6
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(4, 6) = 5 Then
                   x = 4
                   y = 6
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(4, 6) = 3 Then
                        a = 4
                        b = 6
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 39
            If matriz(4, 7) = 4 Then
                x = 4
                y = 7
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(4, 7) = 5 Then
                   x = 4
                   y = 7
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(4, 7) = 3 Then
                        a = 4
                        b = 7
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 40
            If matriz(5, 0) = 4 Then
                x = 5
                y = 0
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(5, 0) = 5 Then
                   x = 5
                   y = 0
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(5, 0) = 3 Then
                        a = 5
                        b = 0
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 41
            If matriz(5, 1) = 4 Then
                x = 5
                y = 1
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(5, 1) = 5 Then
                   x = 5
                   y = 1
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(5, 1) = 3 Then
                        a = 5
                        b = 1
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 42
            If matriz(5, 2) = 4 Then
                x = 5
                y = 2
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(5, 2) = 5 Then
                   x = 5
                   y = 2
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(5, 2) = 3 Then
                        a = 5
                        b = 2
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 43
            If matriz(5, 3) = 4 Then
                x = 5
                y = 3
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(5, 3) = 5 Then
                   x = 5
                   y = 3
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(5, 3) = 3 Then
                        a = 5
                        b = 3
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 44
            If matriz(5, 4) = 4 Then
                x = 5
                y = 4
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(5, 4) = 5 Then
                   x = 5
                   y = 4
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(5, 4) = 3 Then
                        a = 5
                        b = 4
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 45
            If matriz(5, 5) = 4 Then
                x = 5
                y = 5
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(5, 5) = 5 Then
                   x = 5
                   y = 5
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(5, 5) = 3 Then
                        a = 5
                        b = 5
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 46
            If matriz(5, 6) = 4 Then
                x = 5
                y = 6
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(5, 6) = 5 Then
                   x = 5
                   y = 6
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(5, 6) = 3 Then
                        a = 5
                        b = 6
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 47
            If matriz(5, 7) = 4 Then
                x = 5
                y = 7
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(5, 7) = 5 Then
                   x = 5
                   y = 7
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(5, 7) = 3 Then
                        a = 5
                        b = 7
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 48
            If matriz(6, 0) = 4 Then
                x = 6
                y = 0
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(6, 0) = 5 Then
                   x = 6
                   y = 0
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(6, 0) = 3 Then
                        a = 6
                        b = 0
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 49
            If matriz(6, 1) = 4 Then
                x = 6
                y = 1
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(6, 1) = 5 Then
                   x = 6
                   y = 1
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(6, 1) = 3 Then
                        a = 6
                        b = 1
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 50
            If matriz(6, 2) = 4 Then
                x = 6
                y = 2
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(6, 2) = 5 Then
                   x = 6
                   y = 2
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(6, 2) = 3 Then
                        a = 6
                        b = 2
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 51
            If matriz(6, 3) = 4 Then
                x = 6
                y = 3
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(6, 3) = 5 Then
                   x = 6
                   y = 3
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(6, 3) = 3 Then
                        a = 6
                        b = 3
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 52
            If matriz(6, 4) = 4 Then
                x = 6
                y = 4
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(6, 4) = 5 Then
                   x = 6
                   y = 4
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(6, 4) = 3 Then
                        a = 6
                        b = 4
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 53
            If matriz(6, 5) = 4 Then
                x = 6
                y = 5
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(6, 5) = 5 Then
                   x = 6
                   y = 5
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(6, 5) = 3 Then
                        a = 6
                        b = 5
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 54
            If matriz(6, 6) = 4 Then
                x = 6
                y = 6
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(6, 6) = 5 Then
                   x = 6
                   y = 6
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(6, 6) = 3 Then
                        a = 6
                        b = 6
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 55
            If matriz(6, 7) = 4 Then
                x = 6
                y = 7
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(6, 7) = 5 Then
                   x = 6
                   y = 7
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(6, 7) = 3 Then
                        a = 6
                        b = 7
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 56
            If matriz(7, 0) = 4 Then
                x = 7
                y = 0
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(7, 0) = 5 Then
                   x = 7
                   y = 0
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(7, 0) = 3 Then
                        a = 7
                        b = 0
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 57
            If matriz(7, 1) = 4 Then
                x = 7
                y = 1
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(7, 1) = 5 Then
                   x = 7
                   y = 1
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(7, 1) = 3 Then
                        a = 7
                        b = 1
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 58
            If matriz(7, 2) = 4 Then
                x = 7
                y = 2
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(7, 2) = 5 Then
                   x = 7
                   y = 2
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(7, 2) = 3 Then
                        a = 7
                        b = 2
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 59
            If matriz(7, 3) = 4 Then
                x = 7
                y = 3
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(7, 3) = 5 Then
                   x = 7
                   y = 3
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(7, 3) = 3 Then
                        a = 7
                        b = 3
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 60
            If matriz(7, 4) = 4 Then
                x = 7
                y = 4
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(7, 4) = 5 Then
                   x = 7
                   y = 4
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(7, 4) = 3 Then
                        a = 7
                        b = 4
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 61
            If matriz(7, 5) = 4 Then
                x = 7
                y = 5
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(7, 5) = 5 Then
                   x = 7
                   y = 5
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(7, 5) = 3 Then
                        a = 7
                        b = 5
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 62
            If matriz(7, 6) = 4 Then
                x = 7
                y = 6
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(7, 6) = 5 Then
                   x = 7
                   y = 6
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(7, 6) = 3 Then
                        a = 7
                        b = 6
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
        Case 63
            If matriz(7, 7) = 4 Then
                x = 7
                y = 7
                moverde = 4
                l = z
                Label1.Caption = "seleccionaste una ficha blanca" & vbNewLine
            Else
                If matriz(7, 7) = 5 Then
                   x = 7
                   y = 7
                   moverde = 5
                   l = z
                   Label1.Caption = "seleccionaste una ficha negra" & vbNewLine
                Else
                    If matriz(7, 7) = 3 Then
                        a = 7
                        b = 7
                        movera = 3
                        Label1.Caption = "ya tines seleccionado un espacio para mover" & vbNewLine
                    End If
                End If
            End If
    End Select
End Sub

Private Sub Command1_Click() ' para regresar al formulario del menu
    Form3.Hide
    Load Form1
    Form1.Show
End Sub

Private Sub Command2_Click()
    reinicio = 0 ' reinicia el tablero
End Sub


Private Sub Command3_Click() 'este boton es para completar la jugada seleccionada
   If matriz(a, b) = 3 And moverfi = 4 Then 'comprueba que la posicion a mover este vacia y que la ficha selecionada sea blanca
        If turno = 4 And matriz(x, y) = 4 Then 'comprueba que sea el turno de las blancas y la posicion de una fhicha blanca
            celda(z).Picture = LoadPicture(App.Path & "\ficha blanca.jpg")
            matriz(a, b) = 4 ' el comando de arriba carga las imagen en la nueva pocicion y se cambia el valor de la posicion a 4 para indicar que es blanca
        End If
    Else
        If matriz(a, b) = 3 And moverfi = 5 Then 'comprueba que la posicion a mover este vacia y que la ficha selecionada sea negra
            If turno = 5 And matriz(x, y) = 5 Then 'comprueba que sea el turno de las negras y la posicion de una fhicha blanca
                celda(z).Picture = LoadPicture(App.Path & "\ficha negra.jpg")
                matriz(a, b) = 5 ' el comando de arriba carga las imagen en la nueva pocicion y se cambia el valor de la posicion a 5 para indicar que es blanca
            End If
        End If
    End If

    Select Case matriz(x, y)
        Case 4 'si es 4 comprueba el turno y que la ficha sea 4 o 5
            If turno = 4 Then
                If moverfi = 4 Then 'en este caso es para mover la ficha
                    celda(l).Picture = LoadPicture("") ' este comando borra la imagen el la posicion selccionada
                    moverfi = 0
                    matriz(x, y) = 3
                    turno = 5
                End If
                If comerfi = 5 Then ' en este caso es para comer la ficha
                   celda(c).Picture = LoadPicture("")
                    comerfi = 0
                    matriz(x, y) = 3
                    matriz(ca, cb) = 3
                    turno = 4
                End If
            End If
        Case 5 'si es 5 comprueba el turno y que la ficha sea 4 o 5
            If turno = 5 Then
                If moverfi = 5 Then
                   celda(l).Picture = LoadPicture("")
                    moverfi = 0
                    matriz(x, y) = 3
                    turno = 4
                End If
                If comerfi = 4 Then
                   celda(c).Picture = LoadPicture("")
                    comerfi = 0
                    matriz(x, y) = 3
                    matriz(ca, cb) = 3
                    turno = 5
                End If
            End If
    End Select
End Sub







Private Sub Timer1_Timer() ' se usa un timer para que compruebe el estado y ejecute las siguientes intrucciones
    Dim i, j, k, m As Integer
    
    If reinicio = 0 Then ' si esta en cero se reinicia el tablero mediante un for
        k = 0
        For i = 0 To 7 ' se recorre el for y se establecen los valores 5 para las fichas negras 4 para las blancas 3 para los espacios vacios y 0 para las posiciones que no se usan
            For j = 0 To 7
                m = k Mod 2
                If (k >= 0 And k < 8) Or (k > 15 And k < 24) Then
                    If m = 0 Then
                        matriz(i, j) = 5
                        celda(k).Picture = LoadPicture(App.Path & "\ficha negra.jpg")
                        k = k + 1
                    Else
                        matriz(i, j) = 0
                        celda(k).Picture = LoadPicture("")
                        k = k + 1
                    End If
                Else
                    If k > 7 And k < 16 Then
                        If m = 0 Then
                           matriz(i, j) = 0
                           celda(k).Picture = LoadPicture("")
                           k = k + 1
                        Else
                          matriz(i, j) = 5
                          celda(k).Picture = LoadPicture(App.Path & "\ficha negra.jpg")
                          k = k + 1
                        End If
                    Else
                        If k > 23 And k < 32 Then
                            If m = 0 Then
                                matriz(i, j) = 0
                                celda(k).Picture = LoadPicture("")
                                k = k + 1
                            Else
                                matriz(i, j) = 3
                                celda(k).Picture = LoadPicture("")
                                k = k + 1
                            End If
                        Else
                            If k > 31 And k < 40 Then
                                If m = 0 Then
                                    matriz(i, j) = 3
                                    celda(k).Picture = LoadPicture("")
                                    k = k + 1
                                Else
                                    matriz(i, j) = 0
                                    celda(k).Picture = LoadPicture("")
                                    k = k + 1
                                End If
                            Else
                                If k > 47 And k < 56 Then
                                    If m = 0 Then
                                        matriz(i, j) = 4
                                        celda(k).Picture = LoadPicture(App.Path & "\ficha blanca.jpg")
                                        k = k + 1
                                    Else
                                        matriz(i, j) = 0
                                        celda(k).Picture = LoadPicture("")
                                        k = k + 1
                                    End If
                                Else
                                    If m = 0 Then
                                        matriz(i, j) = 0
                                        celda(k).Picture = LoadPicture("")
                                        k = k + 1
                                    Else
                                        matriz(i, j) = 4
                                        celda(k).Picture = LoadPicture(App.Path & "\ficha blanca.jpg")
                                        k = k + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next j
        Next i
        turno = 4
        reinicio = 1
        moverfi = 0
        movera = 0
        Label1.Caption = "comienzan jugando las blancas" & vbNewLine
                
    End If

    If turno = 4 And a < x Then ' si turno es igual a 4 le toca a las blancas
        If moverde = 4 And movera = 3 Then 'si el matriz seleccionado es una ficha blanca y la otra matriz esta vacia se ejecuata
            Select Case y
                Case 0
                    If matriz(x - 1, y + 1) = 3 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                    End If
                    If matriz(x - 2, y + 2) = 3 And matriz(x - 1, y + 1) = 5 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                        comerfi = 5 'modifica la variable a 5 que es para comer una ficha negra
                        ca = x - 1
                        cb = y + 1
                        c = l - 7
                    End If
                Case 1
                    If matriz(x - 1, y + 1) = 3 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                    End If
                    If matriz(x - 2, y + 2) = 3 And matriz(x - 1, y + 1) = 5 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                        comerfi = 5 'modifica la variable a 5 que es para comer una ficha negra
                        ca = x - 1
                        cb = y + 1
                        c = l - 7
                    End If
                    If matriz(x - 1, y - 1) = 3 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                    End If
                Case 6
                    If matriz(x - 1, y + 1) = 3 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                    End If
                    If matriz(x - 1, y - 1) = 3 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                    End If
                    If matriz(x - 2, y - 2) = 3 And matriz(x - 1, y - 1) = 5 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                        comerfi = 5 'modifica la variable a 5 que es para comer una ficha negra
                        ca = x - 1
                        cb = y - 1
                        c = l - 9
                    End If
                Case 7
                    If matriz(x - 1, y - 1) = 3 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                    End If
                    If matriz(x - 2, y - 2) = 3 And matriz(x - 1, y - 1) = 5 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                        comerfi = 5 'modifica la variable a 5 que es para comer una ficha negra
                        ca = x - 1
                        cb = y - 1
                        c = l - 9
                    End If
                Case Else
                    If matriz(x - 1, y + 1) = 3 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                    End If
                    If matriz(x - 2, y + 2) = 3 And matriz(x - 1, y + 1) = 5 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                        comerfi = 5 'modifica la variable a 5 que es para comer una ficha negra
                        ca = x - 1
                        cb = y + 1
                        c = l - 7
                    End If
                    If matriz(x - 1, y - 1) = 3 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                    End If
                    If matriz(x - 2, y - 2) = 3 And matriz(x - 1, y - 1) = 5 Then
                        moverfi = 4 'modifica la variable a 4 que es para mover la ficha blanca
                        comerfi = 5 'modifica la variable a 5 que es para comer una ficha negra
                        ca = x - 1
                        cb = y - 1
                        c = l - 9
                    End If
            End Select
        End If
    End If
    
    If turno = 5 And a > x Then ' si turno es igual a 5 le toca las fichas negras
        If moverde = 5 And movera = 3 Then 'si el matriz seleccionado es una ficha negra y la otra matriz esta vacia se ejecuata
            Select Case y
                Case 0
                    If matriz(x + 1, y + 1) = 3 Then
                        moverfi = 5
                    End If
                    If matriz(x + 2, y + 2) = 3 And matriz(x + 1, y + 1) = 4 Then
                        moverfi = 5
                        comerfi = 4
                        ca = x + 1
                        cb = y + 1
                        c = l + 9
                    End If
                Case 1
                    If matriz(x + 1, y + 1) = 3 Then
                        moverfi = 5
                    End If
                    If matriz(x + 2, y + 2) = 3 And matriz(x + 1, y + 1) = 4 Then
                        moverfi = 5
                        comerfi = 4
                        ca = x + 1
                        cb = y + 1
                        c = l + 9
                    End If
                    If matriz(x + 1, y - 1) = 3 Then
                        moverfi = 5
                    End If
                Case 6
                    If matriz(x + 1, y - 1) = 3 Then
                        moverfi = 5
                    End If
                    If matriz(x + 1, y - 1) = 3 Then
                        moverfi = 5
                    End If
                    If matriz(x + 2, y - 2) = 3 And matriz(x + 1, y - 1) = 4 Then
                        moverfi = 5
                        comerfi = 4
                        ca = x + 1
                        cb = y - 1
                        c = l + 7
                    End If
                Case 7
                    If matriz(x + 1, y - 1) = 3 Then
                        moverfi = 5
                    End If
                    If matriz(x + 2, y - 2) = 3 And matriz(x + 1, y - 1) = 4 Then
                        moverfi = 5
                        comerfi = 4
                        ca = x + 1
                        cb = y - 1
                        c = l + 7
                    End If
                Case Else
                    If matriz(x + 1, y + 1) = 3 Then
                        moverfi = 5
                    End If
                    If matriz(x + 2, y + 2) = 3 And matriz(x + 1, y + 1) = 4 Then
                        moverfi = 5
                        comerfi = 4
                        ca = x + 1
                        cb = y + 1
                        c = l + 9
                    End If
                    If matriz(x + 1, y - 1) = 3 Then
                        moverfi = 5
                    End If
                    If matriz(x + 2, y - 2) = 3 And matriz(x + 1, y - 1) = 4 Then
                        moverfi = 5
                        comerfi = 4
                        ca = x + 1
                        cb = y - 1
                        c = l + 7
                    End If
            End Select
        End If
    End If
End Sub

Private Sub Timer2_Timer() ' este timer actualiza el estado de los label donde te dice el turno
    If turno = 4 Then
        Label2.Caption = "turno para las blancas" & vbNewLine
    Else
        Label2.Caption = "turno para las negras" & vbNewLine
    End If
End Sub

Private Sub Timer3_Timer()

End Sub
