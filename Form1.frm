VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "empezar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8610
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3780
      Width           =   1590
   End
   Begin VB.Label Lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calcular el factorial de un nùmero aleatorio."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5250
      TabIndex        =   1
      Top             =   1890
      Width           =   7680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numeros_array(9), a, numeros, resultado As Integer


Private Sub Command1_Click()
    numerosAleatorios
    calcularFactorial
    mensaje
End Sub
Private Function numerosAleatorios() As Integer

Randomize
For a = 0 To 9
numeros_array(a) = Int(Rnd * 10) + 1
numerosAleatorios = numeros_array(a)
Lbl1 = numeros_array(a)
Next a


End Function

Private Function calcularFactorial() As Integer
Dim resultado, b As Integer

resultado = 1

For b = 1 To numerosAleatorios
 resultado = resultado * b
Next b


Print resultado

End Function
Private Function mensaje()
MsgBox "El factorial de" & numerosAleatorios & " es " & resultado, vbInformation, "Resultado"
End Function
