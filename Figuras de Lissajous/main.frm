VERSION 5.00
Begin VB.Form frmLisaj 
   Caption         =   "curvas de lissajous"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "dibujar p a p"
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "pi *"
      Height          =   255
      Left            =   6000
      TabIndex        =   16
      Top             =   3000
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "borrar pantalla"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   11
      Text            =   "0.5"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   3
      Left            =   6000
      TabIndex        =   9
      Text            =   "4"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   2
      Left            =   6000
      TabIndex        =   7
      Text            =   "4"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   5
      Text            =   "1"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   0
      Left            =   6000
      TabIndex        =   3
      Text            =   "1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "dibujar lissajous"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y = Y0 * cos ( w2 * t + fi )"
      Height          =   195
      Left            =   3480
      TabIndex        =   14
      Top             =   5880
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X = X0 * cos ( w1 * t )"
      Height          =   195
      Left            =   840
      TabIndex        =   13
      Top             =   5880
      Width           =   1515
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Dif. Fase fi"
      Height          =   195
      Index           =   4
      Left            =   6000
      TabIndex        =   12
      Top             =   2760
      Width           =   750
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Frecuencia w2"
      Height          =   195
      Index           =   3
      Left            =   6000
      TabIndex        =   10
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Frecuencia w1"
      Height          =   195
      Index           =   2
      Left            =   6000
      TabIndex        =   8
      Top             =   1560
      Width           =   1050
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Amplitud Y0"
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   6
      Top             =   960
      Width           =   840
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Amplitud X0"
      Height          =   195
      Index           =   0
      Left            =   6000
      TabIndex        =   4
      Top             =   360
      Width           =   840
   End
End
Attribute VB_Name = "frmLisaj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const pi = 3.141516
Dim t, X, Y, z, A, B, w1, w2, fi As Double


Dim xant, yant, x0, y0 As Double
Dim NOprimerGraficar As Boolean


Private Sub Command1_Click()
A = txtParam(0) '1
B = txtParam(1) '1
w1 = txtParam(2) '4
w2 = txtParam(3) '8
fi = IIf(Check1.Value = 0, 1, pi) * Val(txtParam(4)) ' pi / 2
'x=A * sin(alfa * t + gama) y=B * sin (beta * t)
't = t + 0.1
For t = 0 To 10 * IIf(w1 <= w2, 2.1 * pi / w1, 2.1 * pi / w2) Step 0.1

' x = A * Exp(-t) * Sin(alfa * t + gama) * Picture1.ScaleWidth / 4
' y = B * Exp(-t) * Sin(beta * t) * Picture1.ScaleHeight / 4
 
 X = A * Cos(w1 * t) * Picture1.ScaleWidth / 4
 Y = B * Cos(w2 * t + fi) * Picture1.ScaleHeight / 4
Picture1.DrawWidth = 4
Picture1.PSet (X + Picture1.ScaleWidth / 2, -Y + Picture1.ScaleHeight / 2)
Picture1.DrawWidth = 1
graficar X + Picture1.ScaleWidth / 2, -Y + Picture1.ScaleHeight / 2, Picture1, RGB(255, 0, 0)

Next t

NOprimerGraficar = False
End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Command3_Click()
Picture1.Cls

End Sub

Private Sub Command4_Click()
A = txtParam(0) '1
B = txtParam(1) '1
w1 = txtParam(2) '4
w2 = txtParam(3) '8
fi = IIf(Check1.Value = 0, 1, pi) * Val(txtParam(4)) ' pi / 2
'x=A * sin(alfa * t + gama) y=B * sin (beta * t)
t = t + 0.1

' x = A * Exp(-t) * Sin(alfa * t + gama) * Picture1.ScaleWidth / 4
' y = B * Exp(-t) * Sin(beta * t) * Picture1.ScaleHeight / 4
 
 X = A * Cos(w1 * t) * Picture1.ScaleWidth / 4
 Y = B * Cos(w2 * t + fi) * Picture1.ScaleHeight / 4
Picture1.DrawWidth = 4
Picture1.PSet (X + Picture1.ScaleWidth / 2, -Y + Picture1.ScaleHeight / 2)
Picture1.DrawWidth = 1
graficar X + Picture1.ScaleWidth / 2, -Y + Picture1.ScaleHeight / 2, Picture1, RGB(255, 0, 0)



End Sub

Private Sub Form_Load()
    Me.Show
    Picture1.ScaleMode = vbPixels
End Sub

Sub graficar(Xp As Variant, Yp As Variant, pic As PictureBox, coloRin As Long)
If NOprimerGraficar Then
    pic.Line (xant, yant)-(Xp, Yp), coloRin    'RGB(255, 0, 0)
    
    xant = Xp
    yant = Yp

Else
    
    xant = Xp
    yant = Yp
    
    pic.Line (xant, yant)-(Xp, Yp), coloRin    'RGB(255, 0, 0)
    NOprimerGraficar = True
End If
End Sub

