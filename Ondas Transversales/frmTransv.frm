VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransv 
   Caption         =   "Ondas Transversales en Una Cuerda Tensa"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "<-stepback"
      Height          =   495
      Left            =   2640
      TabIndex        =   32
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "detener"
      Height          =   495
      Left            =   600
      TabIndex        =   21
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "step->"
      Height          =   495
      Left            =   2640
      TabIndex        =   20
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "continuar |>"
      Height          =   495
      Left            =   4440
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   4695
      Index           =   0
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   8281
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      SelStart        =   500
      TickStyle       =   2
      Value           =   500
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9360
      Top             =   2880
   End
   Begin VB.CommandButton cmdEqui 
      Caption         =   "Equilibrio"
      Height          =   495
      Left            =   -1560
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.PictureBox pTransv 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   4695
      Index           =   1
      Left            =   6480
      TabIndex        =   3
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   8281
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   4695
      Index           =   2
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   8281
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   4695
      Index           =   3
      Left            =   7440
      TabIndex        =   5
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   8281
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   4695
      Index           =   4
      Left            =   7920
      TabIndex        =   6
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   8281
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   4695
      Index           =   5
      Left            =   8400
      TabIndex        =   7
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   8281
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   4695
      Index           =   6
      Left            =   8880
      TabIndex        =   8
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   8281
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   9375
      Index           =   7
      Left            =   15360
      TabIndex        =   9
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   16536
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   9375
      Index           =   8
      Left            =   15840
      TabIndex        =   22
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   16536
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   9375
      Index           =   9
      Left            =   16320
      TabIndex        =   23
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   16536
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   9375
      Index           =   10
      Left            =   16800
      TabIndex        =   24
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   16536
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   9375
      Index           =   11
      Left            =   17280
      TabIndex        =   25
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   16536
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin MSComctlLib.Slider sldModos 
      Height          =   9375
      Index           =   12
      Left            =   17760
      TabIndex        =   26
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   16536
      _Version        =   393216
      Orientation     =   1
      Max             =   1000
      TickStyle       =   2
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "9"
      Height          =   195
      Index           =   12
      Left            =   16080
      TabIndex        =   31
      Top             =   9600
      Width           =   90
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "10"
      Height          =   195
      Index           =   11
      Left            =   16560
      TabIndex        =   30
      Top             =   9600
      Width           =   180
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "11"
      Height          =   195
      Index           =   10
      Left            =   17040
      TabIndex        =   29
      Top             =   9600
      Width           =   180
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "12"
      Height          =   195
      Index           =   9
      Left            =   17520
      TabIndex        =   28
      Top             =   9600
      Width           =   180
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "13"
      Height          =   195
      Index           =   8
      Left            =   18000
      TabIndex        =   27
      Top             =   9600
      Width           =   180
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "8"
      Height          =   195
      Index           =   7
      Left            =   15600
      TabIndex        =   18
      Top             =   9600
      Width           =   90
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "7"
      Height          =   195
      Index           =   6
      Left            =   9120
      TabIndex        =   17
      Top             =   4800
      Width           =   90
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "6"
      Height          =   195
      Index           =   5
      Left            =   8640
      TabIndex        =   16
      Top             =   4800
      Width           =   90
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "5"
      Height          =   195
      Index           =   4
      Left            =   8160
      TabIndex        =   15
      Top             =   4800
      Width           =   90
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   195
      Index           =   3
      Left            =   7680
      TabIndex        =   14
      Top             =   4800
      Width           =   90
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "3"
      Height          =   195
      Index           =   2
      Left            =   7200
      TabIndex        =   13
      Top             =   4800
      Width           =   90
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   195
      Index           =   1
      Left            =   6720
      TabIndex        =   12
      Top             =   4800
      Width           =   90
   End
   Begin VB.Label lblModo 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Index           =   0
      Left            =   6240
      TabIndex        =   11
      Top             =   4800
      Width           =   90
   End
   Begin VB.Label lblModos 
      AutoSize        =   -1  'True
      Caption         =   "Modo:"
      Height          =   195
      Left            =   5520
      TabIndex        =   10
      Top             =   4800
      Width           =   450
   End
End
Attribute VB_Name = "frmTransv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim dX As Double
Dim x As Double
Dim pi As Double
Dim k As Double
Dim c As Double
Dim w As Double
Dim t As Double

Dim Amps(12) As Double

Const dt = 0.1

Dim espP, eW, eH As Integer 'Ancho y altura efectivas de la pictureBox


Const Np = 800
Const Lp = 100

Const Alt = 100
'Const m = 3
'Const lambda = 4 * Lp / (2 * m - 1)
Const lambda = 2 * Lp

Const T0 = 100
Const Ro0 = 0.1

Const Amp = Lp / 300

Function f(xp As Double) As Double
k = 2 * pi / lambda
c = Sqr(T0 / Ro0)
w = k * c
'f = Amp / 2 * Sin(k * x - w * t)
'f = f + Amp / 2.5 * Sin(k * x + w * t)
'f = Amp * Sin(k * xp) * Cos(w * t)
f = (-1) * (Amps(0) * Sin(k * xp) * Cos(w * t) + Amps(1) * Sin(k * 2 * xp) * Cos(w * 2 * t) + Amps(2) * Sin(k * 3 * xp) * Cos(w * 3 * t) + _
Amps(3) * Sin(k * 4 * xp) * Cos(w * 4 * t) + Amps(4) * Sin(k * 5 * xp) * Cos(w * 5 * t) + Amps(5) * Sin(k * 6 * xp) * Cos(w * 6 * t) + _
Amps(6) * Sin(k * 7 * xp) * Cos(w * 7 * t) + Amps(7) * Sin(k * 8 * xp) * Cos(w * 8 * t) + Amps(8) * Sin(k * 9 * xp) * Cos(w * 9 * t) + _
Amps(9) * Sin(k * 10 * xp) * Cos(w * 10 * t) + Amps(11) * Sin(k * 12 * xp) * Cos(w * 12 * t) + Amps(12) * Sin(k * 13 * xp) * Cos(w * 13 * t))

'f = Amp * Sin(k * xp) * Cos(w * t) + Amp * Sin(3 * k * xp) * Cos(3 * w * t) + Amp * Sin(5 * k * xp) * Cos(5 * w * t)
'f = Amp * Sin(k * xp - w * t) + Amp * Sin(k * xp + w * t)
'f = Amp * Sin(k * xp - w * t) + Amp * Sin(3 * k * xp - 3 * w * t) ' + Sin(5 * k * xp - 5 * w * t)
End Function
Sub MostrarPs()
    pTransv.Cls
    Reglita
    For i = 0 To Np
        x = i * dX
        'pTransv.Line (espP + (f(X) + X / Lp) * eW, 10)-(espP + (f(X) + X / Lp) * eW, pTransv.Height - 10), &HFF
        pTransv.PSet (espP / 2 + (x / Lp) * eW, (1 / 2 - f(x)) * eH), &HFF
    
    Next i
End Sub


Sub Reglita()

    pTransv.Line (0, pTransv.Height / 2)-(pTransv.Width, pTransv.Height / 2)
End Sub

Private Sub cmdEqui_Click()

t = 3 * lambda / 3 * c - dt
Timer1_Timer

End Sub

Private Sub Command1_Click()
Timer1.Enabled = False
End Sub

Private Sub Command2_Click()
Timer1_Timer
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True
End Sub

Private Sub Command4_Click()
t = t - 2 * dt
Timer1_Timer
End Sub

Private Sub Form_Load()

dX = Lp / Np
pi = Atn(1) * 4
pTransv.ScaleMode = vbPixels
Me.ScaleMode = vbPixels
espP = 10
eW = pTransv.Width - 2 * espP
eH = pTransv.Height
pTransv.DrawWidth = 2
For i = 0 To sldModos.Count - 1: sldModos_Click (i): Next i

End Sub

Private Sub sldModos_Click(Index As Integer)
Amps(Index) = Amp * (sldModos(Index).Value - sldModos(Index).Min) / sldModos(Index).Max
End Sub

Private Sub sldModos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Amps(Index) = Amp * (sldModos(Index).Value - sldModos(Index).Min) / sldModos(Index).Max

End Sub

Private Sub sldModos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Amps(Index) = Amp * (sldModos(Index).Value - sldModos(Index).Min) / sldModos(Index).Max

End Sub

Private Sub Timer1_Timer()
t = t + dt
MostrarPs
End Sub
