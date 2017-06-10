VERSION 5.00
Begin VB.Form frmViscoso 
   Caption         =   "Partículas en suspensión en un líquido viscoso"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "frmViscoso.frx":0000
   LinkTopic       =   "frmViscoso"
   ScaleHeight     =   5625
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Exportar a excel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   18
      Top             =   4800
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   2880
      Picture         =   "frmViscoso.frx":0442
      ScaleHeight     =   975
      ScaleWidth      =   4815
      TabIndex        =   17
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Width           =   615
   End
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   3000
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   1
      Top             =   1200
      Width           =   5600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dejar Caer"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   3480
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "Ecuación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   1440
      TabIndex        =   16
      Top             =   360
      Width           =   1200
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "dt (s) ="
      Height          =   195
      Index           =   5
      Left            =   1440
      TabIndex        =   15
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "g (m/s^2) ="
      Height          =   195
      Index           =   4
      Left            =   1080
      TabIndex        =   13
      Top             =   2640
      Width           =   810
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "nu (kg/m·s) ="
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   11
      Top             =   2280
      Width           =   945
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "M (g) ="
      Height          =   195
      Index           =   2
      Left            =   1440
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "ro (g/cm^3) ="
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "R (mm) ="
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Velocidad"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   6000
      TabIndex        =   3
      Top             =   5280
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Posición"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3960
      TabIndex        =   2
      Top             =   5280
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000000&
      Height          =   375
      Left            =   360
      Shape           =   3  'Circle
      Top             =   720
      Width           =   375
   End
   Begin VB.Line Line3 
      X1              =   840
      X2              =   840
      Y1              =   480
      Y2              =   4560
   End
   Begin VB.Line Line2 
      X1              =   840
      X2              =   240
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   240
      Y1              =   480
      Y2              =   4560
   End
End
Attribute VB_Name = "frmViscoso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim t As Double
Dim X As Double
Dim DXDT As Double
Dim D2XDT2 As Double
Dim M As Double
Dim pi As Double
Dim G As Double 'Acel. Gravedad
Dim dt As Double 'Tiempo de integracion
Dim R  As Double 'Radio
Dim ro As Double 'Densidad
Dim nu As Double 'Viscosidad



Private Sub Command1_Click()

R = txtData(0).Text / 1000
ro = txtData(1).Text / 1000
M = txtData(2).Text / 1000
nu = txtData(3).Text
G = txtData(4).Text
dt = txtData(5).Text

p1.Cls
X = 0
t = 0
DXDT = 0
D2XDT2 = 0

p1.Line (0, p1.Height / 2)-(p1.Width, p1.Height / 2)

Timer1.Enabled = True



End Sub

Private Sub Command2_Click()
 ' TODO: AGREGAR EXPORTACION A EXCEL
End Sub

Private Sub Form_Load()

'R
'ro
'M
'nu
'G
'dt


txtData(0).Text = (3.65 / 2)
txtData(1).Text = 0.8 * (10 ^ 3)
txtData(2).Text = 0.162
txtData(3).Text = 0.2
txtData(4).Text = 9.8
txtData(5).Text = 0.01

Me.ScaleMode = vbPixels
p1.ScaleMode = vbPixels

X = 0
DXDT = 0
D2XDT2 = 0



p1.Line (0, p1.Height / 2)-(p1.Width, p1.Height / 2)
pi = Atn(1) * 4

 
R = 3.65 / 2000
ro = (0.8 * (100 ^ 3)) / 1000
M = 0.162 / 1000
nu = 0.2
G = 9.8
dt = 0.01




End Sub

Private Sub Timer1_Timer()
t = t + dt

'D2XDT2 = G - (ROZ * DXDT) / M - (ro * G * ((4 * pi * R ^ 3) / 3)) / M
D2XDT2 = G - 6 * pi * R * nu * DXDT / M - ((ro * 4 * pi * G * (R ^ 3)) / 3) / M

DXDT = DXDT + D2XDT2 * dt


X = X + DXDT * dt

Shape1.Top = X * 300 + 50

p1.PSet (t * 80, -X * 100 + p1.Height / 2), RGB(255, 0, 0)
p1.PSet (t * 80, -DXDT * 100 + p1.Height / 2), RGB(0, 255, 0)
'p1.PSet (t*30, D2XDT2 / 1.5 + p1.Height / 2), RGB(0, 0, 255)

If Shape1.Top + Shape1.Height >= Line2.Y2 Or Shape1.Top <= Line1.X1 Then
 Timer1.Enabled = False

End If
End Sub
