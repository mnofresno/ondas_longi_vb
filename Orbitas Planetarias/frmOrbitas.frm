VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Simulación de órbitas planetarias"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   4095
   ClientWidth     =   11670
   Icon            =   "frmOrbitas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Detener"
      Height          =   495
      Left            =   8880
      TabIndex        =   37
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CheckBox ckGrafi 
      Caption         =   "Todo Junto"
      Height          =   255
      Left            =   10080
      TabIndex        =   35
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   10200
      TabIndex        =   30
      Top             =   4680
      Width           =   1215
   End
   Begin VB.OptionButton chkFza 
      Height          =   195
      Index           =   1
      Left            =   4800
      TabIndex        =   27
      Top             =   4800
      Width           =   195
   End
   Begin VB.OptionButton chkFza 
      Height          =   195
      Index           =   0
      Left            =   4800
      TabIndex        =   26
      Top             =   4560
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   4440
      ScaleHeight     =   4155
      ScaleWidth      =   6915
      TabIndex        =   2
      Top             =   120
      Width           =   6975
      Begin VB.TextBox inicial 
         Height          =   285
         Index           =   7
         Left            =   1680
         TabIndex        =   33
         Text            =   "0"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox inicial 
         Height          =   285
         Index           =   6
         Left            =   1680
         TabIndex        =   31
         Text            =   "4"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.PictureBox picGRadio 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   3000
         ScaleHeight     =   1395
         ScaleWidth      =   3795
         TabIndex        =   29
         Top             =   1920
         Width           =   3855
      End
      Begin VB.PictureBox picGtita 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   3000
         ScaleHeight     =   1395
         ScaleWidth      =   3795
         TabIndex        =   28
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox inicial 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox inicial 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Text            =   "1,92"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox inicial 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   20
         Text            =   "0"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox inicial 
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   19
         Text            =   "30"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox inicial 
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   18
         Text            =   "0"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox inicial 
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   17
         Text            =   "0"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "L0 ="
         Height          =   195
         Index           =   9
         Left            =   1320
         TabIndex        =   34
         Top             =   3720
         Width           =   315
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "K ="
         Height          =   195
         Index           =   8
         Left            =   1320
         TabIndex        =   32
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "P. Iniciales"
         Height          =   195
         Index           =   7
         Left            =   1920
         TabIndex        =   23
         Top             =   240
         Width           =   765
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "d2radio="
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "dradio="
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   525
      End
      Begin VB.Label valor 
         AutoSize        =   -1  'True
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   14
         Top             =   3000
         Width           =   90
      End
      Begin VB.Label valor 
         AutoSize        =   -1  'True
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   13
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "radio="
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "d2tita="
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "dtita="
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "tita="
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   300
      End
      Begin VB.Label etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "tiempo="
         Height          =   195
         Index           =   0
         Left            =   3000
         TabIndex        =   8
         Top             =   3480
         Width           =   555
      End
      Begin VB.Label valor 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   7
         Top             =   2280
         Width           =   90
      End
      Begin VB.Label valor 
         AutoSize        =   -1  'True
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   6
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label valor 
         AutoSize        =   -1  'True
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   5
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label valor 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   90
      End
      Begin VB.Label valor 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   3
         Top             =   3480
         Width           =   90
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picUniv 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.Shape mchica 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Left            =   360
         Shape           =   3  'Circle
         Top             =   480
         Width           =   135
      End
      Begin VB.Shape Mgrande 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   960
         Shape           =   3  'Circle
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Label col 
      AutoSize        =   -1  'True
      Caption         =   "COLISION"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1680
      TabIndex        =   36
      Top             =   4440
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "F = -K · (r - L0)"
      Height          =   195
      Left            =   5040
      TabIndex        =   25
      Top             =   4800
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F = -G·M·m/r^2"
      Height          =   195
      Left            =   5040
      TabIndex        =   24
      Top             =   4560
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'constantes del problema

Dim Mg, mc, g, Gr, l, lc, Pi, kConst, lNat As Double
Dim Rtierra, Rsol, RLuna, DistSol, DistLuna, MTierra, Msol, Mluna As Double

'Parámetros iniciales

Dim tita0, radio0 As Double
Dim dtita0, dradio0 As Double
Dim d2tita0, d2radio0 As Double


'Variables con dependencia temporal

Dim t, fuerzaR, fuerzaT, campo, tita, radio As Double
Dim dtita, dradio, dt As Double
Dim d2tita, d2radio As Double
Dim X, Y, x1, x2, y1, y2 As Double


Dim deTener As Boolean

Private Sub Command1_Click()
deTener = False
t = 0
picUniv.Cls
picGRadio.Cls
picGtita.Cls
'picGtita.Print "Valores Tangenciales:"
'picGRadio.Print "Valores Radiales:"


        picGRadio.Line (0, picGRadio.ScaleHeight / 4)-(picGRadio.ScaleWidth, picGRadio.ScaleHeight / 4), RGB(240, 240, 240)
        picGRadio.Line (0, 2 * picGRadio.ScaleHeight / 4)-(picGRadio.ScaleWidth, 2 * picGRadio.ScaleHeight / 4), RGB(240, 240, 240)
        picGRadio.Line (0, 3 * picGRadio.ScaleHeight / 4)-(picGRadio.ScaleWidth, 3 * picGRadio.ScaleHeight / 4), RGB(240, 240, 240)

        picGtita.Line (0, picGtita.ScaleHeight / 4)-(picGRadio.ScaleWidth, picGtita.ScaleHeight / 4), RGB(240, 240, 240)
        picGtita.Line (0, 2 * picGtita.ScaleHeight / 4)-(picGRadio.ScaleWidth, 2 * picGtita.ScaleHeight / 4), RGB(240, 240, 240)
        picGtita.Line (0, 3 * picGtita.ScaleHeight / 4)-(picGRadio.ScaleWidth, 3 * picGtita.ScaleHeight / 4), RGB(240, 240, 240)


    tita = inicial(0) 'tita0
    dtita = inicial(1) 'dtita0
    d2tita = inicial(2) 'd2tita0
    radio = inicial(3) 'radio0
    dradio = inicial(4) 'dradio0
    d2radio = inicial(5) 'd2radio0
    kConst = inicial(6)
    lNat = inicial(7)

col.visible = False
'Err.Visible = False

On Error GoTo err

    Do
      
    DoEvents
    'Mostrar valores
    valor(0) = Format(t, "00.000")
    valor(1) = Format(tita, "00.000")
    valor(2) = Format(dtita, "00.000")
    valor(3) = Format(d2tita, "00.000")
    valor(4) = Format(radio, "00.000")
    valor(5) = Format(dradio, "00.000")
    valor(6) = Format(d2radio, "00.000")

    
    
    
    
    fuerzaR = IIf(chkFza(0).Value, -Gr * Mg * mc / (radio) ^ 2, -kConst * (radio - lNat))  ' -Gr * mc * Mg / radio  '
    d2radio = (fuerzaR / mc) + radio * dtita ^ 2
    'Debug.Print fuerzaR
    
    dradio = dradio + d2radio * dt
    radio = Abs(radio + dradio * dt)
    
    fuerzaT = 0 ' 100 / radio ^ 2
    d2tita = (fuerzaT / (mc * radio)) - 2 * dradio * dtita / radio
    
    dtita = dtita + d2tita * dt
    tita = tita + dtita * dt
    
    
    X = radio * Cos(tita) + picUniv.ScaleWidth / 2
    Y = -radio * Sin(tita) + picUniv.ScaleHeight / 2
    
    picUniv.PSet (X, Y), RGB(255, 0, 0)

If ckGrafi.Value Then
        picGtita.PSet (t * 2, -Atn(Tan(tita)) * 3 + picGtita.ScaleHeight / 2)
        picGtita.PSet (t * 2, -dtita * 4 + picGtita.ScaleHeight / 2), RGB(255, 0, 0)
        picGtita.PSet (t * 2, -d2tita * 4 + picGtita.ScaleHeight / 2), RGB(0, 0, 255)

        picGRadio.PSet (t * 2, -radio / 4 + picGRadio.ScaleHeight / 2)
        picGRadio.PSet (t * 2, -dradio / 4 + picGRadio.ScaleHeight / 2), RGB(255, 0, 0)
        picGRadio.PSet (t * 2, -d2radio / 4 + picGRadio.ScaleHeight / 2), RGB(0, 0, 255)


Else
        picGtita.PSet (t * 3, -Atn(Tan(tita)) * 3 + picGtita.ScaleHeight / 4)
        picGtita.PSet (t * 3, -dtita * 4 + 2 * picGtita.ScaleHeight / 4), RGB(255, 0, 0)
        picGtita.PSet (t * 3, -d2tita * 4 + 3 * picGtita.ScaleHeight / 4), RGB(0, 0, 255)

        picGRadio.PSet (t * 3, -radio / 4 + picGRadio.ScaleHeight / 4)
        picGRadio.PSet (t * 3, -dradio / 4 + 2 * picGRadio.ScaleHeight / 4), RGB(255, 0, 0)
        picGRadio.PSet (t * 3, -d2radio / 4 + 3 * picGRadio.ScaleHeight / 4), RGB(0, 0, 255)

End If
    mchica.Left = X - mchica.Width / 2
    mchica.Top = Y - mchica.Height / 2
    
    Espera 10
    
    t = t + dt
    
    x1 = mchica.Left
    x2 = Mgrande.Left
    y1 = mchica.Top
    y2 = Mgrande.Top
    
    'If ((x2 <= x1 + lc <= x2 + l) And (y2 <= y1 + lc <= y2 + l)) Then '((x2 <= x1 <= x2 + L) And (y2 <= y1 <= y1 + L)) Or ((x2 <= x1 + lc <= x2 + L) And (y2 <= y1 + lc <= y2 + L)) Then
    '    col.visible = True
    '    Debug.Print y2 & "-" & y1 + lc & "-" & y2 + l
    '    Exit Sub
    'End If
        
    Loop Until t >= 100 Or deTener = True
        Exit Sub
err:
'Debug.Print "ERROR"
col.visible = True
Exit Sub
Resume Next

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Command3_Click()
deTener = True

End Sub

Private Sub Form_Load()

inicial(0) = 0
inicial(1) = 1.92
inicial(2) = 0
inicial(3) = 30
inicial(4) = 1
inicial(5) = 0
inicial(6) = 4
inicial(7) = 0
'inicial(8) = 1.92
'inicial(9) = 1.92
'inicial(10) = 1.92
'inicial(1) = 1.92
'inicial(1) = 1.92

Pi = Atn(1) * 4


Me.Top = 0
    'MKS
    MTierra = 5.97E+24
    Rtierra = 6378140
    DistSol = 149597871000#
    Rsol = 696000000
    Msol = 1.9891E+30
    Mluna = 1.349E+22
    RLuna = 1737400
    DistLuna = 364288000

    Me.ScaleMode = vbPixels
    picUniv.ScaleMode = vbPixels
    picGtita.ScaleMode = vbPixels
    picGRadio.ScaleMode = vbPixels
    
    
     l = Mgrande.Height
     lc = mchica.Height
    
    g = 9.81
    Gr = 1 '0.0000000000667
    
    t = 0
    dt = 0.01
    
    tita = inicial(0) 'tita0
    dtita = inicial(1) 'dtita0
    d2tita = inicial(2) 'd2tita0
    radio = inicial(3) 'radio0
    dradio = inicial(4) 'dradio0
    d2radio = inicial(5) 'd2radio0
    
    kConst = inicial(6)
    lNat = inicial(7)
    
'    tita0 = 0
 '   radio0 = 40
 '
 '   dtita0 = 1.55
 '   dradio0 = 0
 '
 '   d2tita0 = 0
 '   d2radio0 = 0
 '
 '   tita = tita0
 '   dtita = dtita0
 '   radio = radio0
 '   dradio = dradio0
 '   d2tita = d2tita0
 '   d2radio = d2radio0
    
    mc = 1
    Mg = 100000 '1.49925037481259E+15
    
    
    
    Mgrande.Top = picUniv.ScaleHeight / 2 - Mgrande.Height / 2
    Mgrande.Left = picUniv.ScaleWidth / 2 - Mgrande.Width / 2
    
    mchica.Top = radio * Sin(tita) + picUniv.ScaleHeight / 2 - mchica.Height / 2
    mchica.Left = radio * Cos(tita) + picUniv.ScaleWidth / 2 - mchica.Width / 2

    
    Me.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

