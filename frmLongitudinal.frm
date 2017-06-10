VERSION 5.00
Begin VB.Form frmLongitudinal 
   Caption         =   "Ondas Longitudinales En un Tubo con Gas (por MMF)"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ClipControls    =   0   'False
   Icon            =   "frmLongitudinal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMainFrame 
      Caption         =   "Parámetros"
      Height          =   2775
      Left            =   360
      TabIndex        =   2
      Top             =   4920
      Width           =   8175
      Begin VB.VScrollBar sldModo 
         Height          =   2175
         Left            =   7440
         Max             =   14
         Min             =   1
         TabIndex        =   38
         Top             =   360
         Value           =   1
         Width           =   255
      End
      Begin VB.CommandButton cmdEqui 
         Caption         =   "Equilibrio"
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdDetener 
         Caption         =   "detener"
         Height          =   255
         Left            =   3120
         TabIndex        =   32
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdPaso 
         Caption         =   "step"
         Height          =   255
         Left            =   4440
         TabIndex        =   31
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdContinuar 
         Caption         =   "continuar"
         Height          =   255
         Left            =   5760
         TabIndex        =   30
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdAyu 
         Caption         =   "ayuda"
         Height          =   255
         Left            =   2520
         TabIndex        =   29
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "salir"
         Height          =   255
         Left            =   5160
         TabIndex        =   28
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtParam 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2640
         TabIndex        =   25
         Text            =   "340"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtParam 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   960
         TabIndex        =   23
         Text            =   "340"
         Top             =   960
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   6360
         Top             =   840
      End
      Begin VB.VScrollBar vsTiempo 
         Height          =   615
         Left            =   5880
         Max             =   500
         Min             =   1
         TabIndex        =   20
         Top             =   1080
         Value           =   10
         Width           =   255
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   7
         Left            =   5280
         TabIndex        =   17
         Text            =   "0,005"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   6
         Left            =   5280
         TabIndex        =   16
         Text            =   "0,05"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtParam 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   4080
         TabIndex        =   13
         Text            =   "340"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   4
         Left            =   3600
         TabIndex        =   12
         Text            =   "0,5"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtParam 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   8
         Text            =   "1,225"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Text            =   "141610"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtParam 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Text            =   "30"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdAplica 
         Caption         =   "A&plicar"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblModo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Modo: 1"
         Height          =   375
         Left            =   6360
         TabIndex        =   27
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         Caption         =   "k[rad/m]="
         Height          =   195
         Index           =   9
         Left            =   1920
         TabIndex        =   26
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         Caption         =   "Lambda[m]="
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tiempo t=0"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "T=10"
         Height          =   195
         Left            =   5400
         TabIndex        =   21
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         Caption         =   "dt="
         Height          =   195
         Index           =   7
         Left            =   5040
         TabIndex        =   19
         Top             =   600
         Width           =   225
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         Caption         =   "A="
         Height          =   195
         Index           =   6
         Left            =   5040
         TabIndex        =   18
         Top             =   240
         Width           =   195
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         Caption         =   "c[m/s]="
         Height          =   195
         Index           =   5
         Left            =   3480
         TabIndex        =   15
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         Caption         =   "L[m]="
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Width           =   390
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         Caption         =   "f[Hz]="
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   11
         Top             =   600
         Width           =   420
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         Caption         =   "Ro[Kg/m^3]="
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   960
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         Caption         =   "K[Kg/m*s^2]="
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblParam 
         AutoSize        =   -1  'True
         Caption         =   "N="
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   210
      End
   End
   Begin VB.PictureBox pTransv 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   17955
      TabIndex        =   1
      Top             =   1680
      Width           =   18015
      Begin VB.VScrollBar vsD 
         Height          =   3135
         Left            =   360
         Max             =   200
         Min             =   1
         TabIndex        =   37
         Top             =   0
         Value           =   100
         Width           =   255
      End
      Begin VB.VScrollBar vsP 
         Height          =   3135
         Left            =   17160
         Max             =   200
         Min             =   1
         TabIndex        =   36
         Top             =   0
         Value           =   100
         Width           =   255
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Presión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   10680
         TabIndex        =   35
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label lblD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desplazamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   0
         TabIndex        =   34
         Top             =   1680
         Width           =   2235
      End
   End
   Begin VB.PictureBox pLongi 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   1440
      Left            =   0
      ScaleHeight     =   1380
      ScaleWidth      =   9840
      TabIndex        =   0
      Top             =   120
      Width           =   9900
   End
End
Attribute VB_Name = "frmLongitudinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim dX As Double
Dim X As Double
Dim pi As Double
Dim k As Double
Dim Cv As Double 'velocidad de propagacion virtual

Dim C As Double  'Velocidad de propagacion real
Dim w As Double
Dim t As Double


Dim dfdx As Double


Const Kcor = 1000000#

Dim dt As Double

Dim espP, eW, eH As Integer 'Ancho y altura efectivas de la pictureBox
Dim escP, escD As Single

Dim Np As Integer
Dim Lp As Double

Const Alt = 75
Dim m As Integer  'Número de modo normal

Dim lambda As Double

Dim k0 As Double
Dim Ro0 As Double

Dim Amp As Double


Function f(xp As Double) As Double
    lambda = 2 * Lp / m
    k = 2 * pi / lambda
    w = k * Cv
    
    f = Amp * Sin(k * xp) * Cos(w * t)
    dfdx = Amp * k * Cos(k * xp) * Cos(w * t)
    'f = Amp * Sin(k * xp) * Cos(w * t) + Amp * Sin(3 * k * xp) * Cos(3 * w * t) + Amp * Sin(5 * k * xp) * Cos(w * t)
    'f = Amp * Sin(k * xp - w * t)
End Function

Sub MostrarPs()
    eW = pLongi.Width - 2 * espP
    eH = pLongi.Height - espP
    
    On Error Resume Next
    pLongi.Cls
    pTransv.Cls
    DrawRuler
    For i = 0 To Np
        X = i * dX
        pLongi.Line (espP + (f(X) + X / Lp) * eW, (pLongi.Height - Alt) / 2)-(espP + (f(X) + X / Lp) * eW, (pLongi.Height + Alt) / 2), &HFF
    pTransv.PSet (espP + (X / Lp) * (pTransv.Width - espP * 2), (1 / 2 - f(X) * escD / 12) * pTransv.Height), &HFF
    pTransv.PSet (espP + (X / Lp) * (pTransv.Width - espP * 2), (1 / 2 + dfdx * escP / 70) * pTransv.Height), &HFF0000
    
    Next i
End Sub

Sub DrawRuler()
    pTransv.Line (0, pTransv.Height / 2)-(pTransv.Width, pTransv.Height / 2)
    For i = 0 To Np
        X = i * dX
        pLongi.Line (espP + (X / Lp) * eW, pLongi.Height - 10)-(espP + (X / Lp) * eW, pLongi.Height)
        pTransv.Line (pTransv.Width / 2, i * pTransv.Height / Np)-(pTransv.Width / 2, i * pTransv.Height / Np)
    Next i
End Sub

Private Sub cmdAplica_Click()
    CargarParametros
End Sub

Private Sub cmdDetener_Click()
    Timer1.Enabled = False
End Sub

Private Sub cmdEqui_Click()
    t = pi / (2 * w)
    MostrarPs
End Sub



Private Sub cmdPaso_Click()
    Timer1_Timer
End Sub

Private Sub cmdContinuar_Click()
    Timer1.Enabled = True
End Sub
Sub CargarParametros()
    On Error Resume Next
    Np = txtParam(0).Text - 1
    Lp = txtParam(4).Text
    
    pTransv.ScaleMode = vbPixels
    dX = Lp / Np
    pi = Atn(1) * 4
    pLongi.ScaleMode = vbPixels
    Me.ScaleMode = vbPixels
    espP = 60
    
    
    pTransv.DrawWidth = 3
    pLongi.DrawWidth = 1
    Amp = txtParam(6).Text
    k0 = txtParam(1).Text / Kcor
    
     Ro0 = txtParam(2).Text
    dt = txtParam(7).Text
    
    Cv = Sqr(k0 / Ro0)
    C = Sqr((Cv ^ 2) * Kcor)
    txtParam(5).Text = Format(C, "0.00")
    escP = vsP.Value
    escD = vsD.Value
    sldModo_Click
End Sub

Private Sub cmdSalir_Click()
    End
End Sub

Private Sub Form_Load()
    txtParam(0) = 30
    txtParam(1) = 141610
    txtParam(2) = 1.225
    txtParam(4) = 0.5
    txtParam(5) = 340
    txtParam(6) = 0.05
    txtParam(7) = 0.05
    CargarParametros
End Sub

Private Sub Form_Resize()
    pLongi.Left = 3
    pTransv.Left = 3
    
    'frMainFrame.Width = Me.Width / Screen.TwipsPerPixelX - 15
    
    pLongi.Width = Me.Width / Screen.TwipsPerPixelX - 15
    pTransv.Width = Me.Width / Screen.TwipsPerPixelX - 15
    
    lblD.Left = vsD.Width * 1.1
    lblP.Left = pTransv.Width - lblP.Width * 1.1 - vsP.Width * 1.1
    
    frMainFrame.Left = (Me.Width / Screen.TwipsPerPixelX - frMainFrame.Width) / 2
    frMainFrame.Top = (Me.Height / Screen.TwipsPerPixelY - frMainFrame.Height) - 50
    
    vsD.Left = 0
    vsP.Left = pTransv.Width - vsP.Width * 1.1 - 2
    vsD.Height = pTransv.Height - 3
    vsP.Height = pTransv.Height - 3
End Sub

Private Sub sldModo_Click()
    m = sldModo.Value
    txtParam(3).Text = Format((C * m) / (2 * Lp), "0.00")
    lblModo = "Modo: " & m
    txtParam(8).Text = Format(lambda, "0.00")
    txtParam(9).Text = Format(k, "0.00")
End Sub

Private Sub sldModo_Change()
    sldModo_Click
End Sub

Private Sub sldModo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sldModo_Click
End Sub

Private Sub sldModo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sldModo_Click
End Sub

Private Sub Timer1_Timer()
    t = t + dt
    MostrarPs
    Label1.Caption = "Tiempo t=" & Format(t, "0.00")
    CargarParametros
End Sub

Private Sub vsD_Change()
    escD = vsD.Value
End Sub

Private Sub vsP_Change()
    escP = vsP.Value
End Sub

Private Sub vsTiempo_Change()
    Timer1.Interval = vsTiempo.Value
    Label2.Caption = "T=" & Timer1.Interval
End Sub
