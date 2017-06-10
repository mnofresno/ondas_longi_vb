VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Ley de Snell"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   Icon            =   "frmSnell.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3720
      Top             =   3000
   End
   Begin VB.ComboBox Cn 
      Height          =   315
      Index           =   1
      ItemData        =   "frmSnell.frx":0442
      Left            =   6360
      List            =   "frmSnell.frx":0464
      TabIndex        =   22
      Top             =   2400
      Width           =   975
   End
   Begin VB.ComboBox Cn 
      Height          =   315
      Index           =   0
      ItemData        =   "frmSnell.frx":04CA
      Left            =   6360
      List            =   "frmSnell.frx":04EC
      TabIndex        =   21
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox tRef 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6360
      TabIndex        =   16
      Text            =   "1"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox tRefDeg 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6360
      TabIndex        =   15
      Text            =   "1"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox tDeg 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   13
      Text            =   "1"
      Top             =   720
      Width           =   495
   End
   Begin VB.HScrollBar scT 
      Height          =   255
      Left            =   120
      Max             =   3142
      TabIndex        =   5
      Top             =   5160
      Value           =   1000
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox tAng 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Text            =   "1"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox tN 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   3
      Text            =   "1,5"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox tN 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   5400
      TabIndex        =   2
      Text            =   "1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox pSnell 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   720
      Width           =   4335
      Begin VB.Line lcero 
         BorderStyle     =   2  'Dash
         X1              =   1800
         X2              =   1800
         Y1              =   600
         Y2              =   3480
      End
      Begin VB.Line lrefle 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   3360
         X2              =   3120
         Y1              =   240
         Y2              =   1920
      End
      Begin VB.Line lref 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         X1              =   2400
         X2              =   4080
         Y1              =   2640
         Y2              =   3600
      End
      Begin VB.Line lInc 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         X1              =   2640
         X2              =   2880
         Y1              =   1080
         Y2              =   1920
      End
      Begin VB.Line lSup 
         X1              =   480
         X2              =   5160
         Y1              =   2280
         Y2              =   2280
      End
   End
   Begin VB.HScrollBar scAngulo 
      Height          =   255
      LargeChange     =   100
      Left            =   120
      Max             =   3142
      SmallChange     =   20
      TabIndex        =   0
      Top             =   360
      Value           =   1000
      Width           =   4335
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Angulo de refracción:"
      Height          =   195
      Left            =   4680
      TabIndex        =   19
      Top             =   4320
      Width           =   1515
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "*pi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6960
      TabIndex        =   18
      Top             =   4320
      Width           =   180
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "º"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6960
      TabIndex        =   17
      Top             =   4680
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "º"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5160
      TabIndex        =   14
      Top             =   720
      Width           =   75
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Angulo de incidencia"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "*pi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5160
      TabIndex        =   11
      Top             =   360
      Width           =   180
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Transmitido/Refractado"
      Height          =   195
      Left            =   5520
      TabIndex        =   10
      Top             =   3720
      Width           =   1680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Reflejado"
      Height          =   195
      Left            =   5520
      TabIndex        =   9
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Incidente"
      Height          =   195
      Left            =   5520
      TabIndex        =   8
      Top             =   3000
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "n2="
      Height          =   195
      Left            =   5040
      TabIndex        =   7
      Top             =   2400
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "n1="
      Height          =   195
      Left            =   5040
      TabIndex        =   6
      Top             =   1680
      Width           =   270
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   5040
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   5040
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   5040
      Top             =   3000
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Zx, Zy As Long
Const aMp = 2000


Private Sub Cn_Click(Index As Integer)
Select Case Cn(Index).ListIndex

    Case 0: 'Aire
        tN(Index).Text = 1
    Case 1: 'Vidrio
        tN(Index).Text = 1.5
    Case 2: 'Agua
        tN(Index).Text = 1.333
    Case 3: 'Azucar
        tN(Index).Text = 1.56
    Case 4: 'Diamante
        tN(Index).Text = 2.417
    Case 5: 'Mica
        tN(Index).Text = 1.58
    Case 6: 'Benceno
        tN(Index).Text = 1.504
    Case 7: 'Glicerina
        tN(Index).Text = 1.47
    Case 8: 'Alcohol etilico
        tN(Index).Text = 1.362
    Case 9: 'Aceite de oliva
        tN(Index).Text = 1.46
End Select
scAngulo_Change

End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()

pSnell.ScaleMode = vbpixel
pSnell.DrawWidth = pSnell.ScaleHeight / 30
Zx = pSnell.ScaleWidth / 2
Zy = pSnell.ScaleHeight / 2

Cn(0).ListIndex = 0
Cn(1).ListIndex = 1

lSup.X1 = 50
lSup.Y1 = Zy
lSup.Y2 = Zy
lSup.X2 = pSnell.ScaleWidth - 50

lInc.X2 = Zx
lInc.Y2 = Zy

lref.X1 = Zx
lref.Y1 = Zy

lrefle.X2 = Zx
lrefle.Y2 = Zy

lcero.X1 = Zx
lcero.X2 = Zx
lcero.Y1 = 500
lcero.Y2 = pSnell.ScaleHeight - 500

scAngulo_Change
End Sub

Private Sub scAngulo_Change()
lInc.X1 = Zx + aMp * Sin((scAngulo.Value - (scAngulo.Max / 2)) / 1000)
lInc.Y1 = Zy - aMp * Cos((scAngulo.Value - (scAngulo.Max / 2)) / 1000)

tDeg.Text = Format(180 * ((scAngulo.Value - (scAngulo.Max / 2)) / 3142), "0.00")
tAng.Text = Format(((scAngulo.Value - (scAngulo.Max / 2)) / 3142), "0.00")


lrefle.X1 = Zx - aMp * Sin((scAngulo.Value - (scAngulo.Max / 2)) / 1000)
lrefle.Y1 = Zy - aMp * Cos((scAngulo.Value - (scAngulo.Max / 2)) / 1000)

'ley de snell
tiTat = ASin((tN(0).Text / tN(1).Text) * Sin((scAngulo.Value - (scAngulo.Max / 2)) / 1000))

tRef.Text = Format(tiTat / 3.142, "0.00")
tRefDeg.Text = Format(180 * tiTat / 3.142, "0.00")

lref.X2 = Zx - aMp * Sin(tiTat)
lref.Y2 = Zy + aMp * Cos(tiTat)

End Sub

Function ASin(senito As Double) As Double
On Error GoTo pepe
    
        ASin = Atn(senito / Sqr(1 - senito * senito))
    
    Exit Function

pepe:
    ASin = 1.5707963267949 * Sgn(senito)
End Function


Private Sub scT_Change()
lref.X2 = Zx - aMp * Sin((scT.Value - (scT.Max / 2)) / 1000)
lref.Y2 = Zy + aMp * Cos((scT.Value - (scT.Max / 2)) / 1000)
End Sub

Private Sub tN_Change(Index As Integer)
pSnell.Cls

If tN(0) > tN(1) Then
pSnell.Line (0, pSnell.ScaleHeight / 4)-(pSnell.ScaleWidth, pSnell.ScaleHeight / 4), RGB(180, 225, 255)
Else
pSnell.Line (0, 3 * pSnell.ScaleHeight / 4)-(pSnell.ScaleWidth, 3 * pSnell.ScaleHeight / 4), RGB(180, 225, 255)
End If
End Sub
