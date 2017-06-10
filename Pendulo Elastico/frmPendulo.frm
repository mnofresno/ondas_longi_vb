VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPendulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comportamiento acelerómetros v0.89"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   Icon            =   "frmPendulo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   10
      Left            =   7800
      TabIndex        =   27
      Text            =   "0.05"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdExplorar 
      Caption         =   "..."
      Height          =   255
      Left            =   8280
      TabIndex        =   26
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox txtArchivo 
      Height          =   285
      Left            =   4560
      TabIndex        =   25
      Top             =   3840
      Width           =   3495
   End
   Begin VB.CheckBox chkTraza 
      Caption         =   "Traza trayectoria"
      Height          =   495
      Left            =   6120
      TabIndex        =   23
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CheckBox chkInextensible 
      Caption         =   "dr/dt = 0"
      Height          =   255
      Left            =   7440
      TabIndex        =   22
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   9
      Left            =   7800
      TabIndex        =   20
      Text            =   "100"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   8
      Left            =   5640
      TabIndex        =   18
      Text            =   "0.1"
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox picSimulador 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   240
      ScaleHeight     =   4035
      ScaleWidth      =   4035
      TabIndex        =   17
      Top             =   120
      Width           =   4095
      Begin MSComDlg.CommonDialog cdgDialogo 
         Left            =   600
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Line lnTecho 
         X1              =   0
         X2              =   4080
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line lnPend 
         BorderStyle     =   3  'Dot
         X1              =   2040
         X2              =   2040
         Y1              =   840
         Y2              =   2760
      End
      Begin VB.Shape shpMasita 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   960
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Inicio"
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   7
      Left            =   7800
      TabIndex        =   14
      Text            =   "0.5"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   6
      Left            =   7800
      TabIndex        =   12
      Text            =   "0"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   5
      Left            =   7800
      TabIndex        =   10
      Text            =   "59.8"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   4
      Left            =   7800
      TabIndex        =   8
      Text            =   "0"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   6
      Text            =   "9.8"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   4
      Text            =   "1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   2
      Text            =   "50"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   0
      Text            =   "1"
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Rozamiento gamma ="
      Height          =   195
      Index           =   10
      Left            =   6120
      TabIndex        =   28
      Top             =   240
      Width           =   1530
   End
   Begin VB.Label lblNota 
      Caption         =   "Al finalizar la simulación, se guardan los datos r, dr/dt, d2r/dt2, Fi, dFi/dt, d2Fi/dt2 en formato exportable a MS Excel."
      Height          =   915
      Left            =   4560
      TabIndex        =   24
      Top             =   3000
      Width           =   4575
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Tiempo total ="
      Height          =   195
      Index           =   9
      Left            =   6600
      TabIndex        =   21
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Diferencial dt ="
      Height          =   195
      Index           =   8
      Left            =   4440
      TabIndex        =   19
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Fi Inicial ="
      Height          =   195
      Index           =   7
      Left            =   6600
      TabIndex        =   15
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "dFi/dt Inicial ="
      Height          =   195
      Index           =   6
      Left            =   6600
      TabIndex        =   13
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "R Inicial ="
      Height          =   195
      Index           =   5
      Left            =   6600
      TabIndex        =   11
      Top             =   960
      Width           =   705
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "dR/dt Inicial ="
      Height          =   195
      Index           =   4
      Left            =   6600
      TabIndex        =   9
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Parametro g ="
      Height          =   195
      Index           =   3
      Left            =   4560
      TabIndex        =   7
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Parametro m ="
      Height          =   195
      Index           =   2
      Left            =   4560
      TabIndex        =   5
      Top             =   1320
      Width           =   1020
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Parametro l0 ="
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   3
      Top             =   960
      Width           =   1020
   End
   Begin VB.Label lblParam 
      AutoSize        =   -1  'True
      Caption         =   "Parametro k ="
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   990
   End
End
Attribute VB_Name = "frmPendulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Semi Constantes
Dim Masa, Gravedad, kElast, LongL0 As Double

'Variables y primeras derivadas
Dim dRdt, Radio, dTitadt, Tita As Double

'Segundas derivadas
Dim d2Titadt2, d2Rdt2 As Double

'Fuerzas tangencial y radial
Dim Ftang, Fradial As Double

'Variables de tiempo
Dim Tiempo, dTiempo As Double

'Arrays para el armado de la tabla
Dim TitulosTabla(6) As String
Dim DatosTabla As Variant

'Indices para el llenado de la tabla
Dim indiceFila, col As Long

'Vector de posición cartesiana
Dim posM(1) As Double

'Constante de rozamiento viscoso
Dim GamaRoz As Double
'Constante de tiempo máximo(futuramente debería fijarla el usuario)
Dim Tmax As Long
Sub fijarValores()
    'Fijado de constantes
    kElast = Val(txtParam(0).Text)
    LongL0 = Val(txtParam(1).Text)
    Masa = Val(txtParam(2).Text)
    Gravedad = Val(txtParam(3).Text)
    GamaRoz = Val(txtParam(10).Text)
    
    'Condiciones iniciales
    dRdt = Val(txtParam(4).Text)
    Radio = Val(txtParam(5).Text)
    dTitadt = Val(txtParam(6).Text)
    Tita = Val(txtParam(7).Text)
    
    'Fijamos el tiempo de integración numérica
    dTiempo = Val(txtParam(8).Text)
    
    'Fijamos el tiempo total del ensayo
    Tmax = Val(txtParam(9).Text)
    
    'Ponemos el tiempo a cero
    Tiempo = 0
    

    'Tamaño de masita proporcional
    shpMasita.Width = Int(Masa * 9)
    shpMasita.Height = Int(Masa * 9)

    'Ubico resorte punto 1
    lnPend.X1 = picSimulador.ScaleWidth / 2
    lnPend.Y1 = lnTecho.Y1
    
    'Limpiamos el picturebox
    picSimulador.Cls
End Sub

Sub dibujarPendulo()

    posM(0) = picSimulador.ScaleWidth / 2 + Radio * Sin(Tita)
    posM(1) = lnTecho.Y1 + (Radio) * Cos(Tita)
    
    'Ubico resorte punto 2
    lnPend.X2 = posM(0)
    lnPend.Y2 = posM(1)
    
    'Ubico masita en posición según radio y tita
    shpMasita.Left = posM(0) - shpMasita.Width / 2
    shpMasita.Top = posM(1) - shpMasita.Height / 2
    
    If chkTraza.Value = 1 Then picSimulador.PSet (posM(0), posM(1)), RGB(255, 0, 0)
    
End Sub

Sub ArmarEncabezados()
    'Titulos de la tabla
    TitulosTabla(0) = "Tiempo"
    TitulosTabla(1) = "Radio"
    TitulosTabla(2) = "dr/dt"
    TitulosTabla(3) = "d2r/dt2"
    TitulosTabla(4) = "Fi"
    TitulosTabla(5) = "dFi/dt"
    TitulosTabla(6) = "d2Fi/dt2"
End Sub
Sub CargarFila()
        'Imprimo los datos de la simulación sobre la tabla
        DatosTabla(indiceFila, 0) = Tiempo
        DatosTabla(indiceFila, 1) = Radio
        DatosTabla(indiceFila, 2) = dRdt
        DatosTabla(indiceFila, 3) = d2Rdt2
        DatosTabla(indiceFila, 4) = Tita
        DatosTabla(indiceFila, 5) = dTitadt
        DatosTabla(indiceFila, 6) = d2Titadt2
End Sub

Private Sub cmdExplorar_Click()
    'Activo el error por cancelación
    'cdgDialogo.CancelError = True
    'Archivo por defecto y verificación de coherencia (para evitar "\\")
    cdgDialogo.FileName = App.Path & IIf(Right(App.Path, 1) <> "\", "\datos.csv", "datos.csv")
    'Filtros de archivo
    cdgDialogo.Filter = "Valores separados por tabulación (*.csv)|*.csv|Archivo de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    'Mostrar diálogo
    cdgDialogo.ShowSave
    'Fijar archivo elegido
    txtArchivo.Text = cdgDialogo.FileName
End Sub

Private Sub cmdStart_Click()

    fijarValores

    
    'Ponemos a cero el cursor que nos ubica en filas de la tabla
    indiceFila = 0
    
    
    'Redimensionamos el tamaño de la matriz de datos para armar la tabla
    ReDim DatosTabla(Int(Tmax / dTiempo) + 1, 6) As Double
    

        
        
    ArmarEncabezados
    
    Do                 'Bucle
        DoEvents       'No trabar la PC para mientras se realiza la simulación
            
            
        CargarFila

        dibujarPendulo
            
        'Fuerzas tangencial y radial
        Ftang = -Masa * Gravedad * Sin(Tita) - GamaRoz * dTitadt * Radio
        Fradial = Masa * Gravedad * Cos(Tita) - kElast * (Radio - LongL0) - GamaRoz * dRdt '- dRdt * 100
        
        'Aceleraciones (derivadas segundas)
        d2Rdt2 = (Fradial / Masa) + Radio * dTitadt ^ 2
        d2Titadt2 = (Ftang / Masa - 2 * dRdt * dTitadt) / Radio
        
        'Derivadas primeras (integro numericamente)
        dRdt = IIf(chkInextensible, 0, dRdt + d2Rdt2 * dTiempo)
        dTitadt = dTitadt + d2Titadt2 * dTiempo
        
        'Funciones del tiempo (derivada nula, integro numericamente)
        Radio = Radio + dRdt * dTiempo
        Tita = Tita + dTitadt * dTiempo
        
        'Avance del tiempo
        Tiempo = Tiempo + dTiempo
        indiceFila = indiceFila + 1
        
        Espera dTiempo * 10 'Elijo una demora para el sprite proporcional a mi diferencial t
        txtParam(9).Text = Tiempo
    Loop Until Tiempo >= Tmax                      'Iterar repetidamente hasta alcanzar el tiempo prefijado
    txtParam(9).Text = Tmax
    On Error GoTo Errhandle
    If txtArchivo.Text = "" Then 'cmdExplorar_Click    'Si no tenemos nombre de archivo, explorar en busca-de
    
    End If
    
    If Dir(txtArchivo.Text) <> "" Then Kill txtArchivo.Text         'Matar el archivo anterior si existe
    GuardarTablaEnArchivo TitulosTabla, DatosTabla, txtArchivo.Text           'Generar tabla y guardarla
Errhandle:
    Exit Sub
End Sub


Private Sub chkTraza_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not chkTraza.Value Then picSimulador.Cls
End Sub

Private Sub Form_Load()
    picSimulador.ScaleMode = vbPixels           'Cambio la escala para trabajar en pixels
    fijarValores                                'Fijamos los valores de las constantes y parámetros iniciales
    lblNota.Caption = lblNota.Caption & vbCrLf & vbCrLf & "Nombre de archivo:"   'Completa el rótulo explicativo
    dibujarPendulo                              'Dibuja el péndulo y el elástico
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End 'Fin del programa
End Sub

