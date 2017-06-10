Attribute VB_Name = "library"
Option Explicit

Public Const MAX_PATH = 260
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Espera(TmEspera As Long)
    Dim TmFin As Long
    TmFin = GetTickCount + TmEspera
    Do While GetTickCount < TmFin
        DoEvents
    Loop
End Sub


