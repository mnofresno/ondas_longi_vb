Attribute VB_Name = "library"

'API para generar tiempo de delay en la ejecucion de un programa
Public Declare Function GetTickCount Lib "kernel32" () As Long



Public Sub Espera(ByVal TmEspera As Long)
    Dim TmFin As Long
    TmFin = GetTickCount + TmEspera
    Do While GetTickCount < TmFin
        DoEvents
    Loop
    
End Sub



'Función para guardar tablas en archivos
Function GuardarTablaEnArchivo(ByVal vTitulos As Variant, ByVal vDatos As Variant, ByVal sArchivo As String)

    Dim sFila, sTabla As String                        'Buffer de fila
    
    Dim f, c As Long                                   'f e i Indices fila y columna respectivamente
    
    sTabla = vbNullString                              'Inicializo la tabla
       
            For c = LBound(vTitulos) To UBound(vTitulos)    'Recorro el array de titulos
                sFila = sFila & vTitulos(c) & vbTab
            Next c
            
            sTabla = sTabla & sFila & vbCrLf                    'Imprimo la línea con títulos
            sFila = ""                                          'Limpio el buffer de fila

            
    For f = 0 To UBound(vDatos, 1) - 1                          'Recorremos el array por filas
    
     

            For c = LBound(vDatos, 2) To UBound(vDatos, 2) 'Recorremos las columnas c, de la presente fila f
                sFila = sFila & vDatos(f, c) & vbTab
            Next c
        
            sTabla = sTabla & sFila & vbCrLf                    'Imprimo una línea con datos
            sFila = vbNullString                                'Limpio el buffer de fila
    Next f
    
    Open sArchivo For Output As #1                              'Abro el archivo
        Print #1, sTabla                                        'Imprimo la tabla completa
    Close #1                                                    'Cierro el archivo
    sTabla = vbNullString
End Function




