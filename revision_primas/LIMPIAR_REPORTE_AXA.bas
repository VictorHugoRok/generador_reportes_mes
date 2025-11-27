Attribute VB_Name = "LIMPIAR_REPORTE_AXA"
Sub Nuevo_GenerarReportePolizas_FormatoCorrecto()
    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    Dim wbOrigen As Workbook
    Dim wsOrigen As Worksheet
    Dim sufijos As String
    Dim prefijo As String
    Dim listaSufijos As Variant
    Dim archivos As Variant
    Dim ultimaFilaOrigen As Long
    Dim i As Integer, fila As Long
    Dim rutaArchivo As String
    Dim filaDestino As Long
    Dim poliza As String
    Dim suf As Variant
    Dim coincideSufijo As Boolean
    Dim primaValor As Double
    Dim comisionValor As Double
    Dim primaTexto As String
    Dim comisionTexto As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --- Entrada de filtros ---
    sufijos = InputBox("Ingrese uno o varios sufijos separados por coma (por ejemplo: H, AH, 123):", "Filtro de pólizas por sufijo")
    If Trim(sufijos) = "" Then Exit Sub
    listaSufijos = Split(sufijos, ",")
    For i = LBound(listaSufijos) To UBound(listaSufijos)
        listaSufijos(i) = Trim(listaSufijos(i))
    Next i

    prefijo = InputBox("Ingrese el prefijo de las pólizas (por ejemplo: 1):", "Filtro de pólizas por inicio")
    If Trim(prefijo) = "" Then Exit Sub

    archivos = Application.GetOpenFilename(FileFilter:="Archivos de Excel (*.xls*), *.xls*", MultiSelect:=True, Title:="Seleccione los archivos")
    If Not IsArray(archivos) Then Exit Sub

    ' --- Hoja destino ---
    Set wbDestino = ThisWorkbook
    On Error Resume Next
    Set wsDestino = wbDestino.Sheets("Reporte Consolidado")
    If wsDestino Is Nothing Then
        Set wsDestino = wbDestino.Sheets.Add
        wsDestino.Name = "Reporte Consolidado"
    Else
        wsDestino.Cells.Clear
    End If
    On Error GoTo 0

    wsDestino.Range("A1:G1").Value = Array("Numero de Agente", "Nombre", "Poliza", "Vigor", "Dia de Aplicacion", "Prima Total", "Comisión")
    filaDestino = 2

    ' --- Procesar archivos ---
    For i = LBound(archivos) To UBound(archivos)
        rutaArchivo = archivos(i)
        Set wbOrigen = Workbooks.Open(rutaArchivo)
        Set wsOrigen = wbOrigen.Sheets(1)
        ultimaFilaOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row

        For fila = 2 To ultimaFilaOrigen
            poliza = Trim(wsOrigen.Cells(fila, "E").Value)
            coincideSufijo = False

            For Each suf In listaSufijos
                If Len(poliza) >= Len(suf) And Right(poliza, Len(suf)) = suf Then
                    coincideSufijo = True
                    Exit For
                End If
            Next suf

            If coincideSufijo And Left(poliza, Len(prefijo)) = prefijo Then

                ' === CONVERSIÓN DE PRIMA Y COMISIÓN ===
                primaTexto = ConvertirFormatoEuropeo(wsOrigen.Cells(fila, "K").Text)
                comisionTexto = ConvertirFormatoEuropeo(wsOrigen.Cells(fila, "P").Text)

                primaValor = IIf(IsNumeric(primaTexto), CDbl(primaTexto), 0)
                comisionValor = IIf(IsNumeric(comisionTexto), CDbl(comisionTexto), 0)

                ' --- Escribir en reporte ---
                With wsDestino
                    .Cells(filaDestino, 1).Value = wsOrigen.Cells(fila, "A").Value
                    .Cells(filaDestino, 2).Value = wsOrigen.Cells(fila, "D").Value
                    .Cells(filaDestino, 3).Value = wsOrigen.Cells(fila, "E").Value
                    .Cells(filaDestino, 4).Value = wsOrigen.Cells(fila, "G").Value
                    .Cells(filaDestino, 5).Value = wsOrigen.Cells(fila, "H").Value
                    .Cells(filaDestino, 6).Value = primaValor
                    .Cells(filaDestino, 7).Value = comisionValor
                    .Cells(filaDestino, 6).NumberFormat = "#,##0.00"  ' Formato mexicano
                    .Cells(filaDestino, 7).NumberFormat = "#,##0.00"
                End With

                filaDestino = filaDestino + 1
            End If
        Next fila
        wbOrigen.Close SaveChanges:=False
    Next i

    wsDestino.Columns("A:G").AutoFit
    MsgBox "Reporte generado correctamente.", vbInformation

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

' ===================================================================
' FUNCIÓN: Convierte formato europeo (primer punto = miles) ? formato VBA
' ===================================================================
Function ConvertirFormatoEuropeo(texto As String) As String
    Dim limpio As String
    Dim posPunto As Long, posComa As Long
    Dim tienePunto As Boolean, tieneComa As Boolean

    ' --- Limpiar ---
    limpio = Trim(texto)
    limpio = Replace(limpio, "$", "")
    limpio = Replace(limpio, "€", "")
    limpio = Replace(limpio, " ", "")
    limpio = Replace(limpio, Chr(160), "")

    If limpio = "" Then
        ConvertirFormatoEuropeo = "0"
        Exit Function
    End If

    posPunto = InStr(limpio, ".")
    posComa = InStr(limpio, ",")
    tienePunto = (posPunto > 0)
    tieneComa = (posComa > 0)

    ' === CASO 1: PRIMER SEPARADOR ES PUNTO ? FORMATO EUROPEO ===
    If tienePunto And (Not tieneComa Or posPunto < posComa) Then
        ' Ej: 23.342,09 ? punto es miles, coma es decimal
        limpio = Replace(limpio, ".", "")     ' Quitar puntos (miles)
        limpio = Replace(limpio, ",", ".")    ' Cambiar coma por punto decimal

    ' === CASO 2: SOLO PUNTO (sin coma) ? ES DE MILES ===
    ElseIf tienePunto And Not tieneComa Then
        limpio = Replace(limpio, ".", "")     ' Quitar punto ? 1.521 ? 1521

    ' === CASO 3: PRIMER SEPARADOR ES COMA ? FORMATO AMERICANO O SOLO DECIMAL ===
    ElseIf tieneComa Then
        ' Ej: 1,521.00 ? coma es miles ? quitar
        ' Ej: 1521,50 ? coma es decimal ? cambiar a punto
        If tienePunto And posComa < posPunto Then
            limpio = Replace(limpio, ",", "") ' coma es miles
        Else
            limpio = Replace(limpio, ",", ".") ' coma es decimal
        End If

    ' === CASO 4: SIN SEPARADORES O YA ES NÚMERO ===
    Else
        ' Ya está bien
    End If

    ' --- Validar ---
    If IsNumeric(limpio) Then
        ConvertirFormatoEuropeo = limpio
    Else
        ConvertirFormatoEuropeo = "0"
    End If
End Function
End Sub


