Attribute VB_Name = "RANKING_MES_GMM"
Sub GenerarRankingAgentes()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim filaDestino As Long
    Dim dict As Object
    Dim i As Long
    Dim agente As String
    Dim prima As Double
    Dim clave As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Verificar que exista la hoja "Reporte Consolidado"
    On Error Resume Next
    Set wsOrigen = ThisWorkbook.Sheets("Comparativo Polizas")
    On Error GoTo 0

    If wsOrigen Is Nothing Then
        MsgBox "No se encontró la hoja 'Comparativo Polizas'. Primero genere el reporte.", vbExclamation
        Exit Sub
    End If

    ' Crear o limpiar hoja destino
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Ranking de Agentes")
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Sheets.Add
        wsDestino.Name = "Ranking de Agentes"
    Else
        wsDestino.Cells.Clear
    End If
    On Error GoTo 0

    ' Crear diccionario para acumular primas por agente
    Set dict = CreateObject("Scripting.Dictionary")

    ' Determinar última fila con datos
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row

    ' Recorrer los datos del reporte consolidado
    For i = 2 To ultimaFila
        agente = Trim(wsOrigen.Cells(i, "A").Value)
        prima = Val(wsOrigen.Cells(i, "F").Value)

        If agente <> "" Then
            If Not dict.Exists(agente) Then
                dict.Add agente, prima
            Else
                dict(agente) = dict(agente) + prima
            End If
        End If
    Next i

    ' Encabezados del ranking
    wsDestino.Range("A1:C1").Value = Array("Ranking", "Número de Agente", "Prima Total Acumulada")

    filaDestino = 2
    For Each clave In dict.Keys
        wsDestino.Cells(filaDestino, 2).Value = clave
        wsDestino.Cells(filaDestino, 3).Value = dict(clave)
        filaDestino = filaDestino + 1
    Next clave

 
    With wsDestino.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsDestino.Range("C2:C" & filaDestino - 1), _
                        Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange wsDestino.Range("A1:C" & filaDestino - 1)
        .Header = xlYes
        .Apply
    End With

    ' Asignar números de ranking
    Dim r As Long
    For r = 2 To filaDestino - 1
        wsDestino.Cells(r, 1).Value = r - 1
    Next r

    ' Formato final
    wsDestino.Columns("A:C").AutoFit

    MsgBox "Ranking de agentes generado correctamente en la hoja 'Ranking de Agentes'.", vbInformation

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

