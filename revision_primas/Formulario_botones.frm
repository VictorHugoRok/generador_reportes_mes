VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formulario_botones 
   Caption         =   "UserForm1"
   ClientHeight    =   7680
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11304
   OleObjectBlob   =   "Formulario_botones.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Formulario_botones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton3_Click()

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


Private Sub CommandButton4_Click()
    Dim wsConsolidado As Worksheet
    Dim wsGMM As Worksheet
    Dim wsReporte As Worksheet
    Dim ultimaFilaCon As Long, ultimaFilaGMM As Long
    Dim filaDestino As Long
    Dim dicExcluir As Object
    Dim i As Long
    Dim poliza As String
    Dim valorP As Double

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '--- Hojas origen ---
    Set wsConsolidado = ThisWorkbook.Sheets("Reporte Consolidado")
    Set wsGMM = ThisWorkbook.Sheets("Polizas de GMM en 2025")

    '--- Crear diccionario con pólizas a EXCLUIR (P <= 0 o vacía) ---
    Set dicExcluir = CreateObject("Scripting.Dictionary")
    ultimaFilaGMM = wsGMM.Cells(wsGMM.Rows.Count, "E").End(xlUp).Row

    For i = 4 To ultimaFilaGMM
        poliza = UCase(Trim(wsGMM.Cells(i, "E").Value))
        If poliza <> "" Then
            If IsNumeric(wsGMM.Cells(i, "P").Value) Then
                valorP = wsGMM.Cells(i, "P").Value
            Else
                valorP = 0
            End If
            If valorP <= 0 Then
                dicExcluir(poliza) = True
            End If
        End If
    Next i

    Debug.Print "Pólizas a excluir (P=0): " & dicExcluir.Count

    '--- Crear hoja de reporte ---
    On Error Resume Next
    Set wsReporte = ThisWorkbook.Sheets("Comparativo Polizas")
    If wsReporte Is Nothing Then
        Set wsReporte = ThisWorkbook.Sheets.Add
        wsReporte.Name = "Comparativo Polizas"
    Else
        wsReporte.Cells.Clear
    End If
    On Error GoTo 0

    '--- Encabezados ---
    wsReporte.Range("A1:G1").Value = wsConsolidado.Range("A1:G1").Value
    filaDestino = 2

    '--- Recorrer consolidado (columna C = póliza) ---
    ultimaFilaCon = wsConsolidado.Cells(wsConsolidado.Rows.Count, "C").End(xlUp).Row

    For i = 2 To ultimaFilaCon
        poliza = UCase(Trim(wsConsolidado.Cells(i, "C").Value))
        If poliza <> "" Then
            ' Incluir si NO está en la lista de exclusión
            If Not dicExcluir.Exists(poliza) Then
                wsReporte.Range("A" & filaDestino & ":G" & filaDestino).Value = wsConsolidado.Range("A" & i & ":G" & i).Value
                wsReporte.Cells(filaDestino, "F").NumberFormat = "#,##0.00"
                wsReporte.Cells(filaDestino, "G").NumberFormat = "#,##0.00"
                filaDestino = filaDestino + 1
            End If
        End If
    Next i

    wsReporte.Columns("A:G").AutoFit

    MsgBox "Reporte generado en 'Comparativo Polizas'." & vbCrLf & _
           "Pólizas excluidas (P=0): " & dicExcluir.Count & vbCrLf & _
           "Pólizas incluidas: " & filaDestino - 2, vbInformation

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub CommandButton5_Click()
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
        .SortFields.Add Key:=wsDestino.Range("C2:C" & filaDestino - 1), _
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
