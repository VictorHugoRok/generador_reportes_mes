Attribute VB_Name = "FILTRACION_NUEVOS_AXA"
Sub SacarPolizasParaReporte_GMM()
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

   '---- Se excluyen las pólizas sin nuevos miembros 
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

    Debug.Print "Polizas a excluir (P=0): " & dicExcluir.Count

    
    On Error Resume Next
    Set wsReporte = ThisWorkbook.Sheets("Comparativo Polizas")
    If wsReporte Is Nothing Then
        Set wsReporte = ThisWorkbook.Sheets.Add
        wsReporte.Name = "Comparativo Polizas"
    Else
        wsReporte.Cells.Clear
    End If
    On Error GoTo 0

    wsReporte.Range("A1:G1").Value = wsConsolidado.Range("A1:G1").Value
    filaDestino = 2

    ultimaFilaCon = wsConsolidado.Cells(wsConsolidado.Rows.Count, "C").End(xlUp).Row

    For i = 2 To ultimaFilaCon
        poliza = UCase(Trim(wsConsolidado.Cells(i, "C").Value))
        If poliza <> "" Then
       
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
           "P�lizas excluidas (P=0): " & dicExcluir.Count & vbCrLf & _
           "P�lizas incluidas: " & filaDestino - 2, vbInformation

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub


