Attribute VB_Name = "REVISAR_PRIMA_CC_GMM"
Sub ValidarCoincidenciaPolizas_NuevaEstructura_FiltroV2()

    Dim wb As Workbook, wbPagadas As Workbook
    Dim wsRegistro As Worksheet, wsPagadas As Worksheet
    Dim ultimaFilaReg As Long, ultimaFilaPag As Long
    Dim dictRegistro As Object, dictPagadas As Object
    Dim mesSeleccionado As String
    Dim i As Long
    
 
    Const COL_POLIZA_REG As Long = 5        ' E (Columna PÓLIZA en tu nueva estructura)
    Const COL_MES_REG As Long = 7           ' G (Columna MES DE EMISIÓN en tu nueva estructura)
    Const FILA_ENCABEZADOS As Long = 3      ' Los encabezados están en la Fila 3
    Const COL_POLIZA_PAG_EXT As Long = 5    ' Columna E en el archivo externo (Se mantiene del código original)
    
    Dim pagadasOk As Long, pagadasNo As Long
    Dim colorVerde As Long, colorRojo As Long
    Dim clave As Variant
    Dim poliza As String, mesRegistro As String
    Dim rutaArchivo As Variant
    Dim cel As Range
    Dim sec As MsoAutomationSecurity

    
    colorVerde = RGB(198, 239, 206)
    colorRojo = RGB(255, 199, 206)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

   
    rutaArchivo = Application.GetOpenFilename("Archivos de Excel (*.xls*), *.xls*", , "Seleccione el archivo de pólizas pagadas")
    If rutaArchivo = False Then
        MsgBox "No se seleccionó ningún archivo. Operación cancelada.", vbExclamation
        GoTo Salir
    End If

    ' === Abrir el archivo externo sin ejecutar macros ===
    sec = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set wb = ThisWorkbook
    Set wbPagadas = Workbooks.Open(rutaArchivo)
    Application.AutomationSecurity = sec

    
    Set wsRegistro = wb.Sheets("Polizas de GMM en 2025")
    
    
    If wbPagadas.Sheets.Count <> 1 Then
        MsgBox "? El archivo externo debe tener exactamente una hoja.", vbCritical
        wbPagadas.Close False
        GoTo Salir
    End If
    Set wsPagadas = wbPagadas.Sheets(1)

   
    mesSeleccionado = InputBox("Ingrese el nombre del mes del reporte (ejemplo: ENERO, FEBRERO...)", "Seleccionar mes")
    If mesSeleccionado = "" Then
        MsgBox "No se seleccionó mes. Operación cancelada.", vbExclamation
        wbPagadas.Close False
        GoTo Salir
    End If
    mesSeleccionado = UCase(Trim(Replace(mesSeleccionado, "Í", "I"))) ' Normalizar mes

    
    Set dictRegistro = CreateObject("Scripting.Dictionary")
    Set dictPagadas = CreateObject("Scripting.Dictionary")

   
    ultimaFilaReg = wsRegistro.Cells(wsRegistro.Rows.Count, COL_POLIZA_REG).End(xlUp).Row
    ultimaFilaPag = wsPagadas.Cells(wsPagadas.Rows.Count, COL_POLIZA_PAG_EXT).End(xlUp).Row

    
    For i = FILA_ENCABEZADOS + 1 To ultimaFilaReg
        Set cel = wsRegistro.Cells(i, COL_POLIZA_REG)
        If cel.Interior.Color <> colorVerde Then cel.Interior.ColorIndex = xlNone
    Next i
    For i = 2 To ultimaFilaPag
        Set cel = wsPagadas.Cells(i, COL_POLIZA_PAG_EXT)
        If cel.Interior.Color <> colorVerde Then cel.Interior.ColorIndex = xlNone
    Next i


    For i = FILA_ENCABEZADOS + 1 To ultimaFilaReg
        poliza = Trim(wsRegistro.Cells(i, COL_POLIZA_REG).Value)
        mesRegistro = UCase(Trim(wsRegistro.Cells(i, COL_MES_REG).Value))
        
        
        Dim cumplePatron As Boolean
        cumplePatron = (poliza Like "1*U00" Or poliza Like "1*V00")
        
        If cumplePatron And (InStr(1, mesRegistro, mesSeleccionado, vbTextCompare) > 0) Then
            If Not dictRegistro.Exists(poliza) Then dictRegistro(poliza) = i
        End If
    Next i

    
    For i = 2 To ultimaFilaPag
        poliza = Trim(wsPagadas.Cells(i, COL_POLIZA_PAG_EXT).Value)
        
        
        cumplePatron = (poliza Like "1*U00" Or poliza Like "1*V00")
        
        If cumplePatron Then
            If Not dictPagadas.Exists(poliza) Then dictPagadas(poliza) = i
        End If
    Next i

    
    For Each clave In dictRegistro.Keys
        Set cel = wsRegistro.Cells(dictRegistro(clave), COL_POLIZA_REG)
        If dictPagadas.Exists(clave) Then
            cel.Interior.Color = colorVerde
            wsPagadas.Cells(dictPagadas(clave), COL_POLIZA_PAG_EXT).Interior.Color = colorVerde
            pagadasOk = pagadasOk + 1
        Else
            If cel.Interior.Color <> colorVerde Then cel.Interior.Color = colorRojo
            pagadasNo = pagadasNo + 1
        End If
    Next clave

    
    For Each clave In dictPagadas.Keys
        Set cel = wsPagadas.Cells(dictPagadas(clave), COL_POLIZA_PAG_EXT)
        If Not dictRegistro.Exists(clave) Then
            If cel.Interior.Color <> colorVerde Then cel.Interior.Color = colorRojo
        End If
    Next clave

   
    MsgBox "? Validación completada para el mes: " & mesSeleccionado & vbCrLf & _
           "? **Filtro aplicado:** Pólizas que inician con '1' y terminan en 'U00' o 'V00'." & vbCrLf & vbCrLf & _
           "?? Pólizas coincidentes en ambas hojas: " & pagadasOk & vbCrLf & _
           "?? Pólizas en tu archivo no encontradas en el externo: " & pagadasNo & vbCrLf & vbCrLf & _
           "?? Archivo analizado: " & vbCrLf & wbPagadas.Name, vbInformation

   
    wbPagadas.Close False

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub
