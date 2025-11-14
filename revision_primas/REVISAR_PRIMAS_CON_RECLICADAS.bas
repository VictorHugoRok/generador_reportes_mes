Attribute VB_Name = "REVISAR_PRIMAS_CON_RECLICADAS"
Sub ValidarCoincidenciaPolizas_FiltroSufijo()

    Dim wb As Workbook, wbPagadas As Workbook
    Dim wsRegistro As Worksheet, wsPagadas As Worksheet
    Dim ultimaFilaReg As Long, ultimaFilaPag As Long
    Dim dictRegistro As Object, dictPagadas As Object
    Dim sufijoSeleccionado As String ' Nueva variable para el sufijo
    Dim i As Long
    
    ' --- Constantes de Columna y Fila ---
    Const COL_POLIZA_REG As Long = 5        ' E (Columna PÓLIZA en tu nueva estructura)
    ' COL_MES_REG (G) ya no se usa, pero la dejamos como referencia si la necesitas más tarde.
    Const FILA_ENCABEZADOS As Long = 3      ' Los encabezados están en la Fila 3
    Const COL_POLIZA_PAG_EXT As Long = 5    ' Columna E en el archivo externo
    
    Dim pagadasOk As Long, pagadasNo As Long
    Dim colorVerde As Long, colorRojo As Long
    Dim clave As Variant
    Dim poliza As String
    Dim rutaArchivo As Variant
    Dim cel As Range
    Dim sec As MsoAutomationSecurity
    Dim sufijos() As String ' Array para manejar múltiples sufijos

    ' --- Inicialización de Colores ---
    colorVerde = RGB(198, 239, 206)
    sinColor = RGB(255, 255, 255)

    ' --- Desactivar Funciones de Excel ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' --- Solicitud de Sufijo ---
    sufijoSeleccionado = InputBox("Ingrese la **terminación** o sufijo de las pólizas que desea validar (ejemplo: U00, V00, 999)." & vbCrLf & _
                                  "Si son varios, sepárelos por **comas** (ejemplo: U00,V00):", "Ingresar Sufijo de Póliza")
    
    If sufijoSeleccionado = "" Then
        MsgBox "No se ingresó ninguna terminación. Operación cancelada.", vbExclamation
        GoTo Salir
    End If
    
    ' Dividir los sufijos ingresados por coma y convertirlos a mayúsculas
    sufijoSeleccionado = UCase(Trim(sufijoSeleccionado))
    sufijos = Split(sufijoSeleccionado, ",")
    
    ' --- Selección y Apertura del Archivo ---
    rutaArchivo = Application.GetOpenFilename("Archivos de Excel (*.xls*), *.xls*", , "Seleccione el archivo de pólizas pagadas")
    If rutaArchivo = False Then
        MsgBox "No se seleccionó ningún archivo. Operación cancelada.", vbExclamation
        GoTo Salir
    End If

    ' Abrir el archivo externo sin ejecutar macros
    sec = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set wb = ThisWorkbook
    Set wbPagadas = Workbooks.Open(rutaArchivo)
    Application.AutomationSecurity = sec
    
    ' --- Establecer Hojas de Trabajo ---
    Set wsRegistro = wb.Sheets("Polizas de GMM en 2025")
    
    If wbPagadas.Sheets.Count <> 1 Then
        MsgBox "?? El archivo externo debe tener exactamente una hoja.", vbCritical
        wbPagadas.Close False
        GoTo Salir
    End If
    Set wsPagadas = wbPagadas.Sheets(1)

    ' --- Inicializar Diccionarios y Últimas Filas ---
    Set dictRegistro = CreateObject("Scripting.Dictionary")
    Set dictPagadas = CreateObject("Scripting.Dictionary")

    ultimaFilaReg = wsRegistro.Cells(wsRegistro.Rows.Count, COL_POLIZA_REG).End(xlUp).Row
    ultimaFilaPag = wsPagadas.Cells(wsPagadas.Rows.Count, COL_POLIZA_PAG_EXT).End(xlUp).Row

    ' --- Limpiar Formato de Pólizas (Excepto el color verde) ---
    For i = FILA_ENCABEZADOS + 1 To ultimaFilaReg
        Set cel = wsRegistro.Cells(i, COL_POLIZA_REG)
        If cel.Interior.Color <> colorVerde Then cel.Interior.ColorIndex = xlNone
    Next i
    For i = 2 To ultimaFilaPag
        Set cel = wsPagadas.Cells(i, COL_POLIZA_PAG_EXT)
        If cel.Interior.Color <> colorVerde Then cel.Interior.ColorIndex = xlNone
    Next i

    ' --- Función de Validación de Sufijo ---
    ' Verifica si la póliza termina con alguno de los sufijos ingresados
    Dim suf As Variant
    Dim cumpleSufijo As Boolean
    
    ' Llenar diccionario de pólizas de Registro (Ahora sin filtro por mes)
    For i = FILA_ENCABEZADOS + 1 To ultimaFilaReg
        poliza = Trim(wsRegistro.Cells(i, COL_POLIZA_REG).Value)
        cumpleSufijo = False
        
        ' Verificar cada sufijo ingresado
        For Each suf In sufijos
            ' Se usa Like "*sufijo" para buscar la terminación.
            If UCase(poliza) Like "*" & Trim(suf) Then
                cumpleSufijo = True
                Exit For
            End If
        Next suf
        
        If cumpleSufijo Then
            If Not dictRegistro.Exists(poliza) Then dictRegistro(poliza) = i
        End If
    Next i

    ' Llenar diccionario de pólizas Pagadas
    For i = 2 To ultimaFilaPag
        poliza = Trim(wsPagadas.Cells(i, COL_POLIZA_PAG_EXT).Value)
        cumpleSufijo = False
        
        For Each suf In sufijos
            If UCase(poliza) Like "*" & Trim(suf) Then
                cumpleSufijo = True
                Exit For
            End If
        Next suf
        
        If cumpleSufijo Then
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
            
            If cel.Interior.Color <> colorVerde Then cel.Interior.Color = sinColor
            pagadasNo = pagadasNo + 1
        End If
    Next clave

   
    For Each clave In dictPagadas.Keys
        Set cel = wsPagadas.Cells(dictPagadas(clave), COL_POLIZA_PAG_EXT)
        If Not dictRegistro.Exists(clave) Then
            If cel.Interior.Color <> colorVerde Then cel.Interior.Color = colorRojo
        End If
    Next clave

    
    MsgBox "? Validación completada para todas las pólizas." & vbCrLf & _
           "?? **Terminaciones validadas (sufijos):** " & sufijoSeleccionado & vbCrLf & vbCrLf & _
           "?? Pólizas coincidentes en ambas hojas: " & pagadasOk & vbCrLf & _
           "?? Pólizas en tu archivo NO encontradas en el externo: " & pagadasNo & vbCrLf & vbCrLf & _
           "?? Archivo analizado: " & vbCrLf & wbPagadas.Name, vbInformation

    
    wbPagadas.Close False

Salir:
    ' --- Reactivar Funciones de Excel ---
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub

