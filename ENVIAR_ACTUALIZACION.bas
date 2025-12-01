Attribute VB_Name = "ENVIAR_ACTUALIZACION"
Sub EnviarActualizacion()

    Dim wbNew As Workbook
    Dim wsVida As Worksheet, wsGMM As Worksheet
    Dim nameWsVida As String
    Dim nameWsGMM As String
    Dim originalPath As String
    Dim newFileName As String
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Set wsVida = ThisWorkbook.Sheets("Polizas de VIDA en 2025")
    Set wsGMM = ThisWorkbook.Sheets("Polizas de GMM en 2025")
    
    originalPath = ThisWorkbook.Path
    newFileName = "Actualizacion_reporte_" & Format(Date, "yyyy-mm-dd") & ".xlsx"
    
    wsVida.Copy
    Set wbNew = ActiveWorkbook
    wsGMM.Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
    
    
    If originalPath = "" Then
        MsgBox "El libro no se ha guardardo", vbCritical
        wbNew.Close SaveChanges:=False
        GoTo CleanUp
        
    End If
    
    wbNew.SaveAs Filename:=originalPath & "\" & newFileName, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False
    
CleanUp:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    If Err.Number = 0 Then
    MsgBox "Las hojas y el libro se han guardado correctamente en " & vbCrLf & originalPath & "\" & newFileName, vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    Resume CleanUp
    
End Sub
