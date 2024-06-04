Sub ExportToPDF()
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim wb As Workbook
    Dim pdfPath As String
    Dim pdfFilename As String
    Dim ws As Worksheet
    
    ' Ruta de la carpeta que contiene los archivos Excel
    Dim folderPath As String
    folderPath = "modelo6\origen\ficheros\iniciales\salida\" ' Cambia esta ruta por la ruta de tu carpeta
    
    ' Inicializar el objeto FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la referencia a la carpeta
    Set objFolder = objFSO.GetFolder(folderPath)
    
    ' Recorrer todos los archivos en la carpeta
    For Each objFile In objFolder.Files
        ' Verificar si el archivo es un archivo Excel
        If objFSO.GetExtensionName(objFile.Path) = "xlsx" Or objFSO.GetExtensionName(objFile.Path) = "xls" Then
            ' Abrir el archivo Excel
            Set wb = Workbooks.Open(objFile.Path)
            
            ' Exportar cada hoja del libro a un archivo PDF
            For Each ws In wb.Worksheets
                pdfPath = folderPath
                pdfFilename = Replace(objFSO.GetBaseName(objFile.Name), " ", "_") & "_" & Replace(ws.Name, " ", "_") & ".pdf"
                ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath & pdfFilename, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            Next ws
            
            ' Cerrar el libro sin guardar cambios
            wb.Close SaveChanges:=False
        End If
    Next objFile
    
    ' Liberar recursos
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set objFile = Nothing
    Set wb = Nothing
    
    MsgBox "La exportación a PDF ha finalizado.", vbInformation
End Sub
