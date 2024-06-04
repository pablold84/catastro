Attribute VB_Name = "Módulo2"

Sub ExportToPDF()
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim pdfFilename As String
    Dim mergedPDF As String
    
    ' Ruta de la carpeta que contiene los archivos Excel
    Dim folderPath As String
    folderPath = "C:\Trabajo\catastro\modelo6\origen\ficheros\iniciales\salida\" ' Ruta de la carpeta con los archivos Excel
    
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
            
            ' Crear un nuevo archivo PDF temporal
            pdfPath = folderPath
            pdfFilename = Replace(objFSO.GetBaseName(objFile.Name), " ", "_") & ".pdf"
            mergedPDF = pdfPath & pdfFilename & "_temp.pdf"
            ' Crear un archivo PDF temporal vacío
            Open mergedPDF For Output As #1
            Close #1
            
            ' Recorrer todas las hojas del libro
            For Each ws In wb.Worksheets
                ' Exportar la hoja actual al archivo PDF temporal
                ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=mergedPDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            Next ws
            
            ' Renombrar el archivo PDF temporal al destino final
            Name mergedPDF As pdfPath & pdfFilename
            
            ' Cerrar el libro sin guardar cambios
            wb.Close SaveChanges:=False
        End If
    Next objFile
    
    ' Liberar recursos
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set objFile = Nothing
    Set wb = Nothing
    Set ws = Nothing
    
    MsgBox "La exportación a PDF ha finalizado.", vbInformation
End Sub
