Sub ExportToPDF()
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim pdfFilename As String
    Dim folderPath As String

    ' Solicitar al usuario que ingrese el directorio donde se encuentran los archivos Excel
    folderPath = InputBox("Por favor, ingrese el directorio donde se encuentran los archivos Excel:", "Directorio de Archivos Excel")

    ' Verificar si se ingresó un directorio válido
    If folderPath = "" Then
        MsgBox "Se canceló la operación.", vbInformation
        Exit Sub
    ElseIf Not Dir(folderPath, vbDirectory) <> "" Then
        MsgBox "El directorio ingresado no es válido.", vbExclamation
        Exit Sub
    End If

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
            
            ' Recorrer todas las hojas del libro
            For Each ws In wb.Worksheets
                ' Exportar la hoja actual a un archivo PDF
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
    Set ws = Nothing

    MsgBox "La exportación a PDF ha finalizado.", vbInformation
End Sub
