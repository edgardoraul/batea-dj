Attribute VB_Name = "Dj"
Option Explicit
Sub GenerarTablaDesdeCarpeta()
    ' Genera un archivo excel cuyo contenido, es el listado de temas que tiene la carpeta donde se encuentra.
    Dim dlg As FileDialog
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileCount As Integer
    Dim i As Integer
    Dim nombre As String
    
    ' Seleccionar carpeta con el cuadro de diálogo
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    
    If dlg.Show = -1 Then
        folderPath = dlg.SelectedItems(1)
    Else
        MsgBox "Elegí una carpeta, por favor! Caramba!.", vbExclamation
        folderPath = dlg.SelectedItems(1)
    End If
    
    
    nombre = Right(dlg.SelectedItems(1), Len(dlg.SelectedItems(1)) - Len(dlg.InitialFileName))
    Debug.Print nombre
    
    ' Crear nuevo libro de trabajo
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    
    ' Establecer título de la carpeta en A1
    ws.Range("A1").Value = nombre
    
    ' Enumerar archivos en la carpeta
    fileName = Dir(folderPath & "\*")
    fileCount = 2 ' Empezamos desde la fila 2
    
    Do While fileName <> ""
        ' Añadir nombre del archivo a la tabla sin extensión
        ws.Cells(fileCount, 1).Value = Left(fileName, InStrRev(fileName, ".") - 1)
        fileCount = fileCount + 1
        fileName = Dir
    Loop
    
    ' Guardar archivo en la misma carpeta con el nombre de la carpeta y extensión .xlsx
    wb.SaveAs folderPath & "\" & nombre & ".xlsx", FileFormat:=xlWorkbookDefault
    
    ' Dar formato: Centra el titular y formatea para imprimir la etiqueta.
    ' Todavía en desarrollo.

End Sub