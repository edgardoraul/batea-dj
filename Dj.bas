Attribute VB_Name = "Dj"
Option Explicit
Public verde As Variant
Public azul As Variant
Public fucsia As Variant
Public naranja As Variant
Public marron As Variant
Public gris As Variant




Sub GenerarTablaDesdeCarpeta()
Attribute GenerarTablaDesdeCarpeta.VB_ProcData.VB_Invoke_Func = "K\n14"

    ' Genera un archivo excel cuyo contenido, es el listado de temas que tiene la carpeta donde se encuentra.
    Dim dlg As FileDialog
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileCount As Integer
    Dim i As Integer
    Dim nombre As String
    
    ' Seleccionar carpeta con el cuadro de diÃ¡logo
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    
    If dlg.Show = -1 Then
        folderPath = dlg.SelectedItems(1)
    Else
        MsgBox "ElegÃ­ una carpeta, por favor! Caramba!.", vbExclamation
        folderPath = dlg.SelectedItems(1)
    End If
    
    
    nombre = Right(dlg.SelectedItems(1), Len(dlg.SelectedItems(1)) - Len(dlg.InitialFileName))
    Debug.Print nombre
    
    ' Crear nuevo libro de trabajo
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    
    ' Establecer tÃ­tulo de la carpeta en A1
    ws.Range("A1").Value = nombre
    
    ' Enumerar archivos en la carpeta
    fileName = Dir(folderPath & "\*")
    fileCount = 2 ' Empezamos desde la fila 2
    
    Do While fileName <> ""
        ' AÃ±adir nombre del archivo a la tabla sin extensiÃ³n
        ws.Cells(fileCount, 1).Value = Left(fileName, InStrRev(fileName, ".") - 1)
        fileCount = fileCount + 1
        fileName = Dir
    Loop
    
    ' Dar formato
    Call darFormato
    ' Guardar archivo en la misma carpeta con el nombre de la carpeta y extensiÃ³n .xlsx
    wb.SaveAs folderPath & "\" & nombre & ".xlsx", FileFormat:=xlWorkbookDefault
    
    ' Dar formato: Centra el titular y formatea para imprimir la etiqueta.
    ' TodavÃ­a en desarrollo.

End Sub

Sub darFormato()
    ' Colores
    verde = RGB(60, 120, 62)
    azul = RGB(49, 119, 203)
    fucsia = RGB(230, 70, 219)
    naranja = RGB(228, 109, 10)
    marron = RGB(132, 108, 60)
    gris = RGB(145, 135, 125)
    
    Dim color As String
    Dim estilo As String
    Dim Pregunta As String

    Dim seleccion As String
    Dim opciones As String
    Dim resultado As Byte
    
    ' Definir las opciones con sus respectivos números
    opciones = "1. Deep House" & vbCrLf & _
               "2. Progressive House" & vbCrLf & _
               "3. Progressive" & vbCrLf & _
               "4. Trance" & vbCrLf & _
               "5. House" & vbCrLf & _
               "6. Acid Jazz"
    
    ' Mostrar el cuadro de diálogo con las opciones
    seleccion = InputBox("¿Qué estilo es? (ingresa el número correspondiente):" & vbCrLf & vbCrLf & opciones, "Opciones")
    
    ' Verificar si el usuario hizo clic en Cancelar o no ingresó nada
    If seleccion = "" Then
        MsgBox "No se seleccionó ninguna opción.", vbInformation
    End If
    
    ' Determinar la opción seleccionada
    Select Case seleccion
        Case "1"
            resultado = 1
        Case "2"
            resultado = 2
        Case "3"
            resultado = 3
        Case "4"
            resultado = 4
        Case "5"
            resultado = 5
        Case "6"
            resultado = 6
        Case Else
            MsgBox "Selección inválida. Por favor, ingresa un número del 1 al 6.", vbExclamation
    End Select

    
    Select Case resultado
        Case 1
            color = verde
        Case 2
            color = marron
        Case 3
            color = naranja
        Case 4
            color = fucsia
        Case 5
            color = azul
        Case 6
            color = gris
        Case Else
            MsgBox "No hay color"
    End Select
    
    Debug.Print color
    
    'Da formato de impresión
    With Range("A1")
        .Borders(xlEdgeTop).color = gris
        .Font.Size = 20
        .Font.Bold = True
        .Font.color = color
        .HorizontalAlignment = xlCenter
        .EntireColumn.ColumnWidth = 40
        .EntireColumn.WrapText = True
    End With
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Borders(xlEdgeBottom).color = gris

    Call FormatoImpresion

End Sub

Sub FormatoImpresion()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim rangoImpresion As Range

    ' Establecer la hoja de trabajo activa
    Set ws = ActiveSheet

    ' Encontrar la última fila con datos
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Definir el rango de impresión incluyendo una fila adicional
    Set rangoImpresion = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaFila + 1, ws.UsedRange.Columns.Count))

    ' Establecer el rango de impresión
    ws.PageSetup.PrintArea = rangoImpresion.Address

    ' Configurar el ancho de la hoja
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
    End With

    ' Ajustar el ancho de las columnas para que se ajusten al ancho de 15 cm
    AjustarAnchoColumna ws, 40

    ' Previsualizar la impresión
    'ws.PrintPreview
End Sub

Sub AjustarAnchoColumna(ws As Worksheet, anchoTotalCm As Double)
    Dim columna As Range
    Dim totalAnchoPuntos As Double
    Dim anchoCmPorPunto As Double
    Dim factorAjuste As Double
    Dim i As Integer
    
    ' Calcular el ancho total en puntos
    totalAnchoPuntos = 0
    For i = 1 To ws.UsedRange.Columns.Count
        totalAnchoPuntos = totalAnchoPuntos + ws.Columns(i).ColumnWidth
    Next i

    ' Calcular el factor de ajuste
    anchoCmPorPunto = anchoTotalCm / totalAnchoPuntos

    ' Ajustar el ancho de cada columna
    For i = 1 To ws.UsedRange.Columns.Count
        ws.Columns(i).ColumnWidth = ws.Columns(i).ColumnWidth * anchoCmPorPunto
    Next i
End Sub

