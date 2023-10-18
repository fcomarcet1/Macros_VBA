Sub FormaBDPtInventory()

    Dim numCols As Integer 'nº columnas
    Dim iCols As Integer
    Dim col As String
    
    Dim numRows As Long 'nº filas
    Dim iRows As Long
    Dim Fila As String

    Dim Hoja As Worksheet
    Dim ws As Worksheet
    
    Dim lastCol As Integer
    Dim stockCol As Integer

    Dim Celda As Range
    
    Dim filaAlmacenC As Integer
    Dim filaCodA As Integer
    Dim filaCodD As Integer
    Dim filaFamiliaE As Integer
    
    Dim MacroEjecutada As Boolean
    Dim ejecuciones As Integer
    
    'If ejecuciones > 0 Then
    
        Application.ScreenUpdating = False
        
        Set ws = ActiveSheet
        
        ' Eliminar columnas vacías
        numCols = ActiveSheet.UsedRange.Columns.Count
    
        For iCols = numCols To 1 Step -1
            If WorksheetFunction.CountA(Cells(1, iCols).EntireColumn) = 0 Then
                Cells(1, iCols).EntireColumn.Delete
            End If
        Next iCols
        
        ' Eliminar filas vacías
        numRows = ActiveSheet.UsedRange.Rows.Count
    
        For iRows = numRows To 1 Step -1
            Fila = iRows & ":" & iRows
            If WorksheetFunction.CountA(Range(Fila)) = 0 Then
                Range("A" & iRows).EntireRow.Delete
            End If
        Next iRows
        
         ' Eliminar la fila 5
        Rows("5:5").Select
        Selection.Delete Shift:=xlUp
        
        ' Elimina las columnas L y M
        Columns("L:M").Delete
        
        ' Añadir Columnas Amacen y Barra con el titulo en la 5º celda
        lastCol = ActiveSheet.Cells(5, Columns.Count).End(xlToLeft).Column
        Columns(lastCol + 1).Insert Shift:=xlToRight
        Columns(lastCol + 2).Insert Shift:=xlToRight
        
        ' Colocar el título en la celda 5 de las nuevas columnas
        Cells(5, lastCol + 1).Value = "Almacen"
        Cells(5, lastCol + 2).Value = "Barra"
        
        
        ' Poner a 0 los valores de la columna J los cuales tengan cod en la columna
        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
            
        For i = 6 To lastRow
            If Not IsEmpty(Cells(i, "H").Value) Or Cells(i, "H").Value Like "" Then
                ws.Cells(i, "J").NumberFormat = "0.00"
                ws.Cells(i, "L").NumberFormat = "0.00"
                ws.Cells(i, "M").NumberFormat = "0.00"
                ws.Cells(i, "J").Value = 0
            End If
        Next i
        
            
        For i = 6 To lastRow
            If Not IsEmpty(Cells(i, "H").Value) Or (ws.Cells(i, "H").Value <> "") Then
                ' Establece la fórmula en la columna Stock (Columna J) para las filas existentes
                ws.Cells(i, "J").Formula = "=SUM(L" & i & ":M" & i & ")"
            End If
        Next i
        
        ' Eliminar los valores 0 de la columna J los cuales no tengan cod en la columna
        For i = 6 To lastRow
            If IsEmpty(Cells(i, "H").Value) Then
                Cells(i, "J").ClearContents
            End If
        Next i
        
        
        ws.Columns("L:L").Select
        
        Application.CutCopyMode = False
        
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        
        ws.Columns("M:M").Select
        
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        
        ' Mover registros para poder establecer filtro
        lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
        ' Inicializar la variable de la fila de Almacén
        filaAlmacenC = 0
        filaCodA = 0
        filaCodD = 0
        filaFamiliaE = 0
        
        
        ' Recorrer cada celda en la columna C usando un bucle For
        For i = 1 To lastRow
            ' Verificar si la celda contiene "Almacén"
            If Not IsEmpty(ws.Cells(i, "C").Value) And ws.Cells(i, "C").Value Like "*Almacén*" Then
                ' Guardar la fila de la celda que contiene "Almacén"
                filaAlmacenC = i
                Exit For ' Salir del bucle si se encuentra "Almacén"
            End If
        Next i
        
        ' Verificar si se encontró "Almacén"
        If filaAlmacenC > 0 Then
            ' Cortar el contenido de la siguiente celda y pegarlo dos celdas abajo
            ActiveSheet.Cells(filaAlmacenC + 1, 3).Cut Destination:=ActiveSheet.Cells(filaAlmacenC + 3, 3)
    
            ' Limpiar la celda original
            ActiveSheet.Cells(filaAlmacenC + 1, 3).Clear
    
        End If
        
        
    
        ' Recorrer cada celda en la columna A usando un bucle For
        For i = 1 To lastRow
            ' Verificar si la celda contiene "Almacén"
            If Not IsEmpty(ws.Cells(i, "A").Value) And ws.Cells(i, "A").Value Like "*Cód.*" Then
                ' Guardar la fila de la celda que contiene "Cod."
                filaCodA = i
                Exit For ' Salir del bucle si se encuentra "Cod."
            End If
        Next i
        
        ' Verificar si se encontró "Cod."
        
        If filaCodA > 0 Then
            ' Leer el valor de la celda combinada en la columna A y limpiarla
            Dim valorCod As Variant
            valorCod = ws.Cells(filaCodA + 1, 1).MergeArea.Value
            ws.Cells(filaCodA + 1, 1).MergeArea.Clear
            
            ' Pegar el valor dos filas más abajo
            ws.Cells(filaCodA + 3, 1).Value = valorCod
        End If
        
        
        
        ' Recorrer cada celda en la columna D usando un bucle For
        For i = 1 To lastRow
            ' Verificar si la celda contiene "Almacén"
            If Not IsEmpty(ws.Cells(i, "D").Value) And ws.Cells(i, "D").Value Like "*Cód.*" Then
                ' Guardar la fila de la celda que contiene "Cod."
                filaCodD = i
                Exit For ' Salir del bucle si se encuentra "Cod."
            End If
        Next i
        
        ' Verificar si se encontró "Cod."
        
        If filaCodD > 0 Then
            ' Leer el valor de la celda combinada en la columna A y limpiarla
            valorCod = ws.Cells(filaCodD + 2, 4).MergeArea.Value
            ws.Cells(filaCodD + 2, 4).MergeArea.Clear
            
            ' Pegar el valor dos filas más abajo
            ws.Cells(filaCodD + 3, 4).Value = valorCod
        End If
        
        
        ' Recorrer cada celda en la columna F usando un bucle For
        For i = 1 To lastRow
            ' Verificar si la celda contiene "Familia"
            If Not IsEmpty(ws.Cells(i, "E").Value) And ws.Cells(i, "E").Value Like "*Familia*" Then
                ' Guardar la fila de la celda que contiene "Cod."
                filaFamiliaE = i
                Exit For ' Salir del bucle si se encuentra "Cod."
            End If
        Next i
        
        ' Verificar si se encontró "Cod."
        
        If filaFamiliaE > 0 Then
            ' Leer el valor de la celda combinada en la columna E y limpiarla
            valorCod = ws.Cells(filaFamiliaE + 2, 5).MergeArea.Value
            ws.Cells(filaFamiliaE + 2, 5).MergeArea.Clear
            
            ' Pegar el valor dos filas más abajo
            ws.Cells(filaFamiliaE + 3, 5).Value = valorCod
        End If
        
        ws.Columns("A:A").Select
        
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        
        ws.Columns("D:D").Select
        
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        
           ws.Columns("E:E").Select
        
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        
        ' Eliminar filas vacías
        numRows = ActiveSheet.UsedRange.Rows.Count
    
        For iRows = numRows To 1 Step -1
            Fila = iRows & ":" & iRows
            If WorksheetFunction.CountA(Range(Fila)) = 0 Then
                Range("A" & iRows).EntireRow.Delete
            End If
        Next iRows
        
        ' Obtener el número de la última columna en la primera fila
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' Aplicar estilos a la primera fila
        With ws.Rows(1).Font
            .Name = "Arial"
            .Bold = True ' Hacer el texto en negrita
            .Size = 12 ' Tamaño de la fuente
        End With
        With ws.Rows(2).Font
            .Name = "Arial"
            .Bold = True ' Hacer el texto en negrita
            .Size = 12 ' Tamaño de la fuente
        End With
        With ws.Rows(3).Font
            .Name = "Arial"
            .Bold = True ' Hacer el texto en negrita
            .Size = 12 ' Tamaño de la fuente
        End With
        With ws.Rows(4).Font
            .Name = "Arial"
            .Bold = True ' Hacer el texto en negrita
            .Size = 12 ' Tamaño de la fuente
        End With
        
        
        ws.Range("A5:M5").Select
        Selection.Interior.Color = RGB(255, 255, 153) ' Color de fondo amarillo
        
        ' añadir filtro
        ws.Range("A5:M5").Select
        Selection.AutoFilter
        
        ' Encontrar el rango de celdas con contenido
        Set rng = ws.UsedRange
        
        ' Aplicar bordes a todas las celdas con contenido
        With rng.Borders
            .LineStyle = xlContinuous ' Tipo de línea continua
            .Color = RGB(0, 0, 0) ' Color de borde negro
            .Weight = xlThin ' Grosor del borde delgado
        End With
        
        ActiveWindow.DisplayGridlines = False
        
        ' Autoajustar columnas
        ws.Cells.EntireColumn.AutoFit
    
        ' Autoajustar filas
        ws.Cells.EntireRow.AutoFit
        
        Application.ScreenUpdating = True
        
        ' Incrementar el contador de ejecuciones
        ejecuciones = ejecuciones + 1
    'End If
    'Else
        'MsgBox "La macro ya ha sido ejecutada anteriormente. No puedes volver a ejecutarla", vbExclamation
    
    
End Sub

Sub ExportarComoPDF()
    Dim ws As Worksheet
    Dim rng As Range
    Dim savePath As String
    
    ' Establecer la hoja de trabajo
    Set ws = ActiveSheet ' Reemplaza "Nombre de tu Hoja" con el nombre real de tu hoja
    
    ' Ajustar el ancho de las columnas para que quepan los datos
    ws.Cells.Columns.AutoFit
    
    ' Encontrar el rango de celdas con contenido
    Set rng = ws.UsedRange
    
    ' Definir la ruta de guardado del archivo PDF
    savePath = Application.GetSaveAsFilename(FileFilter:="Archivos PDF (*.pdf), *.pdf", Title:="Guardar PDF como")
    
    ' Verificar si se seleccionó un archivo y exportar como PDF
    If savePath <> "Falso" Then
        ' Establecer la orientación del PDF como horizontal
        ws.PageSetup.Orientation = xlLandscape
        
        ' Exportar la hoja como PDF
        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath, Quality:=xlQualityStandard
    Else
        MsgBox "No se seleccionó una ubicación para guardar el PDF.", vbExclamation
    End If
End Sub


Sub EliminarFilasCondicionalmente()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    
    ' Encuentra la última fila con datos en la columna I.
    LastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
     For i = 2 To LastRow
     
        ' Convierte el valor en la celda a formato de tiempo (si es posible).
        On Error Resume Next
            Dim valorComoTiempo As Date
            valorComoTiempo = CDate(ws.Cells(i, 9).Value)
        On Error GoTo 0
        
        ' Verifica si el valor es igual a "00:00:00" en formato de tiempo.
        If (valorComoTiempo = TimeValue("00:00:00")) And (InStr(1, ws.Cells(i, "A").Value, "Empleado :", vbTextCompare) > 0) Then
            ' Pone en blanco la celda si cumple con la condición.
            ws.Cells(i, "J").Value = ""
        End If
        
    Next i
End Sub

Sub ObtenerTipoDeDato()
    Dim MiCelda As Range
    Dim TipoDato As String
    
    ' Establece la celda de la que deseas conocer el tipo de dato (cambia "A1" por la celda que desees).
    Set MiCelda = ThisWorkbook.Sheets("Hoja1").Range("J3")
    
    ' Obtén el tipo de dato de la celda.
    TipoDato = TypeName(MiCelda.Value)
    
    ' Muestra el tipo de dato en una ventana emergente.
    MsgBox "El tipo de dato en la celda es: " & TipoDato
End Sub

Sub InsertarDosFilasEnBlanco()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    
    ' Establece la hoja de trabajo en la que deseas trabajar (cambia "NombreDeTuHoja" al nombre real de tu hoja).
    Set ws = ActiveSheet
    
    ' Encuentra la última fila con datos en la columna V.
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    

    For i = 2 To LastRow
         If ws.Cells(i, "A").Value Like "*Firma Empleado*" Then
                    ws.Rows(i).Copy
                    ws.Rows(i + 1).Resize(2).Insert Shift:=xlDown
                    ws.Cells(i + 1, "A").Value = ""
                    ws.Cells(i + 1, "E").Value = ""
                    ws.Cells(i + 2, "A").Value = ""
                    ws.Cells(i + 2, "E").Value = ""
                    ws.Rows(i + 1).PasteSpecial Paste:=xlPasteFormats
         End If
    Next i
End Sub


Sub ExtraerHora()
    Dim Hoja As Worksheet
    Dim Celda As Range
    
    ' Establece la hoja de trabajo en la que deseas realizar la operación.
    Set Hoja = ActiveSheet
    
    ' Recorre todas las celdas en la columna A hasta la última fila con datos.
    For Each Celda In Hoja.Range("I1:I" & Hoja.Cells(Hoja.Rows.Count, "I").End(xlUp).Row)
        ' Verifica si la celda tiene un valor de fecha válido.
        If IsDate(Celda.Value) Or IsNumeric(Celda.Value) Then
            ' Extrae la parte de la hora y asigna el valor a la misma celda.
            Celda.Value = Format(CDate(Celda.Value), "hh:mm:ss")
        End If
        
    Next Celda
End Sub


Sub SumaHastaCeldaVacia()
        ' Suma Total Horas Reales
        Dim total As Double
        Dim Hoja As Worksheet
        Dim Celda As Range
        Dim LastRow As Long
        Dim ws As Worksheet
        
        Set ws = ActiveSheet
        
        LastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row + 1
        
        total = 0
        
         For i = 1 To LastRow
            If ws.Cells(i, "A").Value Like "Empleado :*" Then
               total = 0
            End If
            
            If (IsDate(ws.Cells(i, "J").Value) Or IsNumeric(ws.Cells(i, "J").Value)) And (IsEmpty(ws.Cells(i, "A").Value)) Then
               total = ws.Cells(i, "J").Value + total
            End If
            
            If ws.Cells(i, "A").Value Like "Total Semana" Then
                ws.Cells(i, "J").NumberFormat = "[h]:mm:ss;@"
                ws.Cells(i, "J").Value = total
                ' ws.Cells(i, "J").Value = Format(CDate(ws.Cells(i, "J").Value), "hh:mm:ss")
                'ws.Cells(i, "J").NumberFormat = "[h]:mm:ss;@"
                total = 0
            End If
        Next i
        
        
End Sub
