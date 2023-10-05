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
    
    
    
    ' Autoajustar columnas
    ws.Cells.EntireColumn.AutoFit

    ' Autoajustar filas
    ws.Cells.EntireRow.AutoFit
            
    
    Application.ScreenUpdating = True
    
End Sub
