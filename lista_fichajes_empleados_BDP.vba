Option Explicit
Sub elige()
    If Range("B1") = 1 Then genera
    If Range("B1") = 2 Then EliminarFilasEnBlanco
    If Range("B1") = 3 Then EliminarColumnasVacias
End Sub

Sub FormateoInicial()
    Range("A1:R8").Select
    Selection.Delete Shift:=xlUp
End Sub


Sub EliminarFilasEnBlanco()

    Dim n As Long 'nº filas
    Dim i As Long
    Dim Fila As String

    n = ActiveSheet.UsedRange.Rows.Count
    For i = n To 1 Step -1
       Fila = i & ":" & i
       If WorksheetFunction.CountA(Range(Fila)) = 0 Then
          Range("A" & i).EntireRow.Delete
       End If
    Next i
End Sub



Sub DeleteRowsWithWord()

    'Declare variables
    Dim rng As Range
    Dim word As String
    Dim cell As Range
    Dim word1 As String
    Dim word2 As String

    word1 = "TOTAL ......."
    word2 = "GENERAL TOTAL ......."
    
    'Set the range to the active sheet
    Set rng = ActiveSheet.UsedRange
    
    'Loop through the range, deleting rows that contain the word
    For Each cell In rng
      If cell.Value Like "*" & word1 & "*" Then
        rng.Rows(cell.Row).Delete
      End If
      If cell.Value Like "*" & word2 & "*" Then
        rng.Rows(cell.Row).Delete
      End If
    Next cell

End Sub

Sub EliminarColumnasVacias()

    Dim n As Integer 'nº columnas
    Dim i As Integer
    Dim col As String

    n = ActiveSheet.UsedRange.Columns.Count

    For i = n To 1 Step -1
        If WorksheetFunction.CountA(Cells(1, i).EntireColumn) = 0 Then
            Cells(1, i).EntireColumn.Delete
        End If
    Next i
End Sub

Sub EliminarFilasEnBlanco_bis()
    Dim Fila As Long

    For Fila = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
        If WorksheetFunction.CountA(ActiveSheet.Rows(Fila)) = 0 Then
            Cells(Fila, 1).EntireRow.Delete
        End If
    Next Fila
End Sub

Sub Elimina_Filas_Vacias()

    Dim n As Long 'nº filas
    Dim i As Long
    Dim Fila As String

    n = ActiveSheet.UsedRange.Rows.Count

    For i = n To 1 Step -1
        Fila = i & ":" & i
        If WorksheetFunction.CountA(Range(Fila)) = 0 Then
            Range("A" & i).EntireRow.Delete
        End If
    Next i
End Sub

Sub genera()
    Dim i As Long, j As Byte
    Dim f As Long, c As Integer
    Dim Zona, Vendedor, Delegacion
    Dim alea1 As Byte
    Dim alea2 As Byte
    Dim alea3 As Byte
    Dim Fila As Byte
    Dim Listado As String
    Randomize
    ActiveSheet.UsedRange.Clear
    Range(Cells(1, 1), Cells(103, 15)).Interior.ColorIndex = 34
    Cells(3, 2) = "Fecha"
    Cells(3, 3) = "Zona"
    Cells(3, 5) = "Vendedor"
    Cells(3, 8) = "Importe"
    Cells(3, 12) = "Comisión"
    Cells(3, 13) = "Delegación"
    Cells(3, 15) = "Km"
    Zona = Array("Norte", "Sur", "Este", "Oeste", "Centro")
    Vendedor = Array("Ruiz", "Lopez", "Martin", "Plaza", "García")
    Delegacion = Array("Madrid", "Barcelona", "Sevilla", "Valencia", "Bilbao", "La Coruña")
    Rows("3:3").HorizontalAlignment = xlCenter
    Rows("3:3").Font.Bold = True
    For i = 1 To 100
        Cells(i + 3, 2) = Date + i - 2
        alea1 = Int(Rnd() * 5) + 1
        Cells(i + 3, 3) = Zona(alea1 - 1)
        alea2 = Int(Rnd() * 5) + 1
        Cells(i + 3, 5) = Vendedor(alea2 - 1)
        Cells(i + 3, 8) = Int(Rnd() * 100000) - 1
        Cells(i + 3, 12) = Cells(i + 3, 8) * (0.04 + (Int(Rnd() * 2) + 1) * 0.03)
        alea3 = Int(Rnd() * 6) + 1
        Cells(i + 3, 13) = Delegacion(alea3 - 1)
        Cells(i + 3, 15) = (Int(Rnd() * 10000) - 1) / 100
    Next
    Formatea
    Range("6:6,9:10,14:14,17:18,25:25,30:31,42:44,57:59,66:66,79:81,94:94,101:102").Insert
    Range("A1") = "Generar Informe"
    Range("A2") = "Eliminar Filas Vacias"
    Range("A3") = "Eliminar Columnas Vacias"
End Sub

Sub Formatea()

    Range("B3:O3").Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
  
    Columns("H:H").Select
    Selection.NumberFormat = "#,##0"
    
    Columns("L:L").Select
    Selection.NumberFormat = "#,##0.00"

    Columns("O:O").Select
    Selection.NumberFormat = "#,##0.00"

    Range("A1:A3").Font.Bold = False
    Range("A1:A3").HorizontalAlignment = xlLeft
    Range("A1").Select
End Sub

Sub AddRowAndDeleteRow()
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Hora ent teorica"
    'rellenamos columna: Si hora<10:00
    
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
End Sub
Sub CrearColumnaYReemplazarHoras()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Hora As Date
    
    Const HoraLimiteInferior As Date = #9:01:00 AM#
    Const HoraLimiteSuperior As Date = #9:59:00 AM#
    Const ValorSustituido As String = "10:00"
    Const EncabezadoNuevaCol As String = "Hora ent teorica"
    
    ' Definir la hoja de trabajo en la que deseas realizar la operación
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Cambia "NombreDeTuHoja" al nombre de tu hoja
    
    Application.ScreenUpdating = False
    
    ' Encontrar la última fila en la columna "Hora ent"
    LastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    With ws
        .Columns("E:E").Insert Shift:=xlToRight
        .Cells(1, "E").Value = EncabezadoNuevaCol
        .Range("E2:E" & LastRow).Value = .Range("D2:D" & LastRow).Value
    End With
    
    For i = 2 To LastRow ' Comenzamos desde la fila 2 asumiendo que la fila 1 contiene encabezados
        ' Obtener el valor de texto de la celda en columna "E"
        Dim ValorTexto As String
        ValorTexto = ws.Cells(i, "E").Text
        
        ' Convertir el valor de texto a una hora
        If IsDate(ValorTexto) Then
            Hora = TimeValue(ValorTexto)
            
            ' Verificar si la hora está dentro del rango definido
            If Hora >= HoraLimiteInferior And Hora <= HoraLimiteSuperior Then
                ' Si está en ese rango, establecer el valor definido
                ws.Cells(i, "E").Value = ValorSustituido

            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub

