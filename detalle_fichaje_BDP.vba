Option Explicit

Sub FormatDetailTimeSheet()

    Dim numCols As Integer 'nº columnas
    Dim iCols As Integer
    Dim col As String

    Dim numRows As Long 'nº filas
    Dim iRows As Long
    Dim Fila As String

    Dim ws As Worksheet
    Dim newDateTimeMañana As String
    Dim newDateTimeTarde As String
    Dim newDateTimeNoche As String
    Dim iSep As Long
    Dim i As Long

    Dim LastRow As Long
    Dim Hoja As Worksheet
    Dim Celda As Range
    Dim UltimaFila As Long
    
    Dim UltimaFilaI As Long
    Dim CeldaI As Range

    Const HoraMañanaLimiteInferior As Date = #9:01:00 AM#
    Const HoraMañanaLimiteSuperior As Date = #9:59:00 AM#
    Const HoraTardeLimiteInferior As Date = #4:01:59 PM#
    Const HoraTardeLimiteSuperior As Date = #4:59:59 PM#
    Const HoraNocheLimiteInferior As Date = #11:01:59 PM#
    Const HoraNocheLimiteSuperior As Date = #11:59:59 PM#
    Const MaxHorasExtras As String = "12:00:00"

    Application.ScreenUpdating = False

    ' Eliminar encabezado
    Range("A1:N5").Select
    Selection.Delete Shift:=xlUp


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


    ' Insertar columna "Hora Ent Teorica"
        newDateTimeMañana = "10:00"
        newDateTimeTarde = "17:00"
        newDateTimeNoche = "00:00"
        
        Application.ScreenUpdating = False

        Set ws = ThisWorkbook.Sheets("Sheet1")

        ' Encontrar la última fila en la columna "Hora ent"
        LastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

        With ws
            .Columns("D:D").Insert Shift:=xlToRight
            .Cells(1, "D").Value = "Hora Ent Teorica"
        End With
        
        ' Recorre las filas y separa las fechas y horas
        For iSep = 2 To LastRow
            Cells(iSep, "D").Value = Format(Cells(iSep, 3), "hh:mm:ss")
        Next iSep

        
        LastRowH = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        
        For i = 2 To LastRowH
            Dim ValorTextoD As Date
            ValorTextoD = ws.Cells(i, "D").Value
            horaEntTeorica = CDate(ValorTextoD)
            
            If horaEntTeorica >= HoraMañanaLimiteInferior And horaEntTeorica <= HoraMañanaLimiteSuperior Then
                ' Si está en ese rango, establecer el valor definido
                ws.Cells(i, "D").Value = newDateTimeMañana
            End If
            If horaEntTeorica >= HoraTardeLimiteInferior And horaEntTeorica <= HoraTardeLimiteSuperior Then
                ' Si está en ese rango, establecer el valor definido
                ws.Cells(i, "D").Value = newDateTimeTarde
            End If
            If horaEntTeorica >= HoraNocheLimiteInferior And horaEntTeorica <= HoraNocheLimiteSuperior Then
                ' Si está en ese rango, establecer el valor definido
                ws.Cells(i, "D").Value = newDateTimeNoche
            End If
        Next i

    ' Calculo horas trabajadas
        Set ws = ThisWorkbook.Sheets("Sheet1")

        ' Encontrar la última fila en la columna "Hora ent"
        LastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

        With ws
            .Columns("I:I").Insert Shift:=xlToRight
            .Cells(1, "I").Value = "Total horas Reales"
        End With

        
        ' Itera a través de las filas con datos
        For i = 2 To LastRow
            If Not IsEmpty(Cells(i, "D").Value) Then
                ' Calcula la diferencia entre las horas
                Dim HorasTrabajadas As Double
                HorasTrabajadas = Cells(i, "F").Value - Cells(i, "D").Value
                
                ' Coloca el resultado en la columna "Total horas reales"
                Cells(i, "I").Value = HorasTrabajadas
                
                ' Cambia el formato de la celda a hora para que se vea bien
                Cells(i, "I").NumberFormat = "hh:mm"
            End If
        Next i
        
        
        ' Encuentra la última fila en la columna H (columna adyacente a la izquierda de la columna I)
        UltimaFila = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        
        ' Recorre las celdas en la columna I y colorea las celdas vacías de gris si la celda en H tiene datos
        For Each Celda In ws.Range("I1:I" & UltimaFila)
            If Celda.Offset(0, -1).Value <> "" And IsEmpty(Celda.Value) Then
                Celda.Interior.Color = RGB(238, 229, 227) ' rosa
            End If
        Next Celda
        
        ' Horas extras
        
    
        ' Opcional: Autoajustar el ancho de la columna J para mostrar correctamente las horas
        ws.Columns("A:A").AutoFit
        ws.Columns("B:B").AutoFit
        ws.Columns("C:C").AutoFit
        ws.Columns("D:D").AutoFit
        ws.Columns("E:E").AutoFit
        ws.Columns("F:F").AutoFit
        ws.Columns("G:G").AutoFit
        ws.Columns("H:H").AutoFit
        ws.Columns("I:I").AutoFit
        ws.Columns("J:J").AutoFit
        

    Application.ScreenUpdating = True
End Sub



Sub CalcularHorasExtras()
    Dim Hoja As Worksheet
    Set Hoja = ThisWorkbook.Sheets("Hoja2") ' Cambia "Nombre de tu hoja" con el nombre de tu hoja
    
    Dim UltimaFila As Long
    UltimaFila = Hoja.Cells(Hoja.Rows.Count, "I").End(xlUp).Row
    
    Dim HoraLimite As Date
    HoraLimite = TimeValue("08:00:00") ' Cambia "08:00:00" según tu hora límite
    
    Dim CeldaI As Range
    Dim CeldaJ As Range
    
    For i = 2 To UltimaFila ' Empezando desde la segunda fila (asumiendo que la fila 1 tiene encabezados)
        Set CeldaI = Hoja.Cells(i, "I")
        Set CeldaJ = Hoja.Cells(i, "J")
        
        If IsNumeric(CeldaI.Value) Then
            Dim HoraReal As Date
            HoraReal = TimeSerial(Hour(CeldaI.Value), Minute(CeldaI.Value), Second(CeldaI.Value))
            
            If HoraReal > HoraLimite Then
                Dim HorasExtras As Date
                HorasExtras = HoraReal - HoraLimite
                CeldaJ.Value = Format(HorasExtras, "hh:mm:ss")
            Else
                CeldaJ.Value = "00:00:00"
            End If
        End If
    Next i
End Sub
