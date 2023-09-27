

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
    Const HoraTarde2LimiteInferior As Date = #6:01:59 PM#
    Const HoraTarde2LimiteSuperior As Date = #6:59:59 PM#
    Const HoraTarde3LimiteInferior As Date = #7:01:59 PM#
    Const HoraTarde3LimiteSuperior As Date = #7:59:59 PM#
    Const HoraNocheLimiteInferior As Date = #11:01:59 PM#
    Const HoraNocheLimiteSuperior As Date = #11:59:59 PM#
    Const MaxHorasExtras As String = "12:00:00"
    Dim MaxHorasSemanales As Double

    

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
        newDateTimeTarde2 = "19:00"
        newDateTimeTarde3 = "20:00"
        newDateTimeNoche = "00:00"
        
        Application.ScreenUpdating = False

        ' Set ws = ThisWorkbook.Sheets("Hoja1")
        Set ws = ActiveSheet
        
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

        Dim LastRowH As Long
        
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
            If horaEntTeorica >= HoraTarde2LimiteInferior And horaEntTeorica <= HoraTarde2LimiteSuperior Then
                ' Si está en ese rango, establecer el valor definido
                ws.Cells(i, "D").Value = newDateTimeTarde2
            End If
             If horaEntTeorica >= HoraTarde3LimiteInferior And horaEntTeorica <= HoraTarde3LimiteSuperior Then
                ' Si está en ese rango, establecer el valor definido
                ws.Cells(i, "D").Value = newDateTimeTarde3
            End If
            If horaEntTeorica >= HoraNocheLimiteInferior And horaEntTeorica <= HoraNocheLimiteSuperior Then
                ' Si está en ese rango, establecer el valor definido
                ws.Cells(i, "D").Value = newDateTimeNoche
            End If
        Next i

    ' Calculo horas trabajadas
        ' Set ws = ThisWorkbook.Sheets("Hoja1")
        Set ws = ActiveSheet
        
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
        ' Set Hoja = ThisWorkbook.Sheets("Hoja1")
        Set Hoja = ActiveSheet
        
        UltimaFila = Hoja.Cells(Hoja.Rows.Count, "I").End(xlUp).Row
         With ws
            .Columns("J:J").Insert Shift:=xlToRight
            .Cells(1, "J").Value = "Total horas extras"
        End With
        
        Dim HoraLimite As Date
        HoraLimite = TimeValue("08:00:00") ' Cambia "08:00:00" según tu hora límite
        
        
        For i = 2 To UltimaFila
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
                    ' CeldaJ.Value = "00:00:00"
                    ' CeldaJ.Interior.Color = RGB(240, 128, 128) ' coral
                    
                    Dim HorasFaltantes As Date
                    HorasFaltantes = HoraLimite - HoraReal
                    CeldaJ.Value = "-" & Format(HorasFaltantes, "hh:mm")
                    ' CeldaJ.Font.Color = RGB(255, 0, 0) ' Rojo
                End If
                
                
            End If
        Next i
        
        ' Eliminar valores 00:00:00 Col J si la columna A contiene "Empleado :" o "Firma Empleado"
           For i = 2 To LastRow
        
            ' Convierte el valor en la celda a formato de tiempo (si es posible).
            On Error Resume Next
                Dim valorComoTiempo As Date
                valorComoTiempo = CDate(ws.Cells(i, 9).Value)
            On Error GoTo 0
            
            ' Verifica si el valor es igual a "00:00:00" en formato de tiempo.
            ' If (valorComoTiempo = TimeValue("00:00:00")) And (InStr(1, ws.Cells(i, "A").Value, "Empleado :", vbTextCompare) > 0) Then
                ' ws.Cells(i, "J").Value = ""
            ' End If
            
            If (valorComoTiempo = TimeValue("00:00:00")) And ((ws.Cells(i, 1).Value Like "*Empleado :*") Or (ws.Cells(i, 1).Value Like "*Firma Empleado*")) Then
                ws.Cells(i, "J").Value = ""
            End If
           Next i
           
        ' Añadir 2 filas debajo Firma Empleado
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
        
        ' Cambiar formato Total horas Reales Fecha + hh:ss --> hh:ss
        Set Hoja = ActiveSheet
        
         For Each Celda In Hoja.Range("I1:I" & Hoja.Cells(Hoja.Rows.Count, "I").End(xlUp).Row)
            ' Verifica si la celda tiene un valor de fecha válido.
            If IsDate(Celda.Value) Or IsNumeric(Celda.Value) Then
                ' Extrae la parte de la hora y asigna el valor a la misma celda.
                Celda.Value = Format(CDate(Celda.Value), "hh:mm:ss")
            End If
            If Celda.Value = TimeValue("00:00:00") Or Celda.Value = "0,00" Then
                Celda.Value = ""
            End If
        Next Celda
        
        ' Suma Total Horas Reales
            LastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row + 1
            
            total = 0
            
             For i = 1 To LastRow
                If ws.Cells(i, "A").Value Like "Empleado :*" Then
                   total = 0
                End If
                
                If (IsDate(ws.Cells(i, "I").Value) Or IsNumeric(ws.Cells(i, "I").Value)) And (IsEmpty(ws.Cells(i, "A").Value)) Then
                   total = ws.Cells(i, "I").Value + total
                End If
                
                If ws.Cells(i, "A").Value Like "Total Semana" Then
                    ws.Cells(i, "I").NumberFormat = "[h]:mm:ss;@"
                    ws.Cells(i, "I").Value = total
                    
                    'Dim ValorCelda As Double
                    'ValorCelda = CDbl(ws.Cells(i, "I").Value)
                    ' MsgBox "El valor total en la celda es: " & total
    
                    'If ValorCelda < MaxHorasSemanales Then
                        'ws.Cells(i, "I").Font.Color = RGB(255, 0, 0) ' Rojo
                    'End If
                    
                    ' ws.Cells(i, "J").Value = Format(CDate(ws.Cells(i, "J").Value), "hh:mm:ss")
                    'ws.Cells(i, "J").NumberFormat = "[h]:mm:ss;@"
                    total = 0
                End If
            Next i
        
        ' Suma Total Horas Extras
            LastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row + 1
                
            total = 0
                
            For i = 1 To LastRow
                If ws.Cells(i, "A").Value Like "Empleado :*" Then
                    total = 0
                End If
                    
                If (IsDate(ws.Cells(i, "J").Value) Or IsNumeric(ws.Cells(i, "J").Value)) And (IsEmpty(ws.Cells(i, "A").Value)) Then
                    total = CDbl(ws.Cells(i, "J").Value) + total
                End If
                    
                If ws.Cells(i, "A").Value Like "Total Semana" Then
                    ws.Cells(i, "J").NumberFormat = "[h]:mm:ss;@"
                    ws.Cells(i, "J").Value = total
                    ' ws.Cells(i, "J").Value = Format(CDate(ws.Cells(i, "J").Value), "hh:mm:ss")
                    total = 0
                End If
            Next i
            
        ' Total periodo Horas Reales
            LastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row + 1
            
            total = 0
                
            For i = 1 To LastRow
                If ws.Cells(i, "I").Interior.Color = RGB(238, 229, 227) Then
                    total = total + ws.Cells(i, "I").Value
                End If
                
                If ws.Cells(i, "A").Value Like "TOTAL PERIODO*" Then
                    ws.Cells(i, "I").NumberFormat = "[h]:mm:ss;@"
                    ws.Cells(i, "I").Value = total
                    ' ws.Cells(i, "J").Value = Format(CDate(ws.Cells(i, "J").Value), "hh:mm:ss")
                    total = 0
                End If
            Next i
        
        ' Total periodo Horas extras
        
        
        
        'Ocultar cols
        
    
        ' Opcional: Autoajustar el ancho de la columna J para mostrar correctamente las horas
        ' ws.Columns("A:A").AutoFit
        
        ' Autoajustar columnas
        Hoja.Cells.EntireColumn.AutoFit

        ' Autoajustar filas
        Hoja.Cells.EntireRow.AutoFit
            

    Application.ScreenUpdating = True
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
