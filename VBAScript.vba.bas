Sub SockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim ticker As Variant ' Declare as Variant to handle any value
    Dim outputRow As Long
    Dim tickerCollection As New Collection
    Dim sheetName As String
    Dim j As Long
    Dim firstRow As Long, lastRowOfTicker As Long
    Dim openingPrice As Double, closingPrice As Double
    Dim lastUniqueRow As Long
    Dim quarterlyChange As Double, percentChange As Double
    Dim totalVolume As Double
    Dim k As Long
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    Dim tickerMaxIncrease As String, tickerMaxDecrease As String, tickerMaxVolume As String

    ' Loop through the required sheets (Q1, Q2, Q3, Q4)
    For Each ws In ThisWorkbook.Worksheets
        sheetName = ws.Name

        ' Process only if the sheet is Q1, Q2, Q3, or Q4
        If sheetName Like "Q*" Then
            ' Find the last row with data in column A
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            ' Clear the collection for each new sheet
            Set tickerCollection = New Collection

            ' Loop through each row in column A to gather unique tickers
            On Error Resume Next ' Ignore errors when adding duplicate items
            For i = 2 To lastRow ' Start at row 2 to skip header
                ticker = Trim(ws.Cells(i, 1).Value) ' Ensure no leading/trailing spaces

                ' Add ticker to the collection if it's not already there
                If ticker <> "" Then tickerCollection.Add ticker, CStr(ticker)
            Next i
            On Error GoTo 0 ' Reset error handling

            ' Output the unique tickers in column I, starting from row 2
            ws.Cells(1, 9).Value = "tickers" ' Header in I1
            outputRow = 2

            ' Write the unique tickers to column I
            For Each ticker In tickerCollection
                ws.Cells(outputRow, 9).Value = ticker
                outputRow = outputRow + 1
            Next ticker
        End If
    Next ws

    ' Loop through each relevant sheet (Q1, Q2, Q3, Q4)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            
            ' Find the last row with data in column I (unique tickers)
            lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

            ' Add header to column J
            ws.Cells(1, 10).Value = "Quarterly Change"

            ' Loop through each unique ticker in column I
            For j = 2 To lastRow
                ticker = ws.Cells(j, 9).Value ' Get the unique ticker

                ' Find the first and last occurrence of the ticker in column A
                firstRow = ws.Columns(1).Find(What:=ticker, LookAt:=xlWhole).Row
                lastRowOfTicker = ws.Columns(1).Find(What:=ticker, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row

                ' Get the opening price from the first occurrence
                openingPrice = ws.Cells(firstRow, 3).Value

                ' Get the closing price from the last occurrence
                closingPrice = ws.Cells(lastRowOfTicker, 6).Value

                ' Calculate the change and write it to column J
                ws.Cells(j, 10).Value = closingPrice - openingPrice
            Next j
        End If
    Next ws

    ' Loop through all relevant sheets (Q1, Q2, Q3, Q4)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q*" Then
            ' Find the last row with data in column J
            lastRow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row

            ' Set the range for conditional formatting
            Dim rng As Range
            Set rng = ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10)) ' From J2 to last used row

            ' Clear previous conditional formatting
            rng.FormatConditions.Delete

            ' Apply green formatting for values > 0
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                .Interior.Color = RGB(144, 238, 144) ' Light Green
            End With

            ' Apply red formatting for values < 0
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                .Interior.Color = RGB(255, 99, 71) ' Light Red
            End With
        End If
    Next ws

    ' Loop a través de todas las hojas relevantes (Q1, Q2, Q3, Q4)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q*" Then
            ' Encuentra la última fila con datos en la columna I (unique tickers)
            lastUniqueRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

            ' Añade el encabezado "Percent Change" en K1
            ws.Cells(1, 11).Value = "Percent Change"

            ' Loop para calcular el cambio porcentual para cada ticker único
            For i = 2 To lastUniqueRow ' Comienza en la fila 2 para omitir el encabezado
                ticker = ws.Cells(i, 9).Value ' Obtener el ticker único

                ' Encuentra la primera ocurrencia del ticker en la columna A
                firstRow = ws.Columns(1).Find(What:=ticker, LookAt:=xlWhole).Row

                ' Obtén el precio de apertura de la primera ocurrencia (columna C)
                openingPrice = ws.Cells(firstRow, 3).Value

                ' Obtén el Quarterly Change de la columna J
                quarterlyChange = ws.Cells(i, 10).Value

                ' Verifica que el precio de apertura no sea 0 para evitar errores
                If openingPrice <> 0 Then
                    ' Calcula el cambio porcentual
                    percentChange = (quarterlyChange / openingPrice)
                Else
                    ' Si el precio de apertura es 0, asigna 0 al cambio porcentual
                    percentChange = 0
                End If

                ' Coloca el resultado en la columna K como porcentaje
                ws.Cells(i, 11).Value = percentChange
            Next i

            ' Formatea la columna K como porcentaje con dos decimales
            ws.Columns(11).NumberFormat = "0.00%"

            ' Ajusta automáticamente las columnas para que se vean bien
            ws.Columns.AutoFit
        End If
    Next ws
   
    ' Loop a través de todas las hojas relevantes (Q1, Q2, Q3, Q4)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q*" Then
            ' Encuentra la última fila con datos en la columna I (unique tickers)
            lastUniqueRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

            ' Añade el encabezado "Total Stock Volume" en L1
            ws.Cells(1, 12).Value = "Total Stock Volume"

            ' Loop para calcular el volumen total para cada ticker único
            For i = 2 To lastUniqueRow ' Comienza en la fila 2 para omitir el encabezado
                ticker = ws.Cells(i, 9).Value ' Obtener el ticker único

                ' Encuentra la primera ocurrencia del ticker en la columna A
                firstRow = ws.Columns(1).Find(What:=ticker, LookAt:=xlWhole).Row

                ' Encuentra la última ocurrencia del ticker en la columna A
                lastRow = ws.Columns(1).Find(What:=ticker, LookAt:=xlWhole, _
                                             SearchDirection:=xlPrevious).Row

                ' Inicializa el total de volumen
                totalVolume = 0

                ' Loop para sumar el volumen de todas las ocurrencias del ticker en la columna G
                For k = firstRow To lastRow
                    totalVolume = totalVolume + ws.Cells(k, 7).Value ' Suma valores de columna G
                Next k

                ' Coloca el total de volumen en la columna L
                ws.Cells(i, 12).Value = totalVolume
            Next i

            ' Ajusta automáticamente las columnas para que se vean bien
            ws.Columns.AutoFit
        End If
    Next ws

    ' Loop a través de todas las hojas relevantes (Q1, Q2, Q3, Q4)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q*" Then
            ' Encuentra la última fila con datos en la columna I (unique tickers)
            lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

            ' Inicializa variables para el máximo y mínimo
            maxIncrease = -1E+308 ' Mínimo posible para buscar el máximo
            maxDecrease = 1E+308  ' Máximo posible para buscar el mínimo
            maxVolume = 0

            ' Loop para encontrar los mayores y menores valores
            For i = 2 To lastRow ' Comienza en la fila 2 para omitir el encabezado
                ' Verifica si el valor actual es mayor al maxIncrease
                If ws.Cells(i, 11).Value > maxIncrease Then
                    maxIncrease = ws.Cells(i, 11).Value
                    tickerMaxIncrease = ws.Cells(i, 9).Value ' Guarda el ticker correspondiente
                End If

                ' Verifica si el valor actual es menor al maxDecrease
                If ws.Cells(i, 11).Value < maxDecrease Then
                    maxDecrease = ws.Cells(i, 11).Value
                    tickerMaxDecrease = ws.Cells(i, 9).Value ' Guarda el ticker correspondiente
                End If

                ' Verifica si el valor actual es mayor al maxVolume
                If ws.Cells(i, 12).Value > maxVolume Then
                    maxVolume = ws.Cells(i, 12).Value
                    tickerMaxVolume = ws.Cells(i, 9).Value ' Guarda el ticker correspondiente
                End If
            Next i

            ' Añade encabezados en P1 y Q1
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"

            ' Añade los resultados en O2, O3 y O4
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"

            ' Coloca los resultados en las columnas P y Q
            ws.Cells(2, 16).Value = tickerMaxIncrease
            ws.Cells(2, 17).Value = maxIncrease
            ws.Cells(3, 16).Value = tickerMaxDecrease
            ws.Cells(3, 17).Value = maxDecrease
            ws.Cells(4, 16).Value = tickerMaxVolume
            ws.Cells(4, 17).Value = maxVolume

            ' Formatea las celdas de Q2 y Q3 como porcentaje con dos decimales
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(3, 17).NumberFormat = "0.00%"

            ' Ajusta automáticamente las columnas para que se vean bien
            ws.Columns.AutoFit
        End If
    Next ws

    ' Informa al usuario que el proceso ha finalizado
    MsgBox "All operations have been completed.", vbInformation
End Sub
