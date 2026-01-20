Public Sub test()
    '<<<<<<<<<<<<<<<<<<<<<<<<< DODANIE NAGŁÓWKA >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    ' Szuka pierwszej wolnej kolumny w wierszu 1 i wstawia nowy nagłówek.
    ' Dla pierwszego wystąpienia operatora (kolumna D) wpisuje "Start".

    Dim ws As Worksheet
    Dim headerRow As Long
    Dim col As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim lastCol As Long
    Dim operatorValue As String
    Dim seenOperators As Object
    Dim threshold As Double
    Dim dtStart As Double
    Dim dtEnd As Double
    Dim dateValue As Variant
    Dim timeValue As Variant
    Dim dataRange As Range
    Dim dataArr As Variant
    Dim outputArr() As Variant
    Dim blockStart As Long
    Dim blockHasSMinus As Boolean
    Dim previousOperator As String
    Dim nightInfo As Object
    Dim nightKey As Variant
    Dim infoArrNight As Variant
    Dim dateOnly As Double
    Dim timeOnly As Double
    Dim startDate As Double
    Dim startRow As Long
    Dim lastRowWindow As Long
    Dim earliestTime As Double
    Dim earliestRow As Long

    ' Krok 1: pracujemy na aktywnym arkuszu.
    Set ws = ActiveSheet
    ' Krok 2: ustawiamy numer wiersza nagłówków.
    headerRow = 1
    ' Krok 3: startujemy od pierwszej kolumny.
    col = 1
    ' Krok 3.1: ustawiamy próg 7 godzin jako ułamek doby.
    threshold = 7# / 24#

    ' Krok 4: pobieramy ostatni wiersz danych w kolumnie D.
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    ' Krok 5: pobieramy ostatnią używaną kolumnę w wierszu nagłówków.
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    ' Krok 6: jeżeli nagłówki są puste, zaczynamy od pierwszej kolumny.
    If ws.Cells(headerRow, lastCol).Value = "" Then
        lastCol = 0
    End If
    col = lastCol + 1

    ' Krok 7: wczytujemy dane A:K do tablicy dla szybszego przetwarzania.
    Set dataRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, 11))
    dataArr = dataRange.Value

    ' Krok 8: przygotowujemy tablicę wynikową dla nowej kolumny.
    ReDim outputArr(1 To lastRow, 1 To 1)
    outputArr(1, 1) = "SesjaZmianowaTok"

    ' Krok 9: tworzymy słownik do zapamiętania operatorów już oznaczonych.
    Set seenOperators = CreateObject("Scripting.Dictionary")

    ' Krok 9.1: przygotowujemy śledzenie bloków operatorów bez S-.
    blockStart = headerRow + 1
    blockHasSMinus = False
    previousOperator = Trim(CStr(dataArr(blockStart, 4)))

    ' Krok 10: przechodzimy po wierszach danych i oznaczamy pierwsze wystąpienia operatorów.
    For rowIndex = headerRow + 1 To lastRow
        ' Krok 11: pobieramy i czyścimy wartość operatora z kolumny D.
        operatorValue = Trim(CStr(dataArr(rowIndex, 4)))
        ' Krok 11.1: aktualizujemy informację o S- w bieżącym bloku operatora.
        If CStr(dataArr(rowIndex, 11)) = "S-" Then
            blockHasSMinus = True
        End If
        If operatorValue <> "" Then
            If Not seenOperators.Exists(operatorValue) Then
                ' Krok 12: dla pierwszego wystąpienia operatora wpisujemy "Start".
                outputArr(rowIndex, 1) = "Start"
                seenOperators.Add operatorValue, True
            End If
        End If

        ' Krok 12.1: jeśli operator się zmienia, a blok nie ma S-, oznaczamy "Koniec" na końcu bloku.
        If operatorValue <> previousOperator And previousOperator <> "" Then
            If Not blockHasSMinus Then
                outputArr(rowIndex - 1, 1) = "Koniec"
            End If
            blockStart = rowIndex
            blockHasSMinus = (CStr(dataArr(rowIndex, 11)) = "S-")
            previousOperator = operatorValue
        End If
    Next rowIndex

    ' Krok 12.2: domknięcie ostatniego bloku operatora bez S-.
    If previousOperator <> "" And Not blockHasSMinus Then
        outputArr(lastRow, 1) = "Koniec"
    End If

    ' Krok 13: sprawdzamy zakończenia sesji na podstawie wzorca S- w kolumnie K.
    For rowIndex = headerRow + 1 To lastRow - 3
        ' Krok 14: szukamy układu S- w bieżącym wierszu i 3 wiersze niżej.
        If CStr(dataArr(rowIndex, 11)) = "S-" And _
           CStr(dataArr(rowIndex + 3, 11)) = "S-" Then

            ' Krok 15: pobieramy datę i godzinę z pierwszego S- (wiersz bieżący).
            dateValue = dataArr(rowIndex, 1)
            timeValue = dataArr(rowIndex, 2)
            dtStart = 0#
            If IsNumeric(dateValue) Then
                dtStart = CDbl(dateValue)
            ElseIf IsDate(dateValue) Then
                dtStart = CDbl(CDate(dateValue))
            End If
            If IsNumeric(timeValue) Then
                dtStart = dtStart + CDbl(timeValue)
            ElseIf IsDate(timeValue) Then
                dtStart = dtStart + CDbl(CDate(timeValue))
            End If

            ' Krok 16: pobieramy datę i godzinę z kolejnego S- (wiersz + 3).
            dateValue = dataArr(rowIndex + 3, 1)
            timeValue = dataArr(rowIndex + 3, 2)
            dtEnd = 0#
            If IsNumeric(dateValue) Then
                dtEnd = CDbl(dateValue)
            ElseIf IsDate(dateValue) Then
                dtEnd = CDbl(CDate(dateValue))
            End If
            If IsNumeric(timeValue) Then
                dtEnd = dtEnd + CDbl(timeValue)
            ElseIf IsDate(timeValue) Then
                dtEnd = dtEnd + CDbl(CDate(timeValue))
            End If

            ' Krok 17: jeśli różnica przekracza 7 godzin, oznaczamy "Koniec" na pierwszym S-.
            If dtEnd - dtStart > threshold Then
                outputArr(rowIndex, 1) = "Koniec"
            End If
        End If
    Next rowIndex

    ' Krok 18: dodatkowa reguła nocnej zmiany (start po 16:00,
    ' koniec do 08:00 dnia następnego dla tego samego operatora).
    Set nightInfo = CreateObject("Scripting.Dictionary")

    ' Krok 18.1: wykrywamy start po 16:00 i szukamy ostatniego logu do 08:00 dnia następnego.
    For rowIndex = headerRow + 1 To lastRow
        operatorValue = Trim(CStr(dataArr(rowIndex, 4)))
        If operatorValue <> "" Then
            dateValue = dataArr(rowIndex, 1)
            timeValue = dataArr(rowIndex, 2)

            dateOnly = 0#
            timeOnly = 0#
            If IsNumeric(dateValue) Then
                dateOnly = Int(CDbl(dateValue))
            ElseIf IsDate(dateValue) Then
                dateOnly = Int(CDbl(CDate(dateValue)))
            End If
            If IsNumeric(timeValue) Then
                timeOnly = CDbl(timeValue)
            ElseIf IsDate(timeValue) Then
                timeOnly = CDbl(CDate(timeValue)) - Int(CDbl(CDate(timeValue)))
            End If

            If Not nightInfo.Exists(operatorValue) Then
                ' [0]=state (0-brak startu,1-ma start), [1]=startDate, [2]=startRow, [3]=lastRowWindow,
                ' [4]=earliestTime, [5]=earliestRow
                infoArrNight = Array(0, 0#, 0, 0, timeOnly, rowIndex)
                nightInfo.Add operatorValue, infoArrNight
            Else
                infoArrNight = nightInfo(operatorValue)
                If timeOnly < infoArrNight(4) Then
                    infoArrNight(4) = timeOnly
                    infoArrNight(5) = rowIndex
                End If
            End If

            If infoArrNight(0) = 0 Then
                ' Start nocnej zmiany: pierwszy log po 16:00 w danym dniu.
                If timeOnly >= (16# / 24#) Then
                    infoArrNight(0) = 1
                    infoArrNight(1) = dateOnly
                    infoArrNight(2) = rowIndex
                    infoArrNight(3) = 0
                    nightInfo(operatorValue) = infoArrNight
                End If
            Else
                startDate = infoArrNight(1)
                If dateOnly = startDate + 1 And timeOnly <= (8# / 24#) Then
                    infoArrNight(3) = rowIndex
                    nightInfo(operatorValue) = infoArrNight
                End If
            End If
        End If
    Next rowIndex

    ' Krok 18.2: zapisujemy Start/Koniec dla operatorów spełniających regułę nocną.
    For Each nightKey In nightInfo.Keys
        infoArrNight = nightInfo(nightKey)
        startDate = infoArrNight(1)
        startRow = infoArrNight(2)
        lastRowWindow = infoArrNight(3)
        earliestTime = infoArrNight(4)
        earliestRow = infoArrNight(5)

        If lastRowWindow <> 0 And startRow <> 0 Then
            Dim hasKoniec As Boolean
            hasKoniec = False
            For rowIndex = headerRow + 1 To lastRow
                If Trim(CStr(dataArr(rowIndex, 4))) = nightKey Then
                    If outputArr(rowIndex, 1) = "Koniec" Then
                        hasKoniec = True
                        Exit For
                    End If
                End If
            Next rowIndex
            If hasKoniec Then
                GoTo NextNightKey
            End If
            ' Jeśli operator zaczyna między 05:00 a 08:00, zostawiamy wcześniejszy Start.
            If earliestTime >= (5# / 24#) And earliestTime <= (8# / 24#) Then
                outputArr(earliestRow, 1) = "Start"
            Else
            ' Czyścimy wcześniejsze "Start" dla tego operatora, aby nie było duplikatów.
            For rowIndex = headerRow + 1 To lastRow
                If Trim(CStr(dataArr(rowIndex, 4))) = nightKey Then
                    If outputArr(rowIndex, 1) = "Start" And rowIndex <> startRow Then
                        outputArr(rowIndex, 1) = ""
                    End If
                End If
            Next rowIndex
            outputArr(startRow, 1) = "Start"
            End If
            outputArr(lastRowWindow, 1) = "Koniec"
        End If
NextNightKey:
    Next nightKey

    ' Krok 19: uzupełniamy "W toku" pomiędzy Start i Koniec dla tego samego operatora.
    Dim startRows As Object
    Dim endRows As Object
    Dim startLookup As Variant
    Dim endLookup As Variant
    Dim colDobowa As Long
    Dim colDobowaL As String
    Dim arrDobowa() As Variant
    Dim minR As Double, maxR As Double
    Dim minP As Double, maxP As Double
    Dim maxNA As Double, minNB As Double
    Dim sesjaDobowa As String
    Dim dictStart As Object
    Dim dictDates As Object
    Dim dictInfo As Object
    Dim infoArrDobowa As Variant
    Dim minDate As Double
    Dim dateKey As String
    Dim startKey As String
    Dim lastTime As Double
    Dim switchedNB As Boolean

    Set startRows = CreateObject("Scripting.Dictionary")
    Set endRows = CreateObject("Scripting.Dictionary")

    For rowIndex = headerRow + 1 To lastRow
        operatorValue = Trim(CStr(dataArr(rowIndex, 4)))
        If operatorValue <> "" Then
            If outputArr(rowIndex, 1) = "Start" Then
                startRows(operatorValue) = rowIndex
            ElseIf outputArr(rowIndex, 1) = "Koniec" Then
                endRows(operatorValue) = rowIndex
            End If
        End If
    Next rowIndex

    For Each startLookup In startRows.Keys
        If endRows.Exists(startLookup) Then
            startRow = startRows(startLookup)
            lastRowWindow = endRows(startLookup)
            If lastRowWindow > startRow Then
                For rowIndex = startRow + 1 To lastRowWindow - 1
                    If Trim(CStr(dataArr(rowIndex, 4))) = startLookup Then
                        If outputArr(rowIndex, 1) = "" Then
                            outputArr(rowIndex, 1) = "W toku"
                        End If
                    End If
                Next rowIndex
            End If
        End If
    Next startLookup

    ' Krok 20: dodajemy kolumnę SesjaDobowaZakres według czasu w kolumnie Godzina.
    colDobowa = col + 1
    ws.Cells(headerRow, colDobowa).Value = "SesjaDobowaZakres"
    colDobowaL = Split(ws.Cells(headerRow, colDobowa).Address(False, False), CStr(headerRow))(0)

    minR = TimeSerial(4, 0, 0)
    maxR = TimeSerial(9, 0, 0)
    minP = TimeSerial(9, 0, 0)
    maxP = TimeSerial(16, 0, 0)
    maxNA = TimeSerial(4, 0, 0)
    minNB = TimeSerial(16, 0, 0)

    ReDim arrDobowa(1 To lastRow, 1 To 1)
    arrDobowa(1, 1) = "SesjaDobowaZakres"

    Set dictDates = CreateObject("Scripting.Dictionary")
    minDate = 0#

    For rowIndex = headerRow + 1 To lastRow
        dateValue = dataArr(rowIndex, 1)
        dateOnly = 0#
        If IsNumeric(dateValue) Then
            dateOnly = Int(CDbl(dateValue))
        ElseIf IsDate(dateValue) Then
            dateOnly = Int(CDbl(CDate(dateValue)))
        End If

        If dateOnly <> 0# Then
            dateKey = CStr(dateOnly)
            If Not dictDates.Exists(dateKey) Then
                dictDates.Add dateKey, True
                If minDate = 0# Or dateOnly < minDate Then
                    minDate = dateOnly
                End If
            End If
        End If
    Next rowIndex

    Set dictStart = CreateObject("Scripting.Dictionary")
    Set dictInfo = CreateObject("Scripting.Dictionary")

    For rowIndex = headerRow + 1 To lastRow
        sesjaDobowa = ""
        operatorValue = Trim$(CStr(dataArr(rowIndex, 4)))
        dateValue = dataArr(rowIndex, 1)
        dateOnly = 0#
        If IsNumeric(dateValue) Then
            dateOnly = Int(CDbl(dateValue))
        ElseIf IsDate(dateValue) Then
            dateOnly = Int(CDbl(CDate(dateValue)))
        End If

        startKey = CStr(dateOnly) & "|" & operatorValue
        If operatorValue <> "" And dateOnly <> 0# And dateOnly = minDate Then
            If CStr(dataArr(rowIndex, 11)) = "S+" Then
                arrDobowa(rowIndex, 1) = ""
                GoTo NextDobowaRow
            End If
            If IsNumeric(dataArr(rowIndex, 2)) Then
                t = CDbl(dataArr(rowIndex, 2))
                If Not dictStart.Exists(startKey) Then
                    If t <= maxNA Then
                        sesjaDobowa = "N-A"
                    ElseIf t > minR And t <= maxR Then
                        sesjaDobowa = "R"
                    ElseIf t > minP And t <= maxP Then
                        sesjaDobowa = "P"
                    Else
                        sesjaDobowa = "N-B"
                    End If
                    dictStart.Add startKey, True
                    dictInfo.Add startKey, Array(sesjaDobowa, t, False)
                Else
                    infoArrDobowa = dictInfo(startKey)
                    sesjaDobowa = CStr(infoArrDobowa(0))
                    lastTime = CDbl(infoArrDobowa(1))
                    switchedNB = CBool(infoArrDobowa(2))

                    If Not switchedNB And (t - lastTime) > (8# / 24#) Then
                        sesjaDobowa = "N-B"
                        switchedNB = True
                    End If

                    infoArrDobowa(0) = sesjaDobowa
                    infoArrDobowa(1) = t
                    infoArrDobowa(2) = switchedNB
                    dictInfo(startKey) = infoArrDobowa
                End If
                arrDobowa(rowIndex, 1) = sesjaDobowa
            End If
        End If
NextDobowaRow:
    Next rowIndex

    ' Krok 21: zapisujemy wyniki do nowych kolumn w arkuszu.
    ws.Cells(headerRow, col).Resize(lastRow, 1).Value = outputArr
    ws.Cells(headerRow, colDobowa).Resize(lastRow, 1).Value = arrDobowa
End Sub
