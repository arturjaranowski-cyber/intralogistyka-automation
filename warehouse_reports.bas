Public Sub PrzetworzWielePlikowTekstowych()

    '<<<<<<<<<<<<<<<<<<<<<<<<< IMPORTOWANIE PLIKÓW DO PRZETWORZENIA >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    '...
    'Importowanie - Wczytywanie - plików tekstowych typu *txt
    Dim files As Variant, i As Long
    Dim daneSheet As Worksheet
    Dim wsTemp As Worksheet
    Dim wbTextImport As Workbook

    ' =========================================================
    ' START: Porządek arkuszy
    ' - zostawiamy WerehouseReports
    ' - kasujemy wszystkie inne arkusze
    ' - tworzymy nowe "Dane" ZA WerehouseReports
    ' =========================================================

    Dim wsWR As Worksheet
    Dim wsX As Worksheet

    Application.DisplayAlerts = False

    ' Upewnij się, że istnieje WerehouseReports (jeśli nie ma – utwórz)
    On Error Resume Next
    Set wsWR = ThisWorkbook.Worksheets("WerehouseReports")
    On Error GoTo 0

    If wsWR Is Nothing Then
        Set wsWR = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsWR.Name = "WerehouseReports"
    End If

    ' Usuń wszystkie arkusze poza WerehouseReports
    Dim idx As Long
    For idx = ThisWorkbook.Worksheets.Count To 1 Step -1
        If ThisWorkbook.Worksheets(idx).Name <> "WerehouseReports" Then
            ThisWorkbook.Worksheets(idx).Delete
        End If
    Next idx

    Application.DisplayAlerts = True

    ' Utwórz nowy arkusz "Dane" ZA WerehouseReports (zawsze nowy, czysty)
    Set daneSheet = ThisWorkbook.Worksheets.Add(After:=wsWR)
    daneSheet.Name = "Dane"
    daneSheet.Cells.Clear
    ' =========================================================
    ' KONIEC: Porządek arkuszy
    ' =========================================================

    ' Komunikat dla użytkownika
    MsgBox "Wczytaj dane do arkusza" & vbNewLine & _
           "Wybierz pliki tekstowe (*.txt) z logami WMS." & vbNewLine & vbNewLine & _
           "Każdy plik zostanie przetworzony według przygotowanej instrukcji.", _
           vbInformation, "Wydruk: Zapis z logami"

    ' Wybór wielu plików tekstowych
    files = Application.GetOpenFilename("Pliki tekstowe (*.txt), *.txt", MultiSelect:=True, _
                                         Title:="Wybierz pliki tekstowe do wczytania")

    If VarType(files) = vbBoolean Then
        If files = False Then
            MsgBox "Nie wybrano plików", vbCritical, "Błąd"
            Exit Sub
        End If
    End If

    For i = LBound(files) To UBound(files)
        ' Otwórz plik tekstowy w nowym skoroszycie z ustawionym kodowaniem (852)
        Workbooks.OpenText fileName:=files(i), Origin:=852, StartRow:=1, _
            DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
            Tab:=False, Semicolon:=True, Comma:=False, Space:=False, Other:=False, _
            FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
        Set wbTextImport = ActiveWorkbook

        ' Utwórz tymczasowy arkusz w bieżącym skoroszycie
        Set wsTemp = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTemp.Name = "Temp_" & Format(Now, "yyyymmdd_hhmmss") & "_" & i

        ' Skopiuj dane z importowanego pliku do tymczasowego arkusza
        wbTextImport.Sheets(1).UsedRange.Copy wsTemp.Range("A1")
        wbTextImport.Close False


        '<<<<<<<<<<<<<<<<<<<<<<<<< MODUŁ PRZETWARZAJĄCY DANE >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

        ' Usuń nagłówek (wiersze 1:8)
        wsTemp.Rows("1:8").Delete

        ' Usunięcie dwóch ostatnich wypełnionych komórek w kolumnie A
        Dim lastRow As Long
        lastRow = wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row
        wsTemp.Cells(lastRow, "A").ClearContents
        wsTemp.Cells(lastRow - 1, "A").ClearContents

        ' Wstępne usuwanie wierszy zawierających niechciane frazy
        Dim lastRowA As Long, iA As Long, jA As Long
        Dim dataA As Variant, resultA() As Variant
        Dim textA As String

        lastRowA = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
        dataA = wsTemp.Range("A1:A" & lastRowA).Value
        ReDim resultA(1 To lastRowA, 1 To 1)
        jA = 1
        For iA = 1 To lastRowA
            textA = dataA(iA, 1)
            If InStr(1, textA, "WS:", vbTextCompare) = 0 And _
               InStr(1, textA, "ODBLOKOWANIE TRASY", vbTextCompare) = 0 And _
               InStr(1, textA, "ZABLOKOWANIE TRASY", vbTextCompare) = 0 And _
               InStr(1, textA, "ZWOLNIENIE BLOKADY", vbTextCompare) = 0 Then
                resultA(jA, 1) = textA
                jA = jA + 1
            End If
        Next iA
        wsTemp.Range("A1:A" & lastRowA).ClearContents
        wsTemp.Range("A1").Resize(jA - 1, 1).Value = resultA

        ' Rozdzielenie kolumny A na kilka kolumn (A-F)
        Dim lastRowB As Long
        Dim dataB As Variant
        Dim resultA_B() As Variant, resultB_B() As Variant, resultC_B() As Variant
        Dim resultD_B() As Variant, resultE_B() As Variant, resultF_B() As Variant
        Dim iB As Long, spacePosB As Long, dotPosB As Long
        Dim secondSpacePosB As Long, thirdSpacePosB As Long, fourthSpacePosB As Long

        lastRowB = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
        dataB = wsTemp.Range("A1:A" & lastRowB).Value
        ReDim resultA_B(1 To lastRowB, 1 To 1)
        ReDim resultB_B(1 To lastRowB, 1 To 1)
        ReDim resultC_B(1 To lastRowB, 1 To 1)
        ReDim resultD_B(1 To lastRowB, 1 To 1)
        ReDim resultE_B(1 To lastRowB, 1 To 1)
        ReDim resultF_B(1 To lastRowB, 1 To 1)

        For iB = 1 To lastRowB
            dataB(iB, 1) = LTrim(dataB(iB, 1))
            spacePosB = InStr(1, dataB(iB, 1), " ")
            If spacePosB > 0 Then
                resultA_B(iB, 1) = Left(dataB(iB, 1), spacePosB - 1)
                resultB_B(iB, 1) = Mid(dataB(iB, 1), spacePosB + 1)
            Else
                resultA_B(iB, 1) = dataB(iB, 1)
                resultB_B(iB, 1) = ""
            End If

            dotPosB = InStr(1, resultB_B(iB, 1), ".")
            If dotPosB > 0 Then
                resultC_B(iB, 1) = Mid(resultB_B(iB, 1), dotPosB + 1)
                resultB_B(iB, 1) = Left(resultB_B(iB, 1), dotPosB - 1)
            Else
                resultC_B(iB, 1) = ""
            End If

            secondSpacePosB = InStr(1, resultC_B(iB, 1), " ")
            If secondSpacePosB > 0 Then
                resultD_B(iB, 1) = Mid(resultC_B(iB, 1), secondSpacePosB + 1)
                resultC_B(iB, 1) = Left(resultC_B(iB, 1), secondSpacePosB - 1)
            Else
                resultD_B(iB, 1) = resultC_B(iB, 1)
                resultC_B(iB, 1) = ""
            End If

            resultD_B(iB, 1) = LTrim(resultD_B(iB, 1))
            thirdSpacePosB = InStr(1, resultD_B(iB, 1), " ")
            If thirdSpacePosB > 0 Then
                resultE_B(iB, 1) = Mid(resultD_B(iB, 1), thirdSpacePosB + 1)
                resultD_B(iB, 1) = Left(resultD_B(iB, 1), thirdSpacePosB - 1)
            Else
                resultE_B(iB, 1) = ""
            End If

            fourthSpacePosB = InStr(1, resultE_B(iB, 1), " ")
            If fourthSpacePosB > 0 Then
                resultF_B(iB, 1) = Mid(resultE_B(iB, 1), fourthSpacePosB + 1)
                resultE_B(iB, 1) = Left(resultE_B(iB, 1), fourthSpacePosB - 1)
            Else
                resultF_B(iB, 1) = ""
            End If

            If resultE_B(iB, 1) <> "" Then
                resultD_B(iB, 1) = resultD_B(iB, 1) & " " & resultE_B(iB, 1)
            End If
        Next iB

        wsTemp.Range("A1:A" & lastRowB).Value = resultA_B
        wsTemp.Range("B1:B" & lastRowB).Value = resultB_B
        wsTemp.Range("C1:C" & lastRowB).Value = resultC_B
        wsTemp.Range("D1:D" & lastRowB).Value = resultD_B
        wsTemp.Range("E1:E" & lastRowB).Value = resultE_B
        wsTemp.Range("F1:F" & lastRowB).Value = resultF_B

        wsTemp.Columns("E:E").Delete shift:=xlToLeft

        Dim cell As Range
        For Each cell In wsTemp.Range("E1:E" & lastRowB)
            cell.Value = Trim(Application.WorksheetFunction.Substitute(cell.Value, "  ", " "))
        Next cell

        ' Wyszukiwanie fragmentów logów
        Dim searchTextC As Variant, lastRowC As Long, iC As Long
        searchTextC = Array( _
            "PRZYJĘCIE TOWARU NA PÓŁKĘ", "[SKRZ.", "DO PACZKI ZE SKRZYNKI", "DODANIE: KARTON", "PRZENIESIENIE TOWARU", _
            "MENU: BRAK", "PRZERWANIE WYB", "SPRAWDZENIE LOK", "ZWROT TOWARU", _
            "INWENT", "POBRANIE DO WYD.", "WYBRANIE", "NA BRAKI", "PRZERWANIE PAK", "PRZERWANIE PRZY", "PRZENIESIENIE NA:", _
            "LOGOWANIE - MAG", "LOGOWANIE - PAK", "LOGOWANIE - ROZ", "LOGOWANIE - SPR", "LOGOWANIE - ZAŁ", "LOGOWANIE - ZWR", _
            "PRZENIESIENIE Ź", "ROZŁOŻENIE Z LOK. ĆÄ", "ROZŁOŻENIE Z LOK. ĆI", "ROZŁOŻENIE Z LOK. ZWR", "ROZŁOŻENIE Z LOK. RE", _
            "ROZŁOŻENIE Z LOK. TOW.", "TOWARU (ZAT", "ZESKANOWANO SKRZYNKĘ", "PRZEŁADUNEK DO PACZKI", "STATUS PACZKI ZWROTÓW", _
            "ZDJĘCIE Z PÓŁKI", _
 _
            "BRAMA", "DYŻUR", "KARTONY ZBIORCZE", _
            "KIEROWNIK", "KLIENT PALETOWY", "PAKOWANIE AUT", "PALARNIA", "PALARNIA 1", "PALARNIA 2", "PORZĄDKI", "PRZERWA -", _
            "ROZKŁADANIE PACZEK", "SZKOLENIE", "TRANSPORT SKRZYNEK", "ZAMYKANIE PACZEK", "WÓZEK HALA M", _
            "BHP", "ZWROTY", "BRYGADZISTA UK", "WINDA", "DEKONSOLIDACJA", "TRANSPORT PALET" _
        )
        lastRowC = wsTemp.Cells(wsTemp.Rows.Count, 5).End(xlUp).Row
        Dim dataArrayC As Variant, resultArrayC() As Variant
        dataArrayC = wsTemp.Range("E1:E" & lastRowC).Value
        ReDim resultArrayC(1 To lastRowC, 1 To 1)
        For iC = 1 To lastRowC
            resultArrayC(iC, 1) = 0
            Dim jC As Integer
            For jC = LBound(searchTextC) To UBound(searchTextC)
                If InStr(dataArrayC(iC, 1), searchTextC(jC)) > 0 Then
                    resultArrayC(iC, 1) = searchTextC(jC)
                    Exit For
                End If
            Next jC
        Next iC
        wsTemp.Range("F1:F" & lastRowC).Value = resultArrayC

        ' Usunięcie wierszy, gdzie kolumna F = 0
        Dim lastRowD As Long, dataArrayD As Variant, resultArrayD As Variant, newRowD As Long
        lastRowD = wsTemp.Cells(wsTemp.Rows.Count, 6).End(xlUp).Row
        dataArrayD = wsTemp.Range(wsTemp.Cells(1, 1), wsTemp.Cells(lastRowD, wsTemp.Cells(1, Columns.Count).End(xlToLeft).Column)).Value
        ReDim resultArrayD(1 To UBound(dataArrayD, 1), 1 To UBound(dataArrayD, 2))
        newRowD = 0
        Dim iD As Long, jD As Long
        For iD = 1 To UBound(dataArrayD, 1)
            If dataArrayD(iD, 6) <> 0 Then
                newRowD = newRowD + 1
                For jD = 1 To UBound(dataArrayD, 2)
                    resultArrayD(newRowD, jD) = dataArrayD(iD, jD)
                Next jD
            End If
        Next iD
        If newRowD = 0 Then
            wsTemp.Cells.Clear
            MsgBox "Wszystkie wiersze zawierały wartość 0. Arkusz został wyczyszczony.", vbExclamation, "Brak danych"
            GoTo NextFile
        End If
        Application.ScreenUpdating = False
        wsTemp.Cells.Clear
        wsTemp.Range("A1").Resize(newRowD, UBound(dataArrayD, 2)).Value = resultArrayD
        Application.ScreenUpdating = True

        ' Kopiowanie przetworzonych danych z tymczasowego arkusza do arkusza "Dane"
        Dim lastRowE As Long
        If Application.WorksheetFunction.CountA(daneSheet.Range("A:A")) = 0 Then
            lastRowE = 1
        Else
            lastRowE = daneSheet.Cells(daneSheet.Rows.Count, 1).End(xlUp).Row + 1
        End If
        wsTemp.UsedRange.Copy
        daneSheet.Cells(lastRowE, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

NextFile:
        Application.DisplayAlerts = False
        wsTemp.Delete
        Application.DisplayAlerts = True
    Next i

        '<<<<<<<<<<<<<<<<<<<<<<<<< MODUŁ PRZETWARZAJĄCY - C.D. >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    '...
    ' Sortowanie – bez nagłówków, kolejność (D, A, B, C)
    Dim wsSort As Worksheet
    Dim lastRowSort As Long
    Dim rngSort As Range

    Set wsSort = daneSheet

    ' Ostatni wiersz danych w kolumnie A
    lastRowSort = wsSort.Cells(wsSort.Rows.Count, "A").End(xlUp).Row

    If lastRowSort > 1 Then

        Set rngSort = wsSort.Range("A1:F" & lastRowSort)

        With wsSort.Sort
            .SortFields.Clear

            ' 1. Kolumna D
            .SortFields.Add key:=wsSort.Range("D1:D" & lastRowSort), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            ' 2. Kolumna A
            .SortFields.Add key:=wsSort.Range("A1:A" & lastRowSort), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            ' 3. Kolumna B
            .SortFields.Add key:=wsSort.Range("B1:B" & lastRowSort), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            ' 4. Kolumna C
            .SortFields.Add key:=wsSort.Range("C1:C" & lastRowSort), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

            .SetRange rngSort
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

    End If

    '...
    ' Usuwanie zbędnych wierszy dla loga: "[SKRZ." (kolumna F) według znalezionej frazy "ID:" w kolumnie E
    Dim wsG As Worksheet
    Dim lastRowG As Long
    Dim arrG As Variant
    Dim iG As Long

    ' Ustaw aktywny arkusz
    Set wsG = ActiveSheet

    ' Znajdź ostatni wiersz w kolumnie F, gdzie szukamy ciągu "[SKRZ."
    lastRowG = wsG.Cells(wsG.Rows.Count, "F").End(xlUp).Row
    If lastRowG < 1 Then Exit Sub

    ' Pobierz dane z kolumn E i F do tablicy (kolumna E = indeks 1, kolumna F = indeks 2)
    arrG = wsG.Range("E1:F" & lastRowG).Value

    ' Iteracja od ostatniego do pierwszego wiersza – przy usuwaniu wierszy nie zaburzy to indeksowania
    For iG = UBound(arrG, 1) To 1 Step -1
        ' Sprawdź: w kolumnie F (drugi element tablicy) musi występować ciąg "[SKRZ."
        ' oraz w kolumnie E (pierwszy element tablicy) musi być fragment "ID:"
        If (InStr(1, arrG(iG, 2), "[SKRZ.", vbTextCompare) > 0) And _
           (InStr(1, arrG(iG, 1), "ID:", vbTextCompare) > 0) Then
            wsG.Rows(iG).Delete
        End If
    Next iG

    '...
    ' Usuwanie duplikatów - sprawdzanie wierszy
    Dim wsH As Worksheet
    Dim lastRowH As Long
    Dim arrH As Variant
    Dim newArrH() As Variant
    Dim keepH() As Boolean
    Dim iH As Long, jH As Long, countH As Long
    Dim numColsH As Long

    ' Wyłączenie odświeżania ekranu, obliczeń i zdarzeń dla przyspieszenia działania
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Ustaw aktywny arkusz
    Set wsH = ActiveSheet

    ' Znajdź ostatni wiersz w kolumnie B (przyjmujemy, że kolumna B zawiera dane)
    lastRowH = wsH.Cells(wsH.Rows.Count, "B").End(xlUp).Row
    If lastRowH < 2 Then GoTo CleanUp

    ' Pobierz dane z zakresu A:F do tablicy
    arrH = wsH.Range("A1:F" & lastRowH).Value
    numColsH = 6  ' Kolumny A:F

    ' Przygotuj tablicę pomocniczą, która oznaczy, które wiersze zachować
    ReDim keepH(1 To UBound(arrH, 1))
    For iH = 1 To UBound(arrH, 1)
        keepH(iH) = True
    Next iH

    ' Porównujemy kolejne wiersze – zaczynamy od wiersza 2 do przedostatniego
    For iH = 2 To UBound(arrH, 1) - 1
        If (arrH(iH, 2) = arrH(iH + 1, 2)) And _
           (arrH(iH, 3) = arrH(iH + 1, 3)) And _
           (arrH(iH, 4) = arrH(iH + 1, 4)) Then
            ' Oznaczamy wiersz (iH+1) jako duplikat
            keepH(iH + 1) = False
        End If
    Next iH

    ' Zliczamy liczbę wierszy do zachowania
    countH = 0
    For iH = 1 To UBound(arrH, 1)
        If keepH(iH) Then countH = countH + 1
    Next iH

    ' Jeżeli nie pozostaje żaden wiersz, wyczyść zakres i zakończ
    If countH = 0 Then
        wsH.Range("A1:F" & lastRowH).ClearContents
        GoTo CleanUp
    End If

    ' Przygotuj nową tablicę o rozmiarze liczby zachowywanych wierszy
    ReDim newArrH(1 To countH, 1 To numColsH)
    countH = 0
    For iH = 1 To UBound(arrH, 1)
        If keepH(iH) Then
            countH = countH + 1
            For jH = 1 To numColsH
                newArrH(countH, jH) = arrH(iH, jH)
            Next jH
        End If
    Next iH

    ' Wyczyść oryginalny zakres i wklej przetworzone dane (tylko unikalne wiersze)
    wsH.Range("A1:F" & lastRowH).ClearContents
    wsH.Range("A1").Resize(UBound(newArrH, 1), numColsH).Value = newArrH

'CleanUp:
    ' Przywróć ustawienia
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    '...
    'Usuwanie zbędnych wierszy w arkuszu według wybranych fraz ............C.D.
    Dim wsI As Worksheet
    Dim lastRowI As Long
    Dim arrI As Variant, newArrI() As Variant, finalArrI() As Variant
    Dim iI As Long, jI As Long, countI As Long
    Dim numColsI As Long

    ' Wyłączenie optymalizacyjne – przyspieszenie działania
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Ustawienie arkusza
    Set wsI = ActiveSheet

    ' Znalezienie ostatniego wiersza (przyjmujemy, że dane zaczynają się w kolumnie A)
    lastRowI = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row
    If lastRowI < 1 Then GoTo CleanUp

    ' Pobranie danych z zakresu A:F do tablicy
    arrI = wsI.Range("A1:F" & lastRowI).Value
    numColsI = 6

    ' Przydzielamy tablicę o maksymalnym rozmiarze – tyle wierszy, ile jest w danych wejściowych
    ReDim newArrI(1 To UBound(arrI, 1), 1 To numColsI)
    countI = 0

    ' Przetwarzanie – iteracja po każdym wierszu tablicy
    For iI = 1 To UBound(arrI, 1)
        ' Sprawdzamy kolumnę E (indeks 5). Używamy CStr, żeby mieć pewność, że traktujemy wartość jako tekst.
        If (InStr(1, CStr(arrI(iI, 5)), "- KONIEC", vbTextCompare) > 0) Or _
           (InStr(1, CStr(arrI(iI, 5)), "AUTOMATYCZNE WYL", vbTextCompare) > 0) Or _
           (InStr(1, CStr(arrI(iI, 5)), "RĘCZNE WYL", vbTextCompare) > 0) Then
            ' Wiersz zawiera jeden z szukanych fragmentów – pomijamy go
        Else
            countI = countI + 1
            For jI = 1 To numColsI
                newArrI(countI, jI) = arrI(iI, jI)
            Next jI
        End If
    Next iI

    ' Jeśli żaden wiersz nie został zachowany, wyczyść zakres i zakończ makro
    If countI = 0 Then
        wsI.Range("A1:F" & lastRowI).ClearContents
        GoTo CleanUp
    End If

    ' Teraz utworzymy nową tablicę o rozmiarze liczby zachowanych wierszy – finalArrI
    ReDim finalArrI(1 To countI, 1 To numColsI)
    For iI = 1 To countI
        For jI = 1 To numColsI
            finalArrI(iI, jI) = newArrI(iI, jI)
        Next jI
    Next iI

    ' Wyczyść oryginalny zakres i wklej przetworzone dane
    wsI.Range("A1:F" & lastRowI).ClearContents
    wsI.Range("A1").Resize(countI, numColsI).Value = finalArrI

CleanUp:
    ' Przywrócenie ustawień
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    '...
    'Wyszukiwanie fragmentu (**) i wypisz w kolumnie F frazę "Komputer"
    Dim wsJ As Worksheet
    Dim lastRowJ As Long
    Dim arrJ As Variant
    Dim iJ As Long, posJ As Long
    Dim candidateJ As String
    Dim foundPatternJ As Boolean

    ' Wyłączenie odświeżania ekranu dla przyspieszenia działania
    Application.ScreenUpdating = False

    ' Ustawienie aktywnego arkusza
    Set wsJ = ActiveSheet

    ' Znalezienie ostatniego wiersza na podstawie kolumny A
    lastRowJ = wsJ.Cells(wsJ.Rows.Count, "A").End(xlUp).Row

    ' Pobranie danych z zakresu A:F do tablicy
    arrJ = wsJ.Range("A1:F" & lastRowJ).Value

    ' Przetwarzanie tablicowe – iteracja po każdym wierszu
    For iJ = 1 To UBound(arrJ, 1)
        ' Pobieramy zawartość komórki w kolumnie E jako tekst
        candidateJ = CStr(arrJ(iJ, 5))
        foundPatternJ = False

        ' Szukamy pierwszego wystąpienia znaku "("
        posJ = InStr(1, candidateJ, "(")
        Do While posJ > 0
            ' Upewnij się, że po znalezionym znaku jest co najmniej 3 znaki
            If posJ + 3 <= Len(candidateJ) Then
                ' Sprawdzamy, czy mamy wzorzec: znak "(" na pozycji posJ, dwa dowolne znaki, a na pozycji posJ+3 znak ")"
                If Mid(candidateJ, posJ, 1) = "(" And Mid(candidateJ, posJ + 3, 1) = ")" Then
                    foundPatternJ = True
                    Exit Do
                End If
            End If
            ' Szukamy kolejnego wystąpienia "("
            posJ = InStr(posJ + 1, candidateJ, "(")
        Loop

        ' Jeśli wzorzec NIE został znaleziony, zmieniamy wartość w kolumnie F na "KOMPUTER"
        If Not foundPatternJ Then
            arrJ(iJ, 6) = "KOMPUTER"
        End If
    Next iJ

    ' Zapisujemy zmodyfikowaną tablicę z powrotem do arkusza
    wsJ.Range("A1:F" & lastRowJ).Value = arrJ

    ' Przywracamy odświeżanie ekranu
    Application.ScreenUpdating = True

    '...
    'Wyszukiwanie czynności i procesu z data.xlsx/Activities2
    Dim lastRowK As Long
    lastRowK = Cells(Rows.Count, "A").End(xlUp).Row

    'Range("G1:G" & lastRowK).FormulaR1C1 = "=VLOOKUP(RC[-1],'[data.xlsx]Activities2'!C1:C3,2,0)"
    Range("G1:G" & lastRowK).FormulaR1C1 = "=VLOOKUP(RC[-1],'G:\JARAN\logistyka\[data.xlsx]Activities2'!C1:C3,2,0)"
    Range("G1:G" & lastRowK).Value = Range("G1:G" & lastRowK).Value

    'Range("H1:H" & lastRowK).FormulaR1C1 = "=VLOOKUP(RC[-2],'[data.xlsx]Activities2'!C1:C3,3,0)"
    Range("H1:H" & lastRowK).FormulaR1C1 = "=VLOOKUP(RC[-2],'G:\JARAN\logistyka\[data.xlsx]Activities2'!C1:C3,3,0)"
    Range("H1:H" & lastRowK).Value = Range("H1:H" & lastRowK).Value

    '...
    'ObliczRoznice(CzasCzynności)I – A (Data) + B (Godzina), warunek D, próg bez zmian
    Dim wsM As Worksheet
    Dim lastRowM As Long
    Dim arrAM As Variant, arrBM As Variant, arrDM As Variant, arrIM As Variant
    Dim iM As Long
    Dim thresholdM As Double
    Dim dtCurrM As Double, dtNextM As Double

    Dim aCurrM As Variant, bCurrM As Variant, aNextM As Variant, bNextM As Variant

    Application.ScreenUpdating = False

    ' Liczymy na arkuszu "Dane"
    Set wsM = daneSheet

    lastRowM = wsM.Cells(wsM.Rows.Count, "B").End(xlUp).Row
    If lastRowM < 1 Then GoTo CleanUpM

    arrAM = wsM.Range("A1:A" & lastRowM).Value
    arrBM = wsM.Range("B1:B" & lastRowM).Value
    arrDM = wsM.Range("D1:D" & lastRowM).Value

    ReDim arrIM(1 To lastRowM, 1 To 1)

    For iM = 1 To lastRowM - 1

        If arrDM(iM + 1, 1) = arrDM(iM, 1) Then

            aCurrM = arrAM(iM, 1)
            bCurrM = arrBM(iM, 1)
            aNextM = arrAM(iM + 1, 1)
            bNextM = arrBM(iM + 1, 1)

            ' --- dtCurrM = Data + Godzina (bezpieczna konwersja) ---
            dtCurrM = 0
            If IsNumeric(aCurrM) Then
                dtCurrM = CDbl(aCurrM)
            ElseIf IsDate(aCurrM) Then
                dtCurrM = CDbl(CDate(aCurrM))
            End If

            If IsNumeric(bCurrM) Then
                dtCurrM = dtCurrM + CDbl(bCurrM)
            ElseIf IsDate(bCurrM) Then
                dtCurrM = dtCurrM + CDbl(CDate(bCurrM))
            End If

            ' --- dtNextM = Data + Godzina (bezpieczna konwersja) ---
            dtNextM = 0
            If IsNumeric(aNextM) Then
                dtNextM = CDbl(aNextM)
            ElseIf IsDate(aNextM) Then
                dtNextM = CDbl(CDate(aNextM))
            End If

            If IsNumeric(bNextM) Then
                dtNextM = dtNextM + CDbl(bNextM)
            ElseIf IsDate(bNextM) Then
                dtNextM = dtNextM + CDbl(CDate(bNextM))
            End If

            arrIM(iM, 1) = dtNextM - dtCurrM

        Else
            arrIM(iM, 1) = 0
        End If

    Next iM

    arrIM(lastRowM, 1) = 0

    wsM.Range("I1").Resize(lastRowM, 1).Value = arrIM

    ' --- Druga część: próg bez zmian ---
    thresholdM = 0.334027777777778

    arrIM = wsM.Range("I1:I" & lastRowM).Value

    For iM = 1 To UBound(arrIM, 1)
        If Not IsEmpty(arrIM(iM, 1)) Then
            If arrIM(iM, 1) > thresholdM Then
                arrIM(iM, 1) = 0
            End If
        End If
    Next iM

    wsM.Range("I1:I" & lastRowM).Value = arrIM

CleanUpM:
    Application.ScreenUpdating = True

        '<<<<<<<<<<<<<<<<<<<<<<<<< KOREKTY >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    '... Wyłączona
    ' SprawdzPrzedzial-KonserwacjaSystemowa(04:00-04:15)
        'Dim wsKOR As Worksheet
        'Dim lastRowKOR As Long
        'Dim arrBKOR As Variant, arrIKOR As Variant
        'Dim iKOR As Long
        'Dim minZakresKOR As Double, maxZakresKOR As Double

        ' Pracujemy w aktywnym arkuszu
        'Set wsKOR = ActiveSheet

        ' Znajdź ostatni wiersz w kolumnie B
        'lastRowKOR = wsKOR.Cells(wsKOR.Rows.Count, "B").End(xlUp).Row

        ' Wczytaj dane z kolumny B i I do tablic
        'arrBKOR = wsKOR.Range("B1:B" & lastRowKOR).Value
        'arrIKOR = wsKOR.Range("I1:I" & lastRowKOR).Value

        ' Zakres do sprawdzania
        'minZakresKOR = 0.166655092592593
        'maxZakresKOR = 0.177094907407407

        ' Przejdź po tablicy od B1 do przedostatniego wiersza
        'For iKOR = 1 To UBound(arrBKOR, 1) - 1
            ' Sprawdź czy obie wartości są liczbami
            'If IsNumeric(arrBKOR(iKOR, 1)) And IsNumeric(arrBKOR(iKOR + 1, 1)) Then
                ' Sprawdź, czy zakres zawiera wartość z przedziału
                'If arrBKOR(iKOR, 1) <= maxZakresKOR And arrBKOR(iKOR + 1, 1) >= minZakresKOR Then
                    'arrIKOR(iKOR, 1) = 0
                'End If
            'End If
        'Next iKOR

        ' Wpisz dane z powrotem do kolumny I
        'wsKOR.Range("I1:I" & lastRowKOR).Value = arrIKOR

    '...
    ' PrzesunLogowanie-mag
    Dim wsKORA As Worksheet
    Set wsKORA = ActiveSheet

    Dim lastRowKORA As Long
    lastRowKORA = wsKORA.Cells(wsKORA.Rows.Count, "D").End(xlUp).Row

    ' Wczytaj dane do tablic
    Dim dataDKORA As Variant, dataFKORA As Variant, dataHKORA As Variant
    dataDKORA = wsKORA.Range("D1:D" & lastRowKORA).Value
    dataFKORA = wsKORA.Range("F1:F" & lastRowKORA).Value
    dataHKORA = wsKORA.Range("H1:H" & lastRowKORA).Value

    Dim iKORA As Long
    For iKORA = lastRowKORA - 1 To 1 Step -1

        If dataDKORA(iKORA, 1) = dataDKORA(iKORA + 1, 1) Then
            If UCase(Trim(dataFKORA(iKORA, 1))) = "LOGOWANIE - MAG" Then

                ' Szukaj w dół pierwszej wartości ? "Czynności inne"
                Dim jKORA As Long
                For jKORA = iKORA + 1 To lastRowKORA
                    If dataDKORA(jKORA, 1) <> dataDKORA(iKORA, 1) Then Exit For
                    If UCase(Trim(dataHKORA(jKORA, 1))) <> "CZYNNOSCI INNE" And Trim(dataHKORA(jKORA, 1)) <> "" Then
                        dataHKORA(iKORA, 1) = dataHKORA(jKORA, 1)
                        Exit For
                    End If
                Next jKORA

            End If
        End If

    Next iKORA

    ' Zapisz zaktualizowaną kolumnę H
    wsKORA.Range("H1:H" & lastRowKORA).Value = dataHKORA




'=========================================================
' KOREKTA PRZEJŚĆ PRZEZ PÓŁNOC (KORB)
'    Statusy:
'      - NOWE wiersze:                      K = "S+"
'      - Wiersz źródłowy górny (kopiowany): K = "S-"
'      - Wiersz źródłowy dolny (kopiowany): K = "S-"
'      - Pozostałe oryginalne:              K = "S"
'
'    + Czas w kolumnie B jako ułamek doby
'    + Warunek: jeśli suma I w dwóch nowych wierszach S+ > 7h (7/24)
'               to w obu tych wierszach S+ ustawiamy I = 0
'
'    KOREKTA (NOWA LOGIKA KOPIOWANIA):
'      Po wstawieniu 2 nowych wierszy:
'        - OBA nowe wiersze (i+1 i i+2) są kopiowane z ostatniego przed północą (wiersz i)
'        - (nie kopiujemy już drugiego nowego z pierwszego po północy)
'=========================================================

Dim wsKORB As Worksheet
Dim lastRowKORB As Long

Dim arrAKORB As Variant, arrBKORB As Variant, arrCKORB As Variant, arrDKORB As Variant
Dim iKORB As Long

Dim dCurrKORB As Double, dNextKORB As Double
Dim tCurrKORB As Double, tNextKORB As Double

Dim vBCurrKORB As Variant, vBNextKORB As Variant
Dim sBCurrKORB As String, sBNextKORB As String

Dim dtUpKORB As Double, dtDownKORB As Double
Dim dtSplitEndKORB As Double, dtSplitStartKORB As Double

Dim lastRow2KORB As Long
Dim arrIKORB As Variant, arrJKORB As Variant, arrKKORB As Variant
Dim rKORB As Long

Dim lim7hKORB As Double
Dim sumNewIKORB As Double

Set wsKORB = daneSheet

lastRowKORB = wsKORB.Cells(wsKORB.Rows.Count, "A").End(xlUp).Row
If lastRowKORB < 2 Then GoTo EndKORB

lim7hKORB = 7# / 24#   ' 7 godzin jako ułamek doby

'=========================================================
' BLOK 1 (KORB): wstawianie 2 wierszy na przejściu północy
'  + ustawienie statusów S+/S-
'  + warunek 7h tylko dla nowych S+
'  + KOREKTA: oba nowe wiersze kopiowane z ostatniego przed północą
'=========================================================
arrAKORB = wsKORB.Range("A1:A" & lastRowKORB).Value   ' Data
arrBKORB = wsKORB.Range("B1:B" & lastRowKORB).Value   ' Godzina (ułamek doby)
arrCKORB = wsKORB.Range("C1:C" & lastRowKORB).Value   ' Milisekundy
arrDKORB = wsKORB.Range("D1:D" & lastRowKORB).Value   ' Operator

For iKORB = lastRowKORB - 1 To 1 Step -1

    ' tylko jeśli operator w D jest ten sam na przejściu
    If arrDKORB(iKORB, 1) = arrDKORB(iKORB + 1, 1) Then

        ' --- DATA (A) ---
        If IsNumeric(arrAKORB(iKORB, 1)) Then
            dCurrKORB = CDbl(arrAKORB(iKORB, 1))
        Else
            dCurrKORB = CDbl(DateValue(CStr(arrAKORB(iKORB, 1))))
        End If

        If IsNumeric(arrAKORB(iKORB + 1, 1)) Then
            dNextKORB = CDbl(arrAKORB(iKORB + 1, 1))
        Else
            dNextKORB = CDbl(DateValue(CStr(arrAKORB(iKORB + 1, 1))))
        End If

        ' --- CZAS (B) jako ułamek doby ---
        vBCurrKORB = arrBKORB(iKORB, 1)
        vBNextKORB = arrBKORB(iKORB + 1, 1)

        If IsNumeric(vBCurrKORB) Then
            tCurrKORB = CDbl(vBCurrKORB)
        Else
            sBCurrKORB = Replace(Trim(CStr(vBCurrKORB)), ",", ".")
            If IsNumeric(sBCurrKORB) Then tCurrKORB = CDbl(sBCurrKORB) Else tCurrKORB = 0#
        End If

        If IsNumeric(vBNextKORB) Then
            tNextKORB = CDbl(vBNextKORB)
        Else
            sBNextKORB = Replace(Trim(CStr(vBNextKORB)), ",", ".")
            If IsNumeric(sBNextKORB) Then tNextKORB = CDbl(sBNextKORB) Else tNextKORB = 0#
        End If

        ' --- PRZEJŚCIE PRZEZ PÓŁNOC ---
        If dNextKORB > dCurrKORB Then
            If tNextKORB < tCurrKORB Then

                dtUpKORB = dCurrKORB + tCurrKORB
                dtDownKORB = dNextKORB + tNextKORB

                ' Wstaw 2 wiersze pomiędzy iKORB oraz iKORB+1
                wsKORB.Rows(iKORB + 1).Resize(2).Insert shift:=xlDown

                '=========================================================
                ' KOREKTA: OBA nowe wiersze kopiujemy z ostatniego przed północą
                '=========================================================
                wsKORB.Rows(iKORB).Copy
                wsKORB.Rows(iKORB + 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

                wsKORB.Rows(iKORB).Copy
                wsKORB.Rows(iKORB + 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

                ' --- NOWY PRZED PÓŁNOCĄ: 23:59:59 + ms=999 ---
                wsKORB.Cells(iKORB + 1, "B").Value = (86399# / 86400#)   ' ułamek doby
                wsKORB.Cells(iKORB + 1, "C").Value = "999"
                wsKORB.Cells(iKORB + 1, "B").NumberFormat = "0.000000000"

                dtSplitEndKORB = dCurrKORB + (86399# / 86400#)
                wsKORB.Cells(iKORB + 1, "I").Value = dtSplitEndKORB - dtUpKORB

                ' --- NOWY PO PÓŁNOCY: 00:00:00 + ms=000 ---
                wsKORB.Cells(iKORB + 2, "B").Value = 0#
                wsKORB.Cells(iKORB + 2, "C").Value = "000"
                wsKORB.Cells(iKORB + 2, "B").NumberFormat = "0.000000000"

                dtSplitStartKORB = dNextKORB + 0#
                wsKORB.Cells(iKORB + 2, "I").Value = dtDownKORB - dtSplitStartKORB

                ' --- STATUSY ---
                wsKORB.Cells(iKORB + 1, "K").Value = "S+"
                wsKORB.Cells(iKORB + 2, "K").Value = "S+"
                wsKORB.Cells(iKORB, "K").Value = "S-"          ' górny źródłowy
                wsKORB.Cells(iKORB + 3, "K").Value = "S-"      ' dolny źródłowy (oryginalny po przesunięciu)

                ' --- WARUNEK 7h tylko dla nowych wierszy S+ (I) ---
                sumNewIKORB = CDbl(wsKORB.Cells(iKORB + 1, "I").Value) + CDbl(wsKORB.Cells(iKORB + 2, "I").Value)
                If sumNewIKORB > lim7hKORB Then
                    wsKORB.Cells(iKORB + 1, "I").Value = 0#
                    wsKORB.Cells(iKORB + 2, "I").Value = 0#
                End If

            End If
        End If

    End If

Next iKORB

Application.CutCopyMode = False

'=========================================================
' BLOK 2 (KORB): I -> J + status "S" dla reszty
'   - kopiujemy I do J (wartości)
'   - K = "S" dla wszystkich wierszy, które nie mają S+ ani S-
'=========================================================
lastRow2KORB = wsKORB.Cells(wsKORB.Rows.Count, "A").End(xlUp).Row
If lastRow2KORB < 1 Then GoTo EndKORB

arrIKORB = wsKORB.Range("I1:I" & lastRow2KORB).Value
wsKORB.Range("J1:J" & lastRow2KORB).Value = arrIKORB

For rKORB = 1 To lastRow2KORB
    If wsKORB.Cells(rKORB, "K").Value <> "S+" And wsKORB.Cells(rKORB, "K").Value <> "S-" Then
        wsKORB.Cells(rKORB, "K").Value = "S"
    End If
Next rKORB

'=========================================================
' BLOK 3 (KORB): zerowanie czasów I/J na sekwencji północnej
'   Układ 4 wierszy:
'     r:   "S-"   (źródłowy górny)
'     r+1: "S+"   (nowy przed północą)
'     r+2: "S+"   (nowy po północy)
'     r+3: "S-"   (źródłowy dolny)
'
'   Wtedy:
'     - r     : I = 0
'     - r+1   : J = 0
'     - r+2   : J = 0
'     - r+3   : I = 0
'=========================================================
arrIKORB = wsKORB.Range("I1:I" & lastRow2KORB).Value
arrJKORB = wsKORB.Range("J1:J" & lastRow2KORB).Value
arrKKORB = wsKORB.Range("K1:K" & lastRow2KORB).Value

For rKORB = 1 To lastRow2KORB - 3

    If CStr(arrKKORB(rKORB, 1)) = "S-" And _
       CStr(arrKKORB(rKORB + 1, 1)) = "S+" And _
       CStr(arrKKORB(rKORB + 2, 1)) = "S+" And _
       CStr(arrKKORB(rKORB + 3, 1)) = "S-" Then

        arrIKORB(rKORB, 1) = 0
        arrJKORB(rKORB + 1, 1) = 0
        arrJKORB(rKORB + 2, 1) = 0
        arrIKORB(rKORB + 3, 1) = 0

    End If

Next rKORB

wsKORB.Range("I1:I" & lastRow2KORB).Value = arrIKORB
wsKORB.Range("J1:J" & lastRow2KORB).Value = arrJKORB

'=========================================================
' BLOK 4 (KORB): korekta daty dla wiersza S+ o czasie 00:00:00
'   - jeśli w kolumnie K = "S+" i czas w kolumnie B = 0
'   - ustaw datę z kolumny A na datę z następnego wiersza
'=========================================================
Dim arrAKORB2 As Variant, arrBKORB2 As Variant, arrKKORB2 As Variant
Dim rKORBX As Long

arrAKORB2 = wsKORB.Range("A1:A" & lastRow2KORB).Value
arrBKORB2 = wsKORB.Range("B1:B" & lastRow2KORB).Value
arrKKORB2 = wsKORB.Range("K1:K" & lastRow2KORB).Value

For rKORBX = 1 To lastRow2KORB - 1
    If CStr(arrKKORB2(rKORBX, 1)) = "S+" Then
        If IsNumeric(arrBKORB2(rKORBX, 1)) Then
            If CDbl(arrBKORB2(rKORBX, 1)) = 0# Then
                arrAKORB2(rKORBX, 1) = arrAKORB2(rKORBX + 1, 1)
            End If
        End If
    End If
Next rKORBX

wsKORB.Range("A1:A" & lastRow2KORB).Value = arrAKORB2

EndKORB:




        '<<<<<<<<<<<<<<<<<<<<<<<<< KOREKTY - KONIEC >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

        '<<<<<<<<<<<<<<<<<<<<<<<<< MODUŁ PRZETWARZAJĄCY - C.D. >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    '...
    'WstawTerminal(KomputerJako"C")

    ' =========================================================
    ' BLOK: Terminal z kolumny E -> wynik do kolumny L (TEKST)
    '
    ' Zasady:
    ' 1) Jeśli w kolumnie F = "KOMPUTER" › w L wpisz "C"
    ' 2) Jeśli w F <> "KOMPUTER":
    '    - z kolumny E wyciągnij PIERWSZY nawias (...) od lewej
    '    - wpisz jego zawartość do kolumny L
    ' 3) WSZYSTKO w kolumnie L ma być TEKSTEM
    ' =========================================================

    Dim wsO As Worksheet
    Dim lastRowO As Long
    Dim arrEO As Variant, arrFO As Variant, arrLO() As Variant
    Dim iO As Long
    Dim sEO As String, sFO As String
    Dim p1O As Long, p2O As Long
    Dim terminalO As String

    Set wsO = ActiveSheet

    lastRowO = wsO.Cells(wsO.Rows.Count, "E").End(xlUp).Row
    If lastRowO < 1 Then Exit Sub

    ' Wymuś format TEKSTOWY kolumny L
    wsO.Columns("L").NumberFormat = "@"

    arrEO = wsO.Range("E1:E" & lastRowO).Value
    arrFO = wsO.Range("F1:F" & lastRowO).Value
    ReDim arrLO(1 To lastRowO, 1 To 1)

    For iO = 1 To lastRowO

        sFO = Trim$(CStr(arrFO(iO, 1)))

        ' --- KOMPUTER ---
        If UCase$(sFO) = "KOMPUTER" Then
            arrLO(iO, 1) = "C"

        ' --- TERMINAL Z NAWIASU ---
        Else
            sEO = CStr(arrEO(iO, 1))
            terminalO = ""

            p1O = InStr(1, sEO, "(", vbTextCompare)
            If p1O > 0 Then
                p2O = InStr(p1O + 1, sEO, ")", vbTextCompare)
                If p2O > p1O Then
                    terminalO = Mid$(sEO, p1O + 1, p2O - p1O - 1)
                End If
            End If

            ' Wymuszenie TEKSTU (żeby np. 28 nie było liczbą)
            arrLO(iO, 1) = CStr(terminalO)
        End If

    Next iO

    wsO.Range("L1").Resize(lastRowO, 1).Value = arrLO


    '...
    'WstawieniaKluczaDoUnikatowychWierszyWBazie-PolaczA_B_C_D_L_WstawDoM

    ' =========================================================
    ' BLOK: Budowa klucza A-B-C-D-L › kolumna M
    '
    ' A – Data
    ' B – Czas (ułamek doby)
    ' C – Milisekundy
    ' D – Operator
    ' L – Terminal
    '
    ' Separator: "-"
    ' Wynik: kolumna M
    ' =========================================================

    Dim wsP As Worksheet
    Dim lastRowP As Long
    Dim arrAP As Variant, arrBP As Variant, arrCP As Variant
    Dim arrDP As Variant, arrLP As Variant
    Dim arrMP() As Variant
    Dim iP As Long

    Set wsP = ActiveSheet

    lastRowP = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    If lastRowP < 1 Then Exit Sub

    arrAP = wsP.Range("A1:A" & lastRowP).Value
    arrBP = wsP.Range("B1:B" & lastRowP).Value
    arrCP = wsP.Range("C1:C" & lastRowP).Value
    arrDP = wsP.Range("D1:D" & lastRowP).Value
    arrLP = wsP.Range("L1:L" & lastRowP).Value

    ReDim arrMP(1 To lastRowP, 1 To 1)

    For iP = 1 To lastRowP
        arrMP(iP, 1) = _
            CStr(arrAP(iP, 1)) & "-" & _
            CStr(arrBP(iP, 1)) & "-" & _
            CStr(arrCP(iP, 1)) & "-" & _
            CStr(arrDP(iP, 1)) & "-" & _
            CStr(arrLP(iP, 1))
    Next iP

    wsP.Range("M1").Resize(lastRowP, 1).Value = arrMP




'...
'WstawNaglowkiZakresu
    Dim wsR As Worksheet
    Set wsR = ActiveSheet

    ' Wstaw pusty wiersz na samej górze
    wsR.Rows(1).Insert shift:=xlDown

    ' Nagłówki od A do M
    wsR.Cells(1, "A").Value = "Data"
    wsR.Cells(1, "B").Value = "Godzina"
    wsR.Cells(1, "C").Value = "Milisekunda"
    wsR.Cells(1, "D").Value = "Operator"
    wsR.Cells(1, "E").Value = "NazwaRealizacji"
    wsR.Cells(1, "F").Value = "SkrotRealizacji"
    wsR.Cells(1, "G").Value = "Czynnosc"
    wsR.Cells(1, "H").Value = "Proces"
    wsR.Cells(1, "I").Value = "SesjaDobowa"
    wsR.Cells(1, "J").Value = "SesjaZmianowa"
    wsR.Cells(1, "K").Value = "Sesja"
    wsR.Cells(1, "L").Value = "Terminal"
    wsR.Cells(1, "M").Value = "Klucz"

Exit Sub



    '<<<<<<<<<<<<<<<<<<<<<<<<< KONIEC PRZETWARZANIA, ZAPIS DO PLIKU >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

MsgBox "Zapisz arkusz do osobnego pliku" & vbNewLine & _
       "Nastąpi teraz otwarcie okna dialogowego do zapisu", vbExclamation, "Uwaga"
Application.DisplayAlerts = False

Dim sheetDate As String
' Pobierz datę z komórki A1 arkusza "Dane" i sformatuj jako rrrrmmdd
sheetDate = Format(ThisWorkbook.Worksheets("Dane").Range("A1").Value, "yyyymmdd")

fileDialog:
Dim saveFileName As Variant
saveFileName = Application.GetSaveAsFilename( _
    "tbl-" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & "-" & sheetDate, _
    "Skoroszyt pliku Excel z obsługą makr,*.xlsm", , "Zapisz plik")

If saveFileName = False Then
    MsgBox "Anulowanie zapisu" & vbNewLine & _
           "Arkusz musi być zapisany w osobnym pliku" & vbNewLine & _
           "Jeszcze raz otwórz WerehouseReports.xlsm i ponów wczytywanie wydruku", _
           vbExclamation, "Uwaga"
    ThisWorkbook.Close
    Exit Sub
End If

If Dir(saveFileName) = "" Then
    ActiveWorkbook.SaveAs saveFileName
Else
    Dim mbox As VbMsgBoxResult
    mbox = MsgBox("Ten plik już istnieje. Czy chcesz zastąpić plik?", vbYesNoCancel)
    If mbox = vbYes Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs saveFileName
        Application.DisplayAlerts = True
    ElseIf mbox = vbNo Then
        GoTo fileDialog
    Else
        Exit Sub
    End If
End If

Application.DisplayAlerts = True


    '<<<<<<<<<<<<<<<<<<<<<<<<< EKSPORTOWANIE DO PLIKU CSV>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    Dim wbCSVN As Workbook
    Dim currentPathN As String
    Dim fileNameN As String
    Dim wsN As Worksheet
    Dim baseNameN As String
    Dim dotPosN As Long

    ' Ustal bieżący katalog – jeśli skoroszyt nie ma zapisanego folderu, użyj CurDir
    If ThisWorkbook.Path = "" Then
        currentPathN = CurDir
    Else
        currentPathN = ThisWorkbook.Path
    End If

    ' Ustaw referencję do jedynego arkusza (przykładowo pierwszy arkusz)
    Set wsN = ThisWorkbook.Worksheets(1)

    ' Ustal nazwę pliku źródłowego (bez rozszerzenia)
    dotPosN = InStrRev(ThisWorkbook.Name, ".")
    If dotPosN > 0 Then
        baseNameN = Left(ThisWorkbook.Name, dotPosN - 1)
    Else
        baseNameN = ThisWorkbook.Name
    End If

    ' Utwórz nazwę pliku CSV – ta sama nazwa co skoroszyt źródłowy z rozszerzeniem .csv
    fileNameN = currentPathN & "\" & baseNameN & ".csv"

    ' Skopiuj arkusz do nowego skoroszytu
    wsN.Copy
    Set wbCSVN = ActiveWorkbook

    ' Wyłącz alerty, aby nie pojawiały się komunikaty przy zapisie
    Application.DisplayAlerts = False

    ' Zapisz nowy skoroszyt jako CSV
    wbCSVN.SaveAs fileName:=fileNameN, FileFormat:=xlCSV, Local:=True

    ' Zamknij nowy skoroszyt bez zapisywania zmian
    wbCSVN.Close SaveChanges:=False

    ' Przywróć alerty
    Application.DisplayAlerts = True

    MsgBox "Eksport arkusza zakończony. Plik CSV został zapisany jako:" & vbNewLine & fileNameN



End Sub
