Public Sub test()
    '<<<<<<<<<<<<<<<<<<<<<<<<< ILOŚĆ + POZYCJE >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    ' Krok 1: ustawiamy arkusz roboczy i wiersz nagłówków.
    Dim ws As Worksheet
    Dim headerRow As Long
    Set ws = ActiveSheet
    headerRow = 1

    ' Krok 2: wyznaczamy ostatni wiersz danych.
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If lastRow < headerRow + 1 Then Exit Sub

    ' Krok 3: szukamy pierwszej wolnej kolumny na nagłówki.
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If ws.Cells(headerRow, lastCol).Value = "" Then
        lastCol = 0
    End If

    ' Krok 4: wstawiamy nagłówek "Ilość" w pierwszej wolnej kolumnie.
    Dim colIlosc As Long
    Dim colIloscL As String
    colIlosc = lastCol + 1
    ws.Cells(headerRow, colIlosc).Value = "Ilość"
    colIloscL = Split(ws.Cells(headerRow, colIlosc).Address(False, False), CStr(headerRow))(0)

    ' Krok 5: wstawiamy nagłówek "Pozycje" w kolejnej wolnej kolumnie.
    Dim colPozycje As Long
    Dim colPozycjeL As String
    colPozycje = colIlosc + 1
    ws.Cells(headerRow, colPozycje).Value = "Pozycje"
    colPozycjeL = Split(ws.Cells(headerRow, colPozycje).Address(False, False), CStr(headerRow))(0)

    ' Krok 6: wczytujemy dane do tablic (kolumny E i F).
    Dim arrNazwa As Variant
    Dim arrSkrot As Variant
    arrNazwa = ws.Range("E2:E" & lastRow).Value
    arrSkrot = ws.Range("F2:F" & lastRow).Value

    ' Krok 7: przygotowujemy tablice wynikowe.
    Dim arrIlosc() As Variant
    Dim arrPozycje() As Variant
    ReDim arrIlosc(1 To UBound(arrNazwa, 1), 1 To 1)
    ReDim arrPozycje(1 To UBound(arrSkrot, 1), 1 To 1)

    ' Krok 8: wyliczamy "Ilość" z NazwaLoga (po znaku ")").
    Dim i As Long
    Dim s As String, afterS As String, numStr As String
    Dim pos As Long, j As Long, ch As String
    Dim startPos As Long, hasDigit As Boolean

    For i = 1 To UBound(arrNazwa, 1)
        s = CStr(arrNazwa(i, 1))
        pos = InStr(1, s, ")", vbTextCompare)

        If pos > 0 Then
            afterS = Trim$(Mid$(s, pos + 1))
            If Len(afterS) > 0 Then
                numStr = ""
                hasDigit = False
                startPos = 1

                If Left$(afterS, 1) = "-" Then
                    numStr = "-"
                    startPos = 2
                End If

                For j = startPos To Len(afterS)
                    ch = Mid$(afterS, j, 1)
                    If ch >= "0" And ch <= "9" Then
                        numStr = numStr & ch
                        hasDigit = True
                    Else
                        Exit For
                    End If
                Next j

                If hasDigit Then
                    arrIlosc(i, 1) = Abs(CDbl(numStr))
                Else
                    arrIlosc(i, 1) = ""
                End If
            Else
                arrIlosc(i, 1) = ""
            End If
        Else
            arrIlosc(i, 1) = ""
        End If
    Next i

    ' Krok 9: wyliczamy "Pozycje" na podstawie SkrotLoga i listy wzorców.
    Dim patternsA As Variant
    patternsA = Array( _
        "PRZYJĘCIE TOWARU NA PÓŁKĘ", _
        "POBRANIE DO WYD.", _
        "ZESKANOWANO SKRZYNKĘ", _
        "DO PACZKI ZE SKRZYNKI", _
        "DODANIE: KARTON", _
        "[SKRZ.", _
        "INWENT", _
        "NA BRAKI", _
        "PRZENIESIENIE NA:", _
        "PRZENIESIENIE TOWARU", _
        "ROZŁOŻENIE Z LOK. ĆÄ", _
        "ROZŁOŻENIE Z LOK. ĆI", _
        "ROZŁOŻENIE Z LOK. RE", _
        "ROZŁOŻENIE Z LOK. ZWR", _
        "SPRAWDZENIE LOK", _
        "ZDJĘCIE Z PÓŁKI" _
    )

    Dim pA As Variant
    Dim sk As String
    Dim isMatch As Boolean

    For i = 1 To UBound(arrSkrot, 1)
        sk = CStr(arrSkrot(i, 1))
        isMatch = False

        If Len(sk) > 0 Then
            For Each pA In patternsA
                If InStr(1, sk, CStr(pA), vbTextCompare) > 0 Then
                    isMatch = True
                    Exit For
                End If
            Next pA
        End If

        If isMatch Then
            arrPozycje(i, 1) = sk
        Else
            arrPozycje(i, 1) = ""
        End If
    Next i

    ' Krok 10: zapisujemy wyniki do arkusza.
    ws.Range(colIloscL & "2:" & colIloscL & lastRow).Value = arrIlosc
    ws.Range(colPozycjeL & "2:" & colPozycjeL & lastRow).Value = arrPozycje
End Sub
