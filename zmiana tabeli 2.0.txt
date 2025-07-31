Sub PrzeksztalcTabeleZagrozen()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, i As Long, wynikRow As Long
    Dim tekstProfilaktyka As String
    Dim zagrozenie As String, zrodlo As String, skutek As String
    Dim c As Long, p As Long, cn As Long, pn As Long
    Dim mapaZamian As Object
    Dim valD As String, valE As String, valH As String, valI As String
    Dim typ As String
    Dim tempCSVPath As String, csvText As String
    Dim fNum As Integer

    ' Mapa zamiany liter M/S/D › liczby 1/2/3
    Set mapaZamian = CreateObject("Scripting.Dictionary")
    mapaZamian.Add "M", 1
    mapaZamian.Add "S", 2
    mapaZamian.Add "D", 3

    Set ws = ActiveSheet
    Set wsOut = Worksheets.Add
    wsOut.Name = "Wynik"

    ' Nagłówki
    wsOut.Range("A1:J1").Value = Array("Zagrożenie", "Typ", "Źródło", "Skutek", "C", "P", "Profilaktyka", "Cn", "Pn", "Uwagi")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    wynikRow = 2

    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            If zagrozenie <> "" Then
                typ = RozpoznajTyp(zagrozenie)
                wsOut.Cells(wynikRow, 1).Value = zagrozenie
                wsOut.Cells(wynikRow, 2).Value = typ
                wsOut.Cells(wynikRow, 3).Value = zrodlo
                wsOut.Cells(wynikRow, 4).Value = skutek
                wsOut.Cells(wynikRow, 5).Value = c
                wsOut.Cells(wynikRow, 6).Value = p
                wsOut.Cells(wynikRow, 7).Value = tekstProfilaktyka
                wsOut.Cells(wynikRow, 8).Value = cn
                wsOut.Cells(wynikRow, 9).Value = pn
                wsOut.Cells(wynikRow, 10).Value = ""
                wynikRow = wynikRow + 1
            End If

            zagrozenie = ws.Cells(i, 1).Value
            zrodlo = ws.Cells(i, 3).Value
            skutek = ws.Cells(i, 2).Value
            tekstProfilaktyka = ws.Cells(i, 7).Value

            valD = Trim(ws.Cells(i, 4).Value)
            valE = Trim(ws.Cells(i, 5).Value)
            valH = Trim(ws.Cells(i, 8).Value)
            valI = Trim(ws.Cells(i, 9).Value)

            If mapaZamian.exists(valD) Then c = mapaZamian(valD) Else c = ""
            If mapaZamian.exists(valE) Then p = mapaZamian(valE) Else p = ""
            If mapaZamian.exists(valH) Then cn = mapaZamian(valH) Else cn = ""
            If mapaZamian.exists(valI) Then pn = mapaZamian(valI) Else pn = ""
        Else
            If ws.Cells(i, 7).Value <> "" Then
                tekstProfilaktyka = tekstProfilaktyka & vbLf & ws.Cells(i, 7).Value
            End If
        End If
    Next i

    If zagrozenie <> "" Then
        typ = RozpoznajTyp(zagrozenie)
        wsOut.Cells(wynikRow, 1).Value = zagrozenie
        wsOut.Cells(wynikRow, 2).Value = typ
        wsOut.Cells(wynikRow, 3).Value = zrodlo
        wsOut.Cells(wynikRow, 4).Value = skutek
        wsOut.Cells(wynikRow, 5).Value = c
        wsOut.Cells(wynikRow, 6).Value = p
        wsOut.Cells(wynikRow, 7).Value = tekstProfilaktyka
        wsOut.Cells(wynikRow, 8).Value = cn
        wsOut.Cells(wynikRow, 9).Value = pn
        wsOut.Cells(wynikRow, 10).Value = ""
    End If

    ' Eksport do CSV
    tempCSVPath = ThisWorkbook.Path & "\wynik.csv"
    csvText = ""

    For i = 1 To wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
        For j = 1 To 10
            csvText = csvText & """" & Replace(wsOut.Cells(i, j).Text, """", """""") & """"
            If j < 10 Then csvText = csvText & ";"
        Next j
        csvText = csvText & vbCrLf
    Next i

        Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Charset = "windows-1250"
        .Open
        .WriteText csvText
        .SaveToFile tempCSVPath, 2 ' nadpisz jeśli istnieje
        .Close
    End With

    MsgBox "Gotowe! Dane przekształcone i zapisane jako CSV: " & tempCSVPath
End Sub

Function RozpoznajTyp(zagrozenie As String) As String
    Dim z As String
    z = LCase(zagrozenie)

    If z Like "*hałas*" Or z Like "*promieniowanie*" Or z Like "*drgania*" Or z Like "*temperatura*" Or z Like "*oświetlenie*" Then
        RozpoznajTyp = "F"
    ElseIf z Like "*upadek*" Or z Like "*uderzenie*" Or z Like "*skaleczenie*" Or z Like "*porażenie*" Or z Like "*zgniecenie*" Then
        RozpoznajTyp = "W"
    ElseIf z Like "*bakterie*" Or z Like "*wirusy*" Or z Like "*biologiczne*" Then
        RozpoznajTyp = "B"
    ElseIf z Like "*substancje*" Or z Like "*chemiczne*" Or z Like "*pyły*" Or z Like "*gazy*" Then
        RozpoznajTyp = "C"
    ElseIf z Like "*obciążenie*" Or z Like "*monotonia*" Or z Like "*stres fizyczny*" Then
        RozpoznajTyp = "E"
    ElseIf z Like "*presja*" Or z Like "*mobbing*" Or z Like "*przemoc*" Or z Like "*psychiczne*" Or z Like "*psychospołeczne*" Then
        RozpoznajTyp = "P"
    Else
        RozpoznajTyp = "W"
    End If
End Function

