Attribute VB_Name = "fkJsp2Java"
Option Explicit

Private Const ART_HTML_QUELLTEXT = 1
Private Const ART_JSP_ZUWEISUNG = 2
Private Const ART_JSP_JAVA = 3

Public Sub ttt()

Dim vb_str As String

    vb_str = vb_str & vbCrLf & "<tbody>"
    vb_str = vb_str & vbCrLf & "<%"
    vb_str = vb_str & vbCrLf & "boolean knz_html_css_schalter = true;"
    vb_str = vb_str & vbCrLf & ""
    vb_str = vb_str & vbCrLf & "int akt_index = 0;"
    vb_str = vb_str & vbCrLf & ""
    vb_str = vb_str & vbCrLf & "Freigabe inst_freigabe = projekt.getFreigabeIndex( akt_index );"
    vb_str = vb_str & vbCrLf & ""
    vb_str = vb_str & vbCrLf & "while( inst_freigabe != null )"
    vb_str = vb_str & vbCrLf & "{"
    vb_str = vb_str & vbCrLf & "knz_html_css_schalter = !knz_html_css_schalter;"
    vb_str = vb_str & vbCrLf & "%>"
    vb_str = vb_str & vbCrLf & "<tr <% if ( knz_html_css_schalter ) { %> class=""farbelistenzeile2"" <% } %> >"
    vb_str = vb_str & vbCrLf & ""
    vb_str = vb_str & vbCrLf & "<td style=""font-size: 100%; text-align:right; "" <%if (inst_freigabe.getAllgemeinerText().length() > 0) {%> title=""<%=inst_freigabe.getAllgemeinerText()%>"" <%};%> > <%=inst_freigabe.getFreigabeNummerInt()%> </td>"
    vb_str = vb_str & vbCrLf & "<td></td>"
    vb_str = vb_str & vbCrLf & "<td style=""text-align:left; overflow: hidden;"" ><%=inst_freigabe.getBezeichnung(60)              %></td>"
    vb_str = vb_str & vbCrLf & "<td style=""text-align:left; overflow: hidden;"" ><%=inst_freigabe.getEntwickler().getNachnameFmt()%></td>"
    vb_str = vb_str & vbCrLf & "<td style=""text-align:left; overflow: hidden;"" ><%=inst_freigabe.getStatusBezeichnung()          %></td>"
    vb_str = vb_str & vbCrLf & "<td style=""text-align:left; ""                  ><%=inst_freigabe.getFreigabeITDatum()            %></td>"
    vb_str = vb_str & vbCrLf & "<td style=""text-align:left; ""                  ><%=inst_freigabe.getFreigabeFBDatum()            %></td>"
    vb_str = vb_str & vbCrLf & "<td style=""text-align:left; ""                  ><%=inst_freigabe.getInstallationDatum()          %></td>"
    vb_str = vb_str & vbCrLf & ""
    vb_str = vb_str & vbCrLf & "<td><button class=""deletebutton"" type=""submit"""
    vb_str = vb_str & vbCrLf & "<%if (!benutzer.isWriteIT() || !inst_freigabe.getFreigabeStatus().isOhneFreigabe() ) {%> disabled=""disabled"" title=""Kann nicht gelöscht werden"" <%} else {%> title=""Löschen"" <%}%>"
    vb_str = vb_str & vbCrLf & "onclick=""if (confirm('Wollen Sie die Freigabe Nr. <%=inst_freigabe.getFreigabeNummerInt()%> (<%=inst_freigabe.getBezeichnung(50).trim()%>) löschen?'))"
    vb_str = vb_str & vbCrLf & "{ document.getElementById('listeFreigabenAktion').value = 'deleteFreigabe';"
    vb_str = vb_str & vbCrLf & "document.getElementById('projektNr').value = '<%=projekt.getProjektNummer()%>';"
    vb_str = vb_str & vbCrLf & "document.getElementById('freigabeNr').value = '<%=inst_freigabe.getFreigabeNummerInt()%>'"
    vb_str = vb_str & vbCrLf & "} else"
    vb_str = vb_str & vbCrLf & "return false"""
    'vb_str = vb_str & vbCrLf & "></button>"
    Debug.Print startJsp2Java(vb_str)

End Sub

'########################################################################################
'
Public Function startJsp2Java(pEingabe As String) As String

    startJsp2Java = False

    On Error GoTo errStartJsp2Java

Dim akt_csv_string        As String
Dim akt_position_end      As Long
Dim akt_position_start    As Long
Dim akt_teil_string       As String
Dim akt_csv_position      As Long
Dim art_string            As Long
Dim fkt_ergebnis          As String
Dim knz_umbruch_vorhanden As Boolean
Dim letzte_csv_position   As Long
Dim zaehler_schleife      As Long
Dim zeichen_zeilenumbruch As String
Dim zeichenfolge_ende      As String
Dim zeichenfolge_start    As String

    '
    ' Ermittlung welches Zeilenumbruchzeichen verwendet in der Eingabe verwendet wird
    '
    zeichen_zeilenumbruch = MY_CHR_13_10

    knz_umbruch_vorhanden = (InStr(1, pEingabe, zeichen_zeilenumbruch, vbBinaryCompare) > 0)

    If knz_umbruch_vorhanden = False Then

        zeichen_zeilenumbruch = Chr(13)

        knz_umbruch_vorhanden = (InStr(1, pEingabe, zeichen_zeilenumbruch, vbBinaryCompare) > 0)

    End If

    zeichenfolge_start = "<%"
    zeichenfolge_ende = "%>"

    '
    ' Start CSV-Zeichen bestimmen
    '
    akt_position_start = InStr(1, pEingabe, zeichenfolge_start)
    akt_position_end = InStr(1, pEingabe, zeichenfolge_ende)
    
    If ((akt_position_start = 0) And (akt_position_end > 0)) Then
    
        akt_csv_string = zeichenfolge_ende
    
    ElseIf ((akt_position_end = 0) And (akt_position_start > 0)) Then
    
        akt_csv_string = zeichenfolge_start
    
    ElseIf (akt_position_end < akt_position_start) Then
    
        akt_csv_string = zeichenfolge_ende
        
    Else
    
        akt_csv_string = zeichenfolge_start
        
    End If
    
    '
    ' Initialisierung der Startparameter
    '
    akt_csv_position = 1
    
    letzte_csv_position = 1

    '
    ' While-Schleife
    '
    While ((akt_csv_position > 0) And (zaehler_schleife < 32123))

        '
        ' Suche die Position des aktuellen CSV-Strings in der Eingabe
        '
        akt_csv_position = InStr(akt_csv_position, pEingabe, akt_csv_string)
        
        '
        ' Pruefung: aktuellen CSV-String gefunden ?
        '
        If (akt_csv_position > 0) Then
            
            '
            ' Heraustrennen des Ergebnisstrings aus der Eingabe fuer die Verarbeitung
            '
            akt_teil_string = Mid(pEingabe, letzte_csv_position, akt_csv_position - letzte_csv_position)
            
            akt_csv_position = akt_csv_position + Len(akt_csv_string)
            
            letzte_csv_position = akt_csv_position
            
        Else
            
            '
            ' Wurde der aktuelle CSV-String nicht mehr gefunden, ist in
            ' der Variablen "akt_csv_position" ein Wert von -1 enthalten.
            '
            ' Pruefung: Noch ungelesender Teilstring vorhanden ?
            '
            ' Dieses ist dann der Fall, wenn die letzte CSV-Position kleiner
            ' gleich der Stringlaenge ist. Von der letzten Position wird bis
            ' zum Stringende der aktuelle Teilstring ermittelt.
            '
            If (letzte_csv_position <= Len(pEingabe)) Then
                
                akt_teil_string = Mid(pEingabe, letzte_csv_position, Len(pEingabe) - letzte_csv_position)

            End If

        End If

        '
        ' Pruefung: String fuer Verarbeitung vorhanden?
        '
        If (akt_teil_string <> LEER_STRING) Then
        
            If (akt_csv_string = zeichenfolge_start) Then
            
                '
                ' Wurde die Startzeichenfolge gesucht, wird im naechsten
                ' Durchlauf die Endzeichenfolge gesucht.
                '
                akt_csv_string = zeichenfolge_ende
                
                '
                ' Wird die Startzeichenfolge gesucht, ist der aktuelle
                ' Teilstring HTML, da erst nach der Startzeichenfolge
                ' JSP-Einbettungen kommen.
                '
                art_string = ART_HTML_QUELLTEXT
                
            Else
            
                '
                ' Wurde die Endzeichenfolge gesucht, wird im naechsten
                ' Durchlauf die Startzeichenfolge gesucht.
                '
                akt_csv_string = zeichenfolge_start
                
                '
                ' Da die Endzeichenfolge gesucht wurde, ist der aktuelle
                ' Teilstring ein eingebettetes Java-Elemente. Initial wird
                ' die Art auf ART_JSP_JAVA gestellt. Startet der aktuelle
                ' Teilstring mit einem Gleichheitszeichen, wird die Art auf
                ' ART_JSP_ZUWEISUNG geaendert.
                '
                art_string = ART_JSP_JAVA
                
                If (akt_teil_string <> LEER_STRING) Then
                
                    If (Left(akt_teil_string, 1) = "=") Then
                    
                        art_string = ART_JSP_ZUWEISUNG
                        
                        akt_teil_string = Right(akt_teil_string, Len(akt_teil_string) - 1)
                        
                    End If
                
                End If
            
            End If
        
            If (art_string = ART_HTML_QUELLTEXT) Then
            
                '
                ' Generierung der Anweisungen fuer HTML-Quelltext
                '
                ' Es werden die im Teilstring vorhandenen Anfuehrungszeichen maskiert.
                '
                fkt_ergebnis = fkt_ergebnis & vbCrLf & getGeneratorXString(Replace(akt_teil_string, """", "\"""), "ART_HTML_QUELLTEXT   str_buf.append( """, """ );")

            ElseIf (art_string = ART_JSP_ZUWEISUNG) Then
                '
                ' Generierung der Anweisungen fuer JSP-Ausgaben
                '
                fkt_ergebnis = fkt_ergebnis & vbCrLf & getGeneratorXString(akt_teil_string, "ART_JSP_ZUWEISUNG    str_buf.append( ", " );")
            
            Else
                '
                ' Generierung der Anweisungen fuer eingebettete Java-Anweisungen
                '
                If (Trim(akt_teil_string) <> LEER_STRING) Then
                
                    fkt_ergebnis = fkt_ergebnis & zeichen_zeilenumbruch & zeichen_zeilenumbruch & getGeneratorXString(akt_teil_string, zeichen_zeilenumbruch & "ART_JSP_JAVA         ", "") & zeichen_zeilenumbruch
                
                End If
            
            End If

        End If

        '
        ' Am Schleifenende wird der Endlosschleifenverhinderrungszaehler um 1 erhoeht.
        '
        zaehler_schleife = zaehler_schleife + 1

    Wend

    '
    ' Entfernung von falschen Zuweisungen mit Leerstring
    '
    fkt_ergebnis = Replace(fkt_ergebnis, "str_buf.append( """" );", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "str_buf.append(  );", "")
    
    akt_teil_string = " "
    
    zaehler_schleife = 1
    
    '
    ' While-Schleife
    '
    While (zaehler_schleife < 100)

        fkt_ergebnis = Replace(fkt_ergebnis, "str_buf.append( """ & akt_teil_string & """ );", "str_buf.append( """" );")
        
        akt_teil_string = akt_teil_string & " "
    
        zaehler_schleife = zaehler_schleife + 1

    Wend

    '
    ' Doppelte Zeilenumbrueche zusammenschrumpfen
    '
    fkt_ergebnis = Replace(fkt_ergebnis, zeichen_zeilenumbruch & zeichen_zeilenumbruch & zeichen_zeilenumbruch, zeichen_zeilenumbruch)
    fkt_ergebnis = Replace(fkt_ergebnis, zeichen_zeilenumbruch & zeichen_zeilenumbruch, zeichen_zeilenumbruch)
    fkt_ergebnis = Replace(fkt_ergebnis, zeichen_zeilenumbruch & zeichen_zeilenumbruch, zeichen_zeilenumbruch)

    '
    ' Entfernung von Debug-Informationen
    '
    fkt_ergebnis = Replace(fkt_ergebnis, "ART_HTML_QUELLTEXT   ", "  ")
    fkt_ergebnis = Replace(fkt_ergebnis, "ART_JSP_ZUWEISUNG    ", "  ")
    fkt_ergebnis = Replace(fkt_ergebnis, "ART_JSP_JAVA         ", "  ")

EndFunktion:

    On Error Resume Next

    '
    ' Zuweisung des Funktionsergebnisses
    '
    startJsp2Java = fkt_ergebnis
    
    '
    ' DoEvents aufrufen
    '
    DoEvents

    Exit Function

errStartJsp2Java:

    Debug.Print ("Fehler: errStartJsp2Java: " & Err & " " & Error & " " & Erl)

    Resume EndFunktion

End Function

'########################################################################################
'
Private Function getGeneratorXString(pString As String, pZeilenPraefix As String, pZeilenSuffix As String) As String

    If (Trim(pString) = "") Then
    
        getGeneratorXString = ""
        
        Exit Function

    End If

Dim fkt_ergebnis               As String
Dim knz_umbruch_vorhanden      As Boolean
Dim zeichen_zeilenumbruch      As String
Dim zeilen_zaehler             As Long
Dim akt_zeile                  As String
Dim akt_position               As Long
Dim letzte_position            As Long
Dim ergebnis                   As String

    ergebnis = ""
    '
    ' Ermittlung welches Zeilenumbruchzeichen verwendet in der Eingabe verwendet wird
    '
    zeichen_zeilenumbruch = MY_CHR_13_10

    knz_umbruch_vorhanden = (InStr(1, pString, zeichen_zeilenumbruch, vbBinaryCompare) > 0)

    If knz_umbruch_vorhanden = False Then

        zeichen_zeilenumbruch = Chr(13)

        knz_umbruch_vorhanden = (InStr(1, pString, zeichen_zeilenumbruch, vbBinaryCompare) > 0)

    End If
    '
    ' Wenn Zeilenumbrueche vorhanden sind, wird eine Schleife gestartet.
    ' Sind keine Zeilenumbrueche vorhanden, gibt es nur einen Aufruf.
    ' Desweiteren wuerde sich die Schleifenkonstruktion verkomplizieren,
    ' wenn diese auch Zeichenketten ohne Zeilenumbruch verarbeiten sollte.
    '
    If (knz_umbruch_vorhanden) Then

        letzte_position = 1
        
        akt_position = InStr(letzte_position, pString, zeichen_zeilenumbruch, vbBinaryCompare)

        While (akt_position > 0) And (zeilen_zaehler < 32220)

            akt_zeile = Mid(pString, letzte_position, akt_position - letzte_position)

            fkt_ergebnis = fkt_ergebnis & zeichen_zeilenumbruch & pZeilenPraefix & akt_zeile & pZeilenSuffix

            letzte_position = akt_position + Len(zeichen_zeilenumbruch)

            akt_position = InStr(letzte_position, pString, zeichen_zeilenumbruch, vbBinaryCompare)

            zeilen_zaehler = zeilen_zaehler + 1
        
        Wend

        If (letzte_position <= Len(pString)) Then

            akt_zeile = Mid(pString, letzte_position, (Len(pString) - letzte_position) + 1)

            fkt_ergebnis = fkt_ergebnis & zeichen_zeilenumbruch & pZeilenPraefix & akt_zeile & pZeilenSuffix

        End If

    Else

        fkt_ergebnis = pZeilenPraefix & pString & pZeilenSuffix

    End If

    getGeneratorXString = fkt_ergebnis

End Function

