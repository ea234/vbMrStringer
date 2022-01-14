Attribute VB_Name = "fkAppMrStringer"
Option Explicit

'
'ISXF = Funktionen welche zusammengefasst werden koennten (= is x Funktion)
'       Speichern alle den "temp_string_3" im Stringarray
'

'a# 3000001
'b# 3000002
'c# 3000003
'd# 3000004
'e# 3000005
'f# 3000006
'g# 3000007

Public m_knz_aktiv               As Boolean ' public wegen initialisierung
Private m_toggle_mr_stringer_fkt As Boolean ' Umschalter fuer die Doppelbelegung bei Funktionen
Private m_zaehler_debug_print    As Integer ' Zaehler fuer fuer die Funktion "Debug-String-Ausgabe"
Private m_zaehler_string_it      As Integer ' Zaehler fuer fuer die Funktion "String-It"

Public Const FKT_CSV_VB_KONVERTER = 125

Public Const FKT_AUSRICHTER_POSITION = 1
Public Const FKT_AUSRICHTER_STRING = 2
Public Const FKT_BLOCK_ZUFALL = 3
Public Const FKT_CALC_SUMME = 4
Public Const FKT_CAMEL_CASE = 5
Public Const FKT_CLIP_ENTFERNE_TEXT = 6
Public Const FKT_CLIP_GET_TEXT = 7
Public Const FKT_CLIP_POSITION = 8
Public Const FKT_CMD_RENAME = 9
Public Const FKT_CSV_2_ZEILE = 10
Public Const FKT_CSV_CR = 201
Public Const FKT_CSV_JAVA_CASE = 11
Public Const FKT_CSV_ERSTELLE_CSV = 12
Public Const FKT_CSV_JAVA_PROP = 13
Public Const FKT_CSV_REPLACE_MARKIERUNG_MIT_CSV = 14
Public Const FKT_CSV_SWAP = 15
Public Const FKT_DUPLIZIERUNG = 16
Public Const FKT_ERSTELLE_BLOCK = 17
Public Const FKT_ERSTELLE_NAMEN = 18
Public Const FKT_ERSTELLE_XML = 19
Public Const FKT_ERSTELLE_XML_2 = 20
Public Const FKT_EXTRAHIERE_WORTE = 21
Public Const FKT_FORMAT_TXT = 22
Public Const FKT_GENERATOR_DEBUG_AUSGABE = 23
Public Const FKT_GENERATOR_IF_JAVA_SCRIPT = 24
Public Const FKT_GENERATOR_IF_JAVA_VB = 25
Public Const FKT_GENERATOR_SET_NULL = 26
Public Const FKT_GENERATOR_STRING_IT = 27
Public Const FKT_GENERATOR_VARIABLEN_DEKLARATION = 28
Public Const FKT_GENERATOR_VB_CHECK_LEER_STRING = 29
Public Const FKT_GETTER_SETTER_JAVA = 30
Public Const FKT_GETTER_SETTER_JAVA_SCRIPT = 31
Public Const FKT_GETTER_SETTER_VB = 32
Public Const FKT_GET_DIR = 33
Public Const FKT_GET_DOPPELTE_VORKOMMEN = 34
Public Const FKT_GET_EINMALIGE_VORKOMMEN = 35
Public Const FKT_GET_UNIQUE = 36
Public Const FKT_GREP_DUPLIZIERE_MARKZEILEN = 37
Public Const FKT_GREP_MARK = 38
Public Const FKT_GREP_PLUS_MINUS = 39
Public Const FKT_GREP_WORT = 40
Public Const FKT_GREP_ZAHLEN = 41
Public Const FKT_GROUP_NACH_STRING = 42
Public Const FKT_HEX_DUMP = 43
Public Const FKT_JAVA_GENERATOR = 44
Public Const FKT_JAVA_XML_WRITER_NUMMER = 45
Public Const FKT_JAVA_XML_WRITER_STRING = 46
Public Const FKT_JSON_LESEN_SCHREIBEN = 47
Public Const FKT_KONSTANTEN_UEBER_SPLIT = 48
Public Const FKT_LEERZEILEN_EINFUEGEN = 49
Public Const FKT_LEERZEILEN_LOESCHEN = 50
Public Const FKT_MAKE_LONG_DATUM = 51
Public Const FKT_MARKIERE_CSV_VORNE_ODER_HINTEN = 52
Public Const FKT_MARKIERE_DOPPELT_PLUS_1_ZEILE = 53
Public Const FKT_MARKIERE_DOPPELT_PLUS_1_ZEILE_MINUS = 54
Public Const FKT_MARKIERE_STR_VORNE_UND_HINTEN = 55
Public Const FKT_MARKIERE_VORNE_FIX = 56
Public Const FKT_MARKIERE_VORNE_ODER_HINTEN = 57
Public Const FKT_MARKIERE_VORNE_UND_HINTEN = 58
Public Const FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE = 59
Public Const FKT_MARKIERE_WORT = 60
Public Const FKT_MASKIERE_ANFZEICHEN = 61
Public Const FKT_NOTES_DEBUG_FELD_WERTE = 62
Public Const FKT_NOTES_LESEN_SCHREIBEN = 63
Public Const FKT_SET_TRENNZEICHEN = 64
Public Const FKT_SET_TRENNZEICHEN_VOR = 65
Public Const FKT_SET_TRENNZEICHEN_ZURUECK = 66
Public Const FKT_SINGLETON_JAVA = 67
Public Const FKT_SORTIEREN_ALPHABETH = 68
Public Const FKT_SORTIEREN_DATUM = 69
Public Const FKT_SORTIEREN_LAENGE = 70
Public Const FKT_SORTIEREN_ZUFALL = 71
Public Const FKT_STRING_LAENGE_AUSGEBEN = 72
Public Const FKT_STRING_LIT = 73
Public Const FKT_STRING_REMOVE = 74
Public Const FKT_STRING_SPLIT = 75
Public Const FKT_STRING_UMDREHEN = 76
Public Const FKT_STRING_VERSCHIEBEN = 77
Public Const FKT_TRIM_AUFEINANDERFOLGENDE_LEERZEICHEN = 78
Public Const FKT_TRIM_STRING_VORNE_UND_HINTEN = 79
Public Const FKT_UCASE_LCASE = 80
Public Const FKT_ZEILEN_ADD = 81
Public Const FKT_ZEILEN_BOOLEAN = 82
Public Const FKT_ZEILEN_ZAEHLER = 83
Public Const FKT_STRING_ERST = 84

Public Const LEER_STRING = ""
Public Const LEER_ZEICHEN = " "
Public Const UNTER_STRICH = "_"
 
Public Const TRENN_STRING_4 = "#5"
Public Const TRENN_STRING_5 = "#4"
Public Const TRENN_STRING_6 = "#6"
Public Const TRENN_STRING_7 = "#7"
Public Const TRENN_STRING_8 = "#8"
Public Const TRENN_STRING_9 = "#9"

'
' Markierungsstrings
' Die Markierungsstrings sind nur fuer die interne Verwendung vorhanden.
' Sie sind tempraere (Hilfs)Markierungsstrings fuer diverse Funktionen.
' Die Markierungsstrings werden bei ihrer Verwendung selber wieder durch
' Leerstrings ersetzt.
'
' Die Markierungsstrings sind so gestaltet, dass deren Zeichenfolge hoffentlich
' nie nicht selber schon Bestandteil des Eingabestrings sind. Sollte das
' doch mal der Fall sein, dann hat man halt mal Pech gehabt.
'
Public Const MARKIER_STRING_INTERN_1 = "#1#-MARKIER_STRING_1-#1#"
Public Const MARKIER_STRING_INTERN_2 = "#2#-MARKIER_STRING_2-#2#"
Public Const MARKIER_STRING_INTERN_3 = "#3#-MARKIER_STRING_3-#3#"
Public Const MARKIER_STRING_INTERN_4 = "#4#-MARKIER_STRING_4-#4#"

'
' Ausrichterstrings
' Es handelt sich hierbei um temporaere Hilfsstrings um eben Texte daran
' ausrichten zu koennen. Die Ausrichterstrings werden nach deren Verwendung
' durch Leerstrings ersetzt.
'
Public Const AUSRICHT_STRING_TEMP_1 = "#A#-AUSRICHT_STRING_TEMP_1-#A#" ' " & AUSRICHT_STRING_TEMP_1 & "
Public Const AUSRICHT_STRING_TEMP_2 = "#A#-AUSRICHT_STRING_TEMP_2-#A#" ' " & AUSRICHT_STRING_TEMP_2 & "

Private Const MARKIERUNG_DOPPELTE_VORKOMMEN = "#D#-DOPPELT-#D#"

Public Const STR_VAR_NAME_PROPERTIES_LOKAL = "inst_properties" ' " & STR_VAR_NAME_PROPERTIES_LOKAL & "

Private Const GUELTIGE_ZEICHEN_DATEI_NAME = "enirstl_audhgocmfbkVvwz1pSDA0E2RBGMIPKF9UNW3L78H4T5CZJy6xjOYXqQ&,'()" ' sortiert Haeufigkeit Deutsch

Private Const NULL_ZIFFERN_100 = "00000000000000000000000000000000000000000000000000000000000000000000000000000000"

Public Const MY_CHR_13_10 = vbCrLf 'Chr(13) & Chr(10)

Private Const POSITION_0 = 0

'################################################################################
'
Public Function getToggleMrStringerFkt() As Boolean

    getToggleMrStringerFkt = m_toggle_mr_stringer_fkt

End Function

'################################################################################
'
Public Sub setToggleMrStringerFkt(pToggleMrStringerFkt As Boolean)

    m_toggle_mr_stringer_fkt = pToggleMrStringerFkt

End Sub

'################################################################################
'
Public Function startMrStringer(pFunktion As Integer, pString As String, pSelStart As Long, pSelLength As Long, Optional pKennzeichen1 As Boolean = False, Optional pOptString1 As String = "", Optional pOptString2 As String = "", Optional pOptString3 As String = "") As String

On Error GoTo errStartMrStringer

Dim cls_string_array        As clsStringArray ' Stringarray fuer die Eingabe als auch in Teilen fuer die Erstellung des Ergebnisses
Dim zeichen_zeilenumbruch   As String         ' Das verwendete Zeilenumbruchszeichen dieser Funktion
Dim str_fkt_ergebnis        As String         ' Alternative Stringvariable fuer den Ergebnisaufbau einiger Funktionen
Dim aktuelle_zeile          As String         ' Die gesamte aktuelle Zeile aus der Eingabe
Dim akt_zeile_mark          As String         ' Die gesamte aktuelle Zeile, oder bei Markierungsverwendung nur die Markierung
Dim inhalt_markierung       As String         ' Inhalt der Markierung aus der Eingabe
Dim temp_string_1           As String         ' Temporaere Stringvariable 1
Dim temp_string_2           As String         ' Temporaere Stringvariable 2
Dim temp_string_3           As String         ' Temporaere Stringvariable 3
Dim temp_double_1           As Double         ' Temporaerer Doublewert 1
Dim temp_double_2           As Double         ' Temporaerer Doublewert 2
Dim ab_position             As Long           ' Position, ab welcher in der Eingabe abgeschnitten wird
Dim bis_position            As Long           ' Position, bis zu welcher abgeschnitten wird
Dim zeilen_zaehler          As Long           ' aktuell bearbeitete Zeile der Eingabe
Dim zeilen_anzahl           As Long           ' Anzahl der Zeilen des Eingabestrings
Dim temp_long_1             As Long           ' Temporaerer Longwert 1
Dim temp_long_2             As Long           ' Temporaerer Longwert 2
Dim temp_long_3             As Long           ' Temporaerer Longwert 3
Dim knz_benutze_markierung  As Boolean        ' Kennzeichen, ob die Markierung verwendet werden soll
Dim knz_schleifen_durchlauf As Boolean        ' Kennzeichen, ob ein Schleifendurchlauf fuer die Erstellung des Ergebnisses gemacht werden soll
    
    '
    ' Pruefung: Funktion "startMrStringer" schon aktiv?
    ' Wenn dem so ist, ist die Rueckgabe des aktuellen Aufrufes ein
    ' Leerstring und die Funktion wird verlassen.
    '
    If (m_knz_aktiv) Then
    
        startMrStringer = LEER_STRING
        
        Exit Function
    
    End If
    
    '
    ' Funktion als "aktiv" markieren
    '
    m_knz_aktiv = True
    
    '
    ' "Toggle"-Schalter
    ' Der Wert des "Toggle"-Schalters wird per Negation umgestellt.
    ' Dieser Schalter dient dazu, um zwei aufeinanderfolgende Aufrufe erkennen zu koennen.
    ' Einige Funktionen haben 2 unterschiedliche Arbeitsweisen.
    '
    ' Achtung: Wird diese Funktion von einer MrStringer-Aufgabe 2 mal aufgerufen, kann das
    '          das Ergebnis sein, dass der Togglewert nie umgestellt wird.
    '
    m_toggle_mr_stringer_fkt = Not m_toggle_mr_stringer_fkt
    
    '
    ' Variable "str_fkt_ergebnis"
    '
    ' Die Variable "str_fkt_ergebnis" speichert schlussendlich das Ergebnis dieser Funktion.
    '
    ' Es gibt 2 Arten von Funktionen:
    '
    ' 1. Funktionen, welche die Strings im Stringarray modifizieren.
    '    Bei diesen Funktionen wird das Funktionsergebnis aus dem Stringarray
    '    gelesen und der Variablen "str_fkt_ergebnis" zugewiesen.
    '
    '    Die Anzahl der Ergebniszeilen ist gleich der Anzahl der Eingabezeilen.
    '
    '
    ' 2. Funktionen, welche nur eine Teilmenge der Eingabe zurueckliefern.
    '    Bei diesen Funktionen wird die Eingabe nur fuer die Erstellung der Ausgabe
    '    benoetigt. Das ist z.B. bei Generator- oder Grep-Funktionen der Fall.
    '
    '    Die Anzahl der Ergebniszeilen weicht von der Anzahl der Eingabezeilen ab.
    '
    '
    ' Ist nach der Hauptschleife die Variable "str_fkt_ergebnis" ein Leerstring, wird
    ' das Funktionsergebnis aus dem Stringarray gelesen.
    '
    ' Es gibt noch eine dritte Art von Funktionen in dieser Funktion. Das sind
    ' diejenigen Funktionen, welche erst gar keine Instanz der Klasse "Stringarray"
    ' erstellen. Diese Funktionen waeren eigentlich auch allein aufzurufen. Es sollten
    ' jedoch alle Funktionen einen einheitlichen Startpunkt haben und das ist diese
    ' Funktion selber.
    '
    ' Initial wird diese Variable auf einen Leerstring gestellt.
    '
    str_fkt_ergebnis = LEER_STRING
    
    If (pFunktion = FKT_CSV_2_ZEILE) Then
        '
        ' Funktion "CSV 2 Zeile"
        ' Nach jedem Trennzeichen aus dem Parameter "pOptString1" wird ein
        ' Zeilenumbruch eingefuegt.
        '
        ' Dieses wird einmal mit Loeschung des Trennzeichens gearbeitet und
        ' einmal bleibt dass Trennzeichen selber im Ergebnisstring erhalten.
        '
        If (pOptString1 = LEER_STRING) Then
        
        Else
            
            If (m_toggle_mr_stringer_fkt) Then
                
                'str_fkt_ergebnis = Replace(pString, pOptString1, pOptString1 & MY_CHR_13_10)
                str_fkt_ergebnis = Replace(pString, pOptString1, MY_CHR_13_10 & pOptString1)
            
            Else
                
                str_fkt_ergebnis = Replace(pString, pOptString1, MY_CHR_13_10)
            
            End If
            
        End If
    
    ElseIf (pFunktion = FKT_CSV_CR) Then
        '
        ' Funktion "CSV CR"
        ' Vor oder Nach jedem Trennzeichen wird ein Zeilenumbruch eingefuegt.
        ' Eigentlich aehnlich der Funktion "FKT_CSV_2_ZEILE", welche aber das
        ' Trennzeichen selber einmal mit ins Ergebnis aufnimmt oder nicht.
        '
        If (pOptString1 = LEER_STRING) Then
        
        Else
            
            If (m_toggle_mr_stringer_fkt) Then
                
                str_fkt_ergebnis = Replace(pString, pOptString1, pOptString1 & MY_CHR_13_10)
            
            Else
                
                str_fkt_ergebnis = Replace(pString, pOptString1, MY_CHR_13_10 & pOptString1)
            
            End If
            
        End If
    
    ElseIf (pFunktion = FKT_STRING_LIT) Then
        '
        ' Funktion "String Literale"
        ' Die Funktion fuer die Ermittlung der String-Literale ist
        ' eine eigenstaendige Funktion und wird nur aufgerufen.
        '
        str_fkt_ergebnis = getStringLitKonst(pString)
    
    ElseIf (pFunktion = FKT_STRING_REMOVE) Then
        
        '
        ' Funktion "Remove"
        ' Die selektierte Zeichenkette wird aus dem Eingabestring geloescht.
        '
        str_fkt_ergebnis = Replace(pString, Mid(pString, pSelStart + 1, pSelLength), LEER_STRING)
    
    ElseIf (pFunktion = FKT_EXTRAHIERE_WORTE) Then
    
        '
        ' Funktion "extrahiere Worte"
        '
        ' Das Funktionsergebnis wird ueber die Funktion "extrahiereWoerter" erstellt.
        ' Abwechselnd werden der Funktion die Ergebniszeilenlaengen von 400 und
        ' 3 Zeichen uebergeben. Das fuehrt dazu, dass die Worte einmal seperat in
        ' einer Zeile stehen und einmal in einem Block von 400 Zeichen stehen.
        '
        If (m_toggle_mr_stringer_fkt) Then
            
            str_fkt_ergebnis = extrahiereWoerter(pString, pOptString1, 400)
        
        Else
            
            str_fkt_ergebnis = extrahiereWoerter(pString, pOptString1, 3)
        
        End If
    
    ElseIf (pFunktion = FKT_FORMAT_TXT) Then
    
        '
        ' Funktion "Format Text"
        '
        ' Formatiert den Text auf 55 oder 80 Stellen Breite.
        '
        If (m_toggle_mr_stringer_fkt) Then
        
            str_fkt_ergebnis = getStringMaxCols(pString, 55, LEER_STRING, MY_CHR_13_10)
            
        Else
        
            str_fkt_ergebnis = getStringMaxCols(pString, 80, LEER_STRING, MY_CHR_13_10)
        
        End If
    
    Else
        '
        ' #####################################################################################
        ' START - VORBEREITENDE ANWEISUNGEN FUER DIER HAUPTSCHLEIFE
        ' #####################################################################################
        '
        
        '
        ' Zeilenumbruchzeichen
        ' Aus der Eingabe wird das benutzte Zeilenumbruchzeichen ermittelt.
        ' Dieses wird zum suchen und ersetzen innerhalb dieser Funktion benoetigt.
        '
        zeichen_zeilenumbruch = getBenutztesChr13(pString)
        
        '
        ' Erstellung String-Array
        ' Der uebergebene Text aus dem Parameter "pString" wird ueber die
        ' Funktion "startMultiline" in eine Instanz der Klasse clsstringArray
        ' ueberfuehrt. Wird von der Funktion "nothing" zurueckgegeben ist die
        ' Funktion "startMrStringer" beendet.
        '
        Set cls_string_array = startMultiline(pString)
        
        If (cls_string_array Is Nothing) Then
        '
        ' keine Aktionen machen, wenn das String-Array-Objekt nicht gesetzt ist.
        '
        Else
            '
            ' Voreinstellung dass ein Schleifendurchlauf notwendig ist. Es gibt
            ' Funktionen, welche komplett durch die Instanz "clsStringArray"
            ' durchgefuehrt werden koennen. Bei diesen Funktionen soll der hier
            ' nachgelagerte Durchlauf durch alle Strings unterbunden werden.
            ' Diese Aufgabe uebernimmt die Variable "knz_schleifen_durchlauf".
            '
            knz_schleifen_durchlauf = True
            
            '
            ' Verarbeitung Selektion
            ' Es wird geprueft, ob eine Position in "pSelStart" vorhanden ist.
            ' Eine Selektion liegt vor, wenn der Parameter "pSelStart" groesser 0 ist.
            '
            ' TEILWEISE FALSCH
            ' Eine Selektion liegt erst vor, wenn der Parameter "pSelLength" auch
            ' groesser als 0 ist.
            '
            If (pSelStart > POSITION_0) Then
                
                '
                ' Zeilenumbruch vor "pSelStart"
                ' Liegt eine Startposition vor, muss dessen relative Startposition
                ' zu dem letzten Zeilenumbruch ermittelt werden. Eine Markierung
                ' gilt nur fuer eine Zeile, nicht fuer den gesamten Text aus "pString".
                '
                ' Es wird das letzte Zeilenumbruchszeichen vor der Startposition gesucht.
                '
                temp_long_1 = getLetztePositionVorPos(pString, getBenutztesChr13(pString), pSelStart)
                
                '
                ' Korrektur Relative-Startposition
                ' Wird ein Zeilenumbruch gefunden, wird dessen absolute Position
                ' von der Selektionsstartposition abgezogen.
                '
                ' Wird kein Zeilenumbruch gefunden (Markierung befindet sich in
                ' Zeile 1), wird auf die Selektionsstartposition eine Position
                ' draufgerechnet.
                '
                ' 1234567890 1234567890 1234567890 1234567890 1234567890 1234567890
                ' 1234567890 1234567890 1234567890 1234567890 1234567890 1234567890
                '
                If (temp_long_1 > POSITION_0) Then
                
                    ab_position = pSelStart - temp_long_1
                
                Else
                    
                    ab_position = pSelStart + 1
                
                End If
                
                '
                ' Bestimmung Bis-Position
                ' Auf die Ab-Position wird die Selektionslaenge hinzugezaehlt und
                ' abschliessend wieder eine Stelle abgezogen (da die Startposition
                ' schon selber enthalten ist).
                '
                bis_position = (ab_position + pSelLength) - 1
                
                '
                ' Kennzeichen "knz_benutze_markierung"
                ' Das Kennzeichen ist TRUE, wenn die Ab-Position groesser gleich 0
                ' ist und die Bis-Position gleich oder groesser ist.
                '
                knz_benutze_markierung = (ab_position >= POSITION_0) And (bis_position >= ab_position)
            
            Else
                
                '
                ' Liegt keine Startposition einer Selektion vor, werden die
                ' Variablen Ab- und Bis-Position auf 0 und das Kennzeichen
                ' fuer die Verwendung der Markierung auf false gestellt.
                '
                ab_position = POSITION_0
                
                bis_position = POSITION_0
                
                knz_benutze_markierung = False
            
            End If
            
            If (pFunktion = FKT_CSV_REPLACE_MARKIERUNG_MIT_CSV) Then
                
                '
                ' Funktion Markierung mit CSV-String ersetzen
                '
                ' Der Suchstring ist die Markierung
                ' Der Suchstring wird in der Variablen "inhalt_markierung" gespeichert.
                ' Ist der Parameter "pSelLength" groesser als 0, liegt eine Markierung vor.
                ' Es wird aus dem Eingabetext der Suchtext herausgelesen.
                '
                ' Liegt keine Markierung vor, oder die Markierung ist ein Leerstring,
                ' ist das Ergebnis dieser Funktion der Eingabetext selber. Ist keine
                ' Markierung vorhanden, ist keine Ersetzung notwendig.
                '
                ' Die Erstzungen werden global mit der Replace-Funktion gemacht.
                ' Ein weiterer Schleifendurchlauf ist nicht notwendig.
                '
                
                If (pSelStart > POSITION_0) Then
                
                    inhalt_markierung = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                End If
                
                If (inhalt_markierung <> LEER_STRING) Then
                
                    str_fkt_ergebnis = Replace(pString, inhalt_markierung, pOptString1)
                    
                Else
                
                    str_fkt_ergebnis = pString
                
                End If
                
                knz_schleifen_durchlauf = False
            
            ElseIf (pFunktion = FKT_SORTIEREN_ALPHABETH) Then
            
                '
                ' Funktion "Sortierung Alphabeth"
                '
                ' Das Sortieren der Zeilen wird in der Klasse clsStringArray
                ' gemacht. Ein weiterer Schleifendurchlauf ist nicht notwendig.
                '
                Call cls_string_array.startSortierung(1, m_toggle_mr_stringer_fkt, False, knz_benutze_markierung, ab_position, bis_position)
                
                str_fkt_ergebnis = cls_string_array.toString(zeichen_zeilenumbruch, True)
                
                knz_schleifen_durchlauf = False
                                
            ElseIf (pFunktion = FKT_MARKIERE_VORNE_UND_HINTEN) Then 'ISXF
            
                temp_string_1 = IIf(Len(pOptString1) = 0, TRENN_STRING_7, pOptString1)
                temp_string_2 = IIf(Len(pOptString2) = 0, TRENN_STRING_8, pOptString2)

            ElseIf (pFunktion = FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE) Then 'ISXF
            
                temp_string_1 = IIf(Len(pOptString1) = 0, TRENN_STRING_6, pOptString1)
                temp_string_2 = IIf(Len(pOptString2) = 0, TRENN_STRING_7, pOptString2)
                pOptString3 = IIf(Len(pOptString3) = 0, TRENN_STRING_8, pOptString3)

            ElseIf (pFunktion = FKT_SORTIEREN_ZUFALL) Then
            
                '
                ' Funktion "Umstellung Zufall"
                '
                ' Die zufaellige Umstellung der Zeilen wird in der Klasse
                ' "clsStringArray" gemacht. Ein weiterer Schleifendurchlauf
                ' ist nicht notwendig.
                '
                Call cls_string_array.startZufallsUmsortierung
                
                str_fkt_ergebnis = cls_string_array.toString(zeichen_zeilenumbruch, True)
                
                knz_schleifen_durchlauf = False
            
            ElseIf (pFunktion = FKT_MARKIERE_WORT) Then
                
                '
                ' Funktion Markiere Wort
                '
                ' Suchwort
                ' Ermittlung der zu markierenden Zeichenkette. Das ist der
                ' String aus der Markierung, welches ein Wort, oder auch
                ' mehrere Zeichen sein koennen.
                '
                inhalt_markierung = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                '
                ' Das Suchwort wird mit dem Trennstring 8 vor oder hinter
                ' dem Suchwort versehen und im Eingabetext ersetzt.
                '
                If (m_toggle_mr_stringer_fkt) Then
                
                    str_fkt_ergebnis = Replace(pString, inhalt_markierung, inhalt_markierung & TRENN_STRING_8)
                    
                Else
                
                    str_fkt_ergebnis = Replace(pString, inhalt_markierung, TRENN_STRING_8 & inhalt_markierung)

                End If
                
                '
                ' Es muss bei der Funktion "FKT_MARKIERE_WORT" kein weiterer
                ' Schleifendurchlauf gemacht werden. Die Variable "knz_schleifen_durchlauf"
                ' wird auf FALSE gestellt.
                '
                knz_schleifen_durchlauf = False
                
            ElseIf (pFunktion = FKT_SORTIEREN_LAENGE) Then
            
                '
                ' Funktion "Sortieren nach Zeilenlaenge"
                '
                ' In der Klasse clsStringArry wird die Funktion "sortieren" mit
                ' dem Aktionscode 234 aufgerufen. Das Funktionsergebnis wird
                ' direkt aus dem Stringarry gelesen. Es muss kein weiterer
                ' Schleifendurchlauf gemacht werden.
                '
                Call cls_string_array.startSortierung(234, m_toggle_mr_stringer_fkt, False, knz_benutze_markierung, ab_position, bis_position)
                
                str_fkt_ergebnis = cls_string_array.toString(zeichen_zeilenumbruch, True)
                
                knz_schleifen_durchlauf = False
            
            ElseIf (pFunktion = FKT_SORTIEREN_DATUM) Then
            
                '
                ' Funktion "Sortieren nach Datum"
                '
                ' Das Sortieren nach Datum wird von der String-Array-Klasse gemacht.
                ' Es muss kein weiterer Schleifendurchlauf gemacht werden.
                '
            
                If (ab_position = POSITION_0) Then
                    
                    ab_position = 1
                
                End If
                
                Call cls_string_array.startSortierung(1, m_toggle_mr_stringer_fkt, True, True, ab_position, bis_position)
                
                str_fkt_ergebnis = cls_string_array.toString(zeichen_zeilenumbruch, True)

                knz_schleifen_durchlauf = False
                
            ElseIf (pFunktion = FKT_LEERZEILEN_EINFUEGEN) Then
            
                '
                ' Vermeidet dass bei der ersten Zeile ein Zeilenumbruch hinzugefuegt wird.
                '
                temp_long_1 = 0
                
            ElseIf (pFunktion = FKT_CSV_VB_KONVERTER) Then
            
                temp_string_3 = "    If"
            
            ElseIf (pFunktion = FKT_STRING_ERST) Then
            
                If (m_toggle_mr_stringer_fkt) Then
                    
                    temp_string_1 = "\"""
    
                Else
                    temp_string_1 = """"""
                
                End If
            
            ElseIf (pFunktion = FKT_GENERATOR_STRING_IT) Then
            
                '
                ' Vorbereitung Funktion StringIt
                '
            
                m_zaehler_string_it = m_zaehler_string_it + 1
                
                If (m_zaehler_string_it = 2) Then
                    '
                    ' Verkettung als Java-String
                    '
                    temp_string_1 = "j_str += ""\n"
                    temp_string_2 = """;"
                    temp_string_3 = "\"""
    
                ElseIf (m_zaehler_string_it = 3) Then
                    '
                    ' Stringbuffer append
                    '
                    temp_string_1 = "str_buf.append( """
                    temp_string_2 = """ );"
                    temp_string_3 = "\"""
    
                Else
                    '
                    ' Verkettung als VB normaler Variablenname
                    '
                    temp_string_1 = "vb_str = vb_str & """
                    temp_string_2 = """"
                    temp_string_3 = """"""
                    
                    m_zaehler_string_it = 1
                
                End If
            
            ElseIf (pFunktion = FKT_MASKIERE_ANFZEICHEN) Then
                '
                ' Funktion "Maskiere Anfuehrungszeichen"
                '
                ' Die Anfuehrungszeichen werden abwechselnd einmal fuer Java und einmal fuer VB maskiert.
                ' Es wird nicht global ein replace gemacht, da in diesem Fall die Markierungsfunktion
                ' nicht funktionieren wuerde.
                '
                temp_string_1 = """"
                
                If (m_toggle_mr_stringer_fkt) Then
                
                    temp_string_2 = "\"""
                    
                Else
                
                    temp_string_2 = """"""
                
                End If

            ElseIf ((pFunktion = FKT_ERSTELLE_XML) Or (pFunktion = FKT_ERSTELLE_XML_2)) Then
        
                '
                ' Funktion XML-Erstellung
                ' Umdrehen der boolschen Variable (mit / ohne Vorlauf)
                ' Parameterkennzeichen bestimmt ob volle oder nur einzelne TAGS
                '
                If (m_toggle_mr_stringer_fkt) Then
                    
                    temp_string_1 = "<"
                    temp_string_2 = "></"
                    temp_string_3 = ">"
            
                Else
                
                    temp_string_1 = "<"
                    temp_string_2 = " />"
                    temp_string_3 = LEER_STRING
                
                End If
            
            ElseIf (pFunktion = FKT_ERSTELLE_BLOCK) Then
                
                '
                ' Vorbereitung Funktion Erstelle Block
                '
                ' Die Laenge der Zeilen wird in der Variablen "temp_long_1" gespeichert.
                ' Die Zeilenbreite wird ueber die Selektionslaenge uebergeben.
                ' Die Zeilenbreite betraegt minimal 30 und maximal 1000 Zeichen.
                '
                
                If (pSelLength <= 30) Then
                    
                    temp_long_1 = 30
                    
                ElseIf (pSelLength > 1000) Then
                    
                    temp_long_1 = 1000
                
                Else
                    
                    temp_long_1 = pSelLength
                
                End If
                
            ElseIf ((pFunktion = FKT_STRING_SPLIT) Or (pFunktion = FKT_KONSTANTEN_UEBER_SPLIT)) Then
                '
                ' Vorbereitung Funktion Split
                '
                ' Ein Schleifendurclauf muss nur gemacht werden, wenn die Cursorposition groesser als 0 ist
                '
                knz_schleifen_durchlauf = ab_position > POSITION_0
                
                '
                ' In der Variablen "inhalt_markierung" wird ein eventuell markierter
                ' Trennstring gespeichert.
                '
                inhalt_markierung = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                '
                ' Die Variable "temp_long_1" gibt die Position in der Zeile vor, an welcher
                ' die Zeilen gesplittet werden sollen. Ist keine Markierung vorhanden wird
                ' an dieser festen Position getrennt.
                '
                temp_long_1 = ab_position
            
            ElseIf ((pFunktion = FKT_GREP_DUPLIZIERE_MARKZEILEN) Or (pFunktion = FKT_GREP_PLUS_MINUS) Or (pFunktion = FKT_GREP_MARK)) Then
            
                '
                ' Vorbereitung Funktion Grep Minus, Grep Mark, Duplizieren markierte Zeilen
                '
                
                If (pSelLength > POSITION_0) Then
                
                    '
                    ' Ist die Selektionslaenge groesser als 0, wird das zu
                    ' suchende Wort aus dem Eingabetext ermittelt.
                    '
                    inhalt_markierung = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                Else
                    '
                    ' Ist die Selektionslaenge gleich 0, wird ab dem Selektionsstart
                    ' der Anfang und das Ende des Wortes unter dem Cursor ermittelt.
                    '
                    ' Kommt dabei eine Wortlaenge grosser als 0 raus, wird das Wort
                    ' als Suchbegriff gesetzt.
                    '
                    temp_long_1 = getPosWortAnfang(pString, pSelStart)

                    temp_long_2 = getPosWortende(pString, pSelStart)
                                
                    If ((temp_long_1 > POSITION_0) And (temp_long_2 > temp_long_1)) Then
                        
                        inhalt_markierung = Mid(pString, temp_long_1, (temp_long_2 - (temp_long_1)) + 1)
                        
                    End If
    
                End If
                
                '
                ' Ein Schleifendurchlauf muss nur gestartet werden, wenn ein Suchwort vorhanden ist.
                '
                If (Len(Trim(inhalt_markierung)) = 0) Then
                
                    knz_schleifen_durchlauf = False
                    
                Else
                                        
                    knz_schleifen_durchlauf = True
                    
                    If (pFunktion = FKT_GREP_DUPLIZIERE_MARKZEILEN) Then
                        
                        '
                        ' keine weiteren Anweisungen mehr. Vermeidung, das
                        ' in den letzten ELSE-Zweig verzweigt wird
                        '
                        
                    ElseIf (m_toggle_mr_stringer_fkt) Then
                    
                        '
                        ' Markierung vorne setzen = "temp_string_2" gesetzt und "temp_string_3" ein Leerstring
                        '
                        
                        temp_string_2 = TRENN_STRING_7
                        temp_string_3 = LEER_STRING
                    
                    Else
                        '
                        ' Markierung hinten setzen = "temp_string_2" ein Leerstring und "temp_string_3" gesetzt
                        '
                        
                        temp_string_2 = LEER_STRING
                        temp_string_3 = TRENN_STRING_7
                    
                    End If
                    
                End If
            
            ElseIf (pFunktion = FKT_GREP_WORT) Then
                
                '
                ' Vorbereitung Funktion Grep Wort
                '
                ' In "temp_string_1" wird das zu suchende Wort gespeichert.
                ' Das zu suchende Wort ist die Markierung aus der Eingabe.
                '
            
                inhalt_markierung = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
            
            ElseIf (pFunktion = FKT_MARKIERE_VORNE_FIX) Then
            
                '
                ' Funktion "Markiere Vorne mit String"
                ' Setzen eines Strings vorne am String. Intern wird die Funktion auf
                ' die Funktion "FKT_MARKIERE_VORNE_ODER_HINTEN" gestellt.
                '
                ' Die zu setzende Zeichenfolge wird im Parameter "pOptString1" erwartet.
                ' Ist die Laenge des Parameters gleich 0, wird der Trennstring 7 genommen.
                '
            
                temp_string_1 = IIf(Len(pOptString1) = 0, TRENN_STRING_7, pOptString1)
            
                m_toggle_mr_stringer_fkt = True
                
                knz_benutze_markierung = False
                
                pFunktion = FKT_MARKIERE_VORNE_ODER_HINTEN
            
            ElseIf (pFunktion = FKT_MARKIERE_VORNE_ODER_HINTEN) Then 'ISXF
                    
                temp_string_1 = TRENN_STRING_7
            
            ElseIf ((pFunktion = FKT_GENERATOR_IF_JAVA_VB) Or (pFunktion = FKT_GENERATOR_IF_JAVA_SCRIPT)) Then
                
                '
                ' Vorbereitung Funktionen Generator If
                '
            
                temp_string_3 = "if"
                
                inhalt_markierung = getStringAbBis(pString, ab_position, bis_position)
                
                temp_long_2 = Len(inhalt_markierung)
            
            ElseIf (pFunktion = FKT_ZEILEN_ADD) Then
                
                '
                ' Vorbereitung Funktion Zeilen ADD
                '
                ' Zusammenfassung von Zeilen bis eine Anzahl von Zeilenzusammenfuehrungen erreicht ist.
                ' (10 Zeilen Eingabe ergeben 1 Zeile Ausgabe)
                '
                ' Die Variable "temp_long_1" gibt die Anzahl der zusammenzufassenden Zeilen vor.
                ' Die Variable "temp_long_2" zaehlt die zusammengefassten Zeilen.
                '
                ' Die Variable "temp_string_1" speichert einen optionalen Zusatzstring, welcher
                ' nach jeder hinzugefuegten Zeile angefuegt wird.
                '
                
                temp_long_2 = 1
                
                '
                ' Ist eine Markierung vorhanden, ist die Zeilen-Sollanzahl gleich der Anzahl der
                ' Zeilenumbruecke innerhalb der Markierung.
                '
                If (pSelStart > POSITION_0) Then
                    
                    temp_long_1 = getAnzahlVorkommen(Mid(pString, pSelStart, pSelLength), zeichen_zeilenumbruch) + 1
                
                Else
                    
                    temp_long_1 = getAnzahlVorkommen(Mid(pString, 1, pSelLength), zeichen_zeilenumbruch) + 1
                
                End If
                
                If (m_toggle_mr_stringer_fkt) Then
                    
                    temp_string_1 = pOptString1
                
                Else
                    
                    temp_string_1 = LEER_STRING
                
                End If
            
            ElseIf (pFunktion = FKT_GENERATOR_DEBUG_AUSGABE) Then
            
                '
                ' Vorbereitung Funktion Debugausgabe
                '
                ' Mit der Variablen "m_zaehler_debug_print" werden die verfuegbaren
                ' Ergebnissausgaben durchgeschaltet. Es gibt 6 Arten der Debugausgabeerstellung.
                '
            
                m_zaehler_debug_print = m_zaehler_debug_print + 1
                
                If (m_zaehler_debug_print > 6) Then
                    
                    m_zaehler_debug_print = 1
                
                End If

            ElseIf ((pFunktion = FKT_SET_TRENNZEICHEN) Or (pFunktion = FKT_SET_TRENNZEICHEN_VOR) Or (pFunktion = FKT_SET_TRENNZEICHEN_ZURUECK)) Then
                
                '
                ' Vorbereitung Funktion "SetTrennzeichen"
                '
                ' Der Funktion kann mit dem Parameter "pOptString1" ein zu setzender
                ' Trennstring uebergeben werden. Ist der Parameter ein Leerstring,
                ' wird abwechselnd ein internes Trennzeichen genommen.
                '
                
                temp_string_3 = IIf(Len(pOptString1) = 0, IIf(m_toggle_mr_stringer_fkt, TRENN_STRING_6, TRENN_STRING_9), pOptString1)

            ElseIf ((pFunktion = FKT_CLIP_POSITION) Or (pFunktion = FKT_CLIP_GET_TEXT) Or (pFunktion = FKT_CLIP_ENTFERNE_TEXT)) Then
                    
                '
                ' Vorbereitung Funktion Clip
                '
                
                If ((pSelLength = Len(pString)) Or (knz_benutze_markierung = False)) Then
                
                    '
                    ' Ist der gesamte Text markiert, muss kein Schleifendurchlauf
                    ' gestartet werden. Der Aufrufer bekommt den Eingabestring zurueck.
                    '
                
                    str_fkt_ergebnis = pString
                    
                    knz_schleifen_durchlauf = False
    
                End If
                
                If (pFunktion = FKT_CLIP_POSITION) Then
                
                    '
                    ' Ist die Funktion gleich "Clip Position", wird die von der
                    ' Schleife auszufuehrende Aktion umgeschaltet. Je nach dem
                    ' Wert von "m_toggle_mr_stringer_fkt" ist das einmal die
                    ' Funktion "FKT_CLIP_ENTFERNE_TEXT" oder "FKT_CLIP_GET_TEXT".
                    '

                    If (m_toggle_mr_stringer_fkt) Then
                    
                        pFunktion = FKT_CLIP_ENTFERNE_TEXT
                    
                    Else
                    
                        pFunktion = FKT_CLIP_GET_TEXT
                    
                    End If
                
                End If
            
            ElseIf (pFunktion = FKT_CALC_SUMME) Then
            
                '
                ' Vorbereitung Funktion "Calc Summe"
                '
                ' Die verwendeten Variablen fuer die Aufsummierung werden auf 0 gestellt.
                '
            
                temp_long_1 = 0
                temp_double_1 = 0#
                temp_double_2 = 0
            
            ElseIf ((pFunktion = FKT_AUSRICHTER_POSITION) Or (pFunktion = FKT_AUSRICHTER_STRING)) Then
            
                '
                ' Vorbereitung Funktion Ausrichter
                '
                ' Fuer die Ausrichter-Funktion muss in einem ersten Durchlauf die
                ' maximale Ausdehnung des "Suchstrings" ermittelt werden. Der
                ' Suchstring ist dass markierte Wort, oder der String aus dem
                ' Parameter "pOptString1". Die ermittelte maximale Position
                ' wird in der Variablen "temp_long_2" gespeichert.
                '
            
                If (pFunktion = FKT_AUSRICHTER_STRING) Then
                    
                    temp_string_1 = pOptString1
                    
                    pFunktion = FKT_AUSRICHTER_POSITION
                
                Else
                    
                    temp_string_1 = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                End If
            
                '
                ' Gibt es einen Suchstring, wird die Suchschleife fuer die maximale
                ' Position des Suchstrings gestartet. Ist kein Ausrichtungstext vorhanden,
                ' bekommt der Aufrufer den Eingabestring unbehandelt zurueck.
                '
                knz_schleifen_durchlauf = Len(temp_string_1) > 0
                
                If (knz_schleifen_durchlauf) Then
                    
                    zeilen_anzahl = cls_string_array.getAnzahlStrings
                    
                    '
                    ' Die Speichervariable fuer die maximale Ausdehung wird auf 1 gesetzt.
                    '
                    temp_long_2 = 1
                    
                    zeilen_zaehler = 1
                    
                    '
                    ' Suchschleife fuer die maximale Position des Suchstrings.
                    '
                    While (zeilen_zaehler <= zeilen_anzahl)
                    
                        '
                        ' Aktuelle zeile aus dem Stringarray holen
                        '
                        aktuelle_zeile = cls_string_array.getString(zeilen_zaehler)
                        
                        '
                        ' Pruefung: aktuelle Zeile ungleich Leerstring ?
                        '
                        ' Ist die aktuelle Zeile ein Leerstring, muss der
                        ' Suchstring nicht gesucht werden.
                        '
                        If (aktuelle_zeile <> LEER_STRING) Then
                            
                            '
                            ' Ist die aktuelle Zeile ungleich einem Leerstring,
                            ' wird in dieser Zeile die Position des Suchstrings
                            ' ermittelt.
                            '
                            temp_long_1 = InStr(aktuelle_zeile, temp_string_1)
                            
                            '
                            ' Ist die Position groesser als die bisherige maximale
                            ' Position, ist eine neue Max-Position gefunden worden.
                            ' Die neue Max-Position wird in der Variablen "temp_long_2"
                            ' gespeichert.
                            '
                            If (temp_long_1 > temp_long_2) Then
            
                                temp_long_2 = temp_long_1
                                
                            End If
                        
                        End If
                    
                        zeilen_zaehler = zeilen_zaehler + 1
                    
                    Wend
                    
                    '
                    ' In der Variablen "temp_string_3" wird ein String aus Leerzeichen mit der
                    ' Laenge der maximalen Ausdehnung gespeichert.
                    '
                    temp_string_3 = String(temp_long_2, " ") & "  "
    
                End If
            
            ElseIf (pFunktion = FKT_ZEILEN_ZAEHLER) Then
            
                '
                ' Vorbereitung Funktion Zeilen Zaehler
                '
                ' In der Variablen "temp_long_1" wird die Anzahl der gewuenschten
                ' vorlaufenden 0en gespeichert. Bei einer Markierung von mehr als
                ' 1 Stelle und weniger als 100 Stellen, wird die Markierungslaenge
                ' genommen. Ansonsten werden 6 Stellen genommen.
                '
                ' Bei jedem zweiten Aufruf, wird die Anzahl auf 0 gestellt.
                '
                If (m_toggle_mr_stringer_fkt) Then
                
                    temp_long_1 = 0 ' Keine fuehrenden 0en
                    
                ElseIf ((pSelLength >= 1) And (pSelLength <= 100)) Then
                    
                    temp_long_1 = pSelLength
                    
                Else
                        
                    temp_long_1 = 6
                    
                End If
            
            End If
            
            '
            ' #####################################################################################
            ' START - HAUPTSCHLEIFE UEBER ALLE ZEILEN, WENN NOTWENDIG
            ' #####################################################################################
            '
            
            '
            ' Pruefung: Haupt-Schleifendurchlauf machen?
            '
            ' Start der Hauptverarbeitungsschleife.
            '
            If (knz_schleifen_durchlauf) Then
            
                '
                ' Anzahl der insgesamt vorhandenen Zeilen lesen
                '
                zeilen_anzahl = cls_string_array.getAnzahlStrings
                
                '
                ' Zeilenzaehler auf 1 stellen.
                '
                zeilen_zaehler = 1
                
                '
                ' Schleifendurchlauf von 1 bis zu der Anzahl der vorhandenen Zeilen.
                ' Es ist kein Endlossschleifen-Verhinderungszaehler vorhanden.
                '
                While (zeilen_zaehler <= zeilen_anzahl)
                
                    '
                    ' Aktuelle Zeile
                    ' Die aktuelle Zeile wird per Index aus dem Zeilenobjekt gelesen
                    ' und in der Variablen "aktuelle_zeile" gespeichert.
                    '
                    aktuelle_zeile = cls_string_array.getString(zeilen_zaehler)
                    
                    '
                    ' Variable "akt_zeile_mark"
                    ' Soll die Markierung benutzt werden, wird in dieser Variablen nur der
                    ' sich ergebende String aus der Ab- bis zur Bis-Position gespeichert.
                    '
                    ' Soll die Markierung nicht benutzt werden, wird die aktuelle
                    ' Zeile gespeichert.
                    '
                    If (knz_benutze_markierung) Then
                        
                        akt_zeile_mark = getStringAbBis(aktuelle_zeile, ab_position, bis_position)
                    
                    Else
                        
                        akt_zeile_mark = aktuelle_zeile
                    
                    End If
                    
                    '
                    ' Funktionsbestimmung
                    ' Ueber If-Abfragen wird die auszufuerhende Aktion ermittelt.
                    '
                    If (pFunktion = FKT_GREP_WORT) Then
                    
                        '
                        ' Funktion "Grep Wort"
                        '
                        ' Bedingung ist, dass die aktuelle Zeile kein Leerstring ist.
                        ' Aus einem Leerstring kann kein Wort rausgezogen werden.
                        '
                        ' Wurde der Wortanfang mittels Markierung vorgegeben, wird die
                        ' Funktion "getGrepSuchwort" aufgerufen, welche alle Worte aus
                        ' der aktuellen Zeile raussucht, welche mit der vorgegebenen Zeichen-
                        ' folge starten.
                        '
                        ' Liegt keine Markierung vor, wird von der aktuellen Position ein
                        ' Anfang und ein Ende gesucht (... ein Leerzeichen, oder Satzzeichen).
                        ' Wenn ein Wort gefunden werden konnte, wird dieses extrahiert.
                        '
                        If (aktuelle_zeile <> LEER_STRING) Then
                        
                            '
                            ' Ergebnisvariable "temp_string_3" wird auf einen Leerstring gesetzt.
                            ' Gleichbedeudent mit "es gibt keine Ergebnisse aus der aktuellen Zeile".
                            '
                            temp_string_3 = LEER_STRING
                            
                            If (pSelStart > POSITION_0) Then
                            
                                temp_string_3 = getGrepSuchwort(aktuelle_zeile, inhalt_markierung, zeichen_zeilenumbruch)
                            
                            Else
                            
                                temp_long_2 = getPosWortAnfang(aktuelle_zeile, ab_position)
                                
                                temp_long_3 = getPosWortende(aktuelle_zeile, ab_position)
                                
                                If ((temp_long_3 > POSITION_0) And (temp_long_2 > POSITION_0) And (temp_long_3 > temp_long_2)) Then
                                    
                                    temp_string_3 = Mid(aktuelle_zeile, temp_long_2, (temp_long_3 - (temp_long_2)) + 1)
                                
                                End If
                            
                            End If
                            
                            If (temp_string_3 <> LEER_STRING) Then
                            
                                If (zeilen_zaehler = 1) Then
                                    
                                    str_fkt_ergebnis = str_fkt_ergebnis & temp_string_3
                                
                                Else
                                    
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & temp_string_3
                                
                                End If
                                
                            End If
                            
                        End If
                 
                    ElseIf (pFunktion = FKT_AUSRICHTER_POSITION) Then
                    
                        '
                        ' Funktion Ausrichter
                        ' In der aktuellen Zeile wird die Positon des Suchbegriffs gesucht.
                        '
                        temp_long_1 = InStr(aktuelle_zeile, temp_string_1)
                        
                        '
                        ' Pruefung: Suchbegriff gefunden?
                        ' Bedingung: Position groesser als 0
                        '
                        ' Wenn der Suchbegriff in der aktuellen Zeile gefunden wird, werden
                        ' entsprechend viele Leerzeichen bis zur ermittelten Max-Ausdehnung
                        ' eingefuegt. Die MaxPosition wurde weiter oben ermittelt.
                        '
                        ' Wird der Suchbegriff nicht gefunden, muss auch nichts eingefuegt werden.
                        '
                        If (temp_long_1 > POSITION_0) Then
                        
                            Call cls_string_array.setString(zeilen_zaehler, Left(aktuelle_zeile, temp_long_1 - 1) & Left(temp_string_3, temp_long_2 - temp_long_1) & Mid(aktuelle_zeile, temp_long_1))
                            
                        End If
                    
                    ElseIf (pFunktion = FKT_GREP_PLUS_MINUS) Then
                        '
                        ' Funktion Grep + und Grep -
                        '
                        ' Grep + = alle Zeilen mit dem Suchwort
                        ' Grep - = alle Zeilen ohne dem Suchwort
                        '
                        ' Pruefung: wird das Suchwort in der aktuellen Zeile gefunden?
                        '
                        ' Ist das Suchwort enthalten, wird die aktuelle Zeile aufgenommen, wenn es sich um
                        ' die Funktion Grep+ handelt.
                        '
                        ' Ist das Suchwort nicht enthalten, wird die aktuelle Zeile aufgenommen, wenn es sich
                        ' um die Funktion Grep- handelt.
                        '
                        ' Diese Funktion reduziert die Anzahl der Ergebniszeilen.
                        ' Das Ergebnis wird in der Varaibeln "str_fkt_ergebnis" aufgebaut.
                        '
                        If (InStr(aktuelle_zeile, inhalt_markierung) > POSITION_0) Then
                            
                            If (pKennzeichen1) Then
                                
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & aktuelle_zeile
                            
                            End If
                            
                        Else
                        
                            If (pKennzeichen1 = False) Then
                                
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & aktuelle_zeile
                            
                            End If
                        
                        End If
                    
                    ElseIf (pFunktion = FKT_GREP_DUPLIZIERE_MARKZEILEN) Then
                    
                        '
                        ' Funktion "Grep und Dupliziere"
                        ' Es wird in der aktuellen Zeile nachgesehen, ob der zu suchende String enthalten ist.
                        ' Ist der String enthalten, wird die Zeile dupliziert (nicht Zeilenweise, sondern der
                        ' String wird 2 mal hintereinandergehaengt).
                        '
                        
                        If (InStr(aktuelle_zeile, inhalt_markierung) > POSITION_0) Then
                            
                            Call cls_string_array.setString(zeilen_zaehler, temp_string_2 & aktuelle_zeile & temp_string_3 & aktuelle_zeile)
                            
                        End If
                    
                    ElseIf (pFunktion = FKT_GREP_MARK) Then
                    
                        '
                        ' Funktion "Grep Mark"
                        '
                        ' In der aktuellen Zeile wird das Wort aus der Makrierung gesucht.
                        '
                        ' Wird das Wort gefunden, wird nachgesehen, ob die Zeile markiert werden soll.
                        ' Soll eine Markierung angebracht werden, wird die Zeile mit der Markierung versehen.
                        '
                        ' Wird das Wort nicht gefunden, wird die Markierung nur angebracht, wenn die
                        ' Markierung im Negativen Fall angebracht werden soll.
                        '
                        ' Die Markierug besteht ist abwechselnd in "temp_string_2" oder "temp_string_3" enthalten.
                        '

                        If (InStr(aktuelle_zeile, inhalt_markierung) > POSITION_0) Then
                        
                            If (pKennzeichen1) Then
                                
                                Call cls_string_array.setString(zeilen_zaehler, temp_string_2 & aktuelle_zeile & temp_string_3)
                                
                            End If
                            
                        Else
                        
                            If (pKennzeichen1 = False) Then
                                
                                Call cls_string_array.setString(zeilen_zaehler, temp_string_2 & aktuelle_zeile & temp_string_3)
                            
                            End If
                        
                        End If
                    
                    ElseIf (pFunktion = FKT_MARKIERE_CSV_VORNE_ODER_HINTEN) Then 'ISXF
                        '
                        ' Funktion "CSV Markiere vorne oder hinten"
                        '
                        ' Die aktuelle Zeile wird abwechselnd vorne oder hinten mit
                        ' dem uebergebenen String im Parameter "pOptString1" versehen.
                        ' Soll die Markierung benutzt werden, wird in der aktuellen
                        ' Zeile der "temp_string_3" ersetzt.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            temp_string_3 = pOptString1 & akt_zeile_mark
                        
                        Else
                        
                            temp_string_3 = akt_zeile_mark & pOptString1
                            
                        End If
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)
                        
                    
                    ElseIf (pFunktion = FKT_STRING_ERST) Then 'ISXF
                        
                        '
                        ' Funktion "String Erstellung"
                        ' Generator fuer die Stringerzeugung in Programmiersprachen.
                        '
                        ' Beruecksichtigt die Markierung.
                        '
                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            temp_string_3 = Replace(Trim(akt_zeile_mark), """", temp_string_1)
                            
                        End If
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, """" & temp_string_3 & """")
                        
                    ElseIf (pFunktion = FKT_MARKIERE_VORNE_ODER_HINTEN) Then 'ISXF
                    
                        '
                        ' Funktion "Markiere vorne oder hinten"
                        '
                        ' Die aktuelle Zeile wird abwechselnd vorne oder hinten mit
                        ' einem Suchstring versehen. Der hinzugefuegte Suchstring
                        ' ist in der Variablen "temp_string_1" enhalten.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            temp_string_3 = temp_string_1 & akt_zeile_mark
                        
                        Else
                        
                            temp_string_3 = akt_zeile_mark & temp_string_1
                            
                        End If
                        
                        '
                        ' Soll die Markierung genutzt werden, wird im aktuellem String,
                        ' ab der Ab-Position bis zur Bis-Position der String aus
                        ' "temp_string_1" ersetzt.
                        '
                        If (knz_benutze_markierung) Then
                        
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)
                        
                    ElseIf (pFunktion = FKT_MARKIERE_VORNE_UND_HINTEN) Then 'ISXF
                        '
                        ' Funktion "Markiere vorne und hinten"
                        '
                        ' Es wird vorne und hinten ein Suchzeichen gesezt.
                        ' Das kann auf die gesamte Zeile oder aber nur auf den Markierungsbereich erfolgen.
                        '
                        temp_string_3 = temp_string_1 & akt_zeile_mark & temp_string_2
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)

                    ElseIf (pFunktion = FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE) Then 'ISXF
                        '
                        ' Funktion "Markierung und Dopplung"
                        ' Die aktuelle Zeile/Markierung wird gedoppelt. Die sich ergebenden
                        ' beiden Spalten werden durch Trennzeichen voneinander getrennt.
                        '
                        temp_string_3 = temp_string_1 & akt_zeile_mark & temp_string_2 & akt_zeile_mark & pOptString3
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)
                    
                    ElseIf (pFunktion = FKT_STRING_UMDREHEN) Then 'ISXF
                        '
                        ' Funktion "Umdrehen"
                        ' Dreht die Zeichen der aktuellen Zeile oder Markierung um.
                        '
                        temp_string_3 = StrReverse(akt_zeile_mark)

                        If (knz_benutze_markierung) Then
                    
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                        
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)

                    ElseIf (pFunktion = FKT_MARKIERE_STR_VORNE_UND_HINTEN) Then
                        '
                        ' Funktion "Markiere vorne und hinten mit String"
                        '
                        ' Die aktuelle Zeile wird vorne mit dem Optstring 1 und hinten mit
                        ' dem Optstring 2 versehen. Eine Verwendung der Markierung ist
                        ' nicht vorhanden.
                        '
                        Call cls_string_array.setString(zeilen_zaehler, pOptString1 & akt_zeile_mark & pOptString2)
        
                    ElseIf (pFunktion = FKT_STRING_LAENGE_AUSGEBEN) Then
                    
                        '
                        ' Funktion "Stringlaenge"
                        '
                        ' Die Laenge der aktuellen Zeile wird in der Variablen "temp_long_1" gespeichert.
                        ' In der Variablen "temp_long_2" wird die bisherige Gesamtlaenge gespeichert.
                        ' Mit der bisherigen Gesamtlaenge kann ermittelt werden, wann eine bestimmte
                        ' Stringlaenge erreicht worden ist. Die beiden Laengenangaben werden im
                        ' Stringarray gespeichert.
                        '
                        
                        temp_long_1 = Len(akt_zeile_mark)
                        
                        temp_long_2 = temp_long_2 + temp_long_1
                         
                        Call cls_string_array.setString(zeilen_zaehler, temp_long_1 & " " & TRENN_STRING_6 & " " & temp_long_2)
                    
                    ElseIf (pFunktion = FKT_TRIM_STRING_VORNE_UND_HINTEN) Then 'ISXF
                        '
                        ' Funktion "Trim"
                        '
                        ' Auf jede Zeile wird ein Trim ausgefuehrt. Der getrimmte String, wird in
                        ' der Variablen "temp_string_3" gespeichert.
                        '
                        ' Soll nur innerhalb der Markierung getrimmt werden, wird das Ergebnis der
                        ' Trim-Funktion nur innherhalb der Positionen Ab und Bis ausgetauscht.
                        '
                        ' Am Ende wird die Zeile im Stringarray gesetzt.
                        '
                        temp_string_3 = Trim(akt_zeile_mark)
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)
                    
                    ElseIf (pFunktion = FKT_MASKIERE_ANFZEICHEN) Then 'ISXF
                        
                        '
                        ' Funktion "Maskiere Anfuerhungszeichen"
                        '
                        ' Der Suchstring ist in "temp_string_1" gespeichert.
                        ' Der Ersatzstring ist in "temp_string_2" gespeichert.
                        '
                        ' In der aktuellen Zeile wird die Erstzung durchgefuehrt.
                        '
                        ' ACHTUNG: Es koennte auch gleich im gesamten Eingabestring
                        '          diese Ersetzung gemacht werden. In diesem Fall
                        '          wuerde das Ersetzen innerhalb der Markierung aber
                        '          nicht mehr funktionieren.
                        '
                        ' Soll nur innerhalb der Markierung getrimmt werden, wird das Ergebnis der
                        ' Ersetzungs-Funktion nur innherhalb der Positionen Ab und Bis ausgetauscht.
                        '
                        ' Am Ende wird die Zeile im Stringarray gesetzt.
                        '
                        
                        temp_string_3 = Replace(akt_zeile_mark, temp_string_1, temp_string_2)
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)

                    ElseIf (pFunktion = FKT_UCASE_LCASE) Then 'ISXF
                        '
                        ' Funktion Upper- Lower-Case
                        ' Die Funktionen "UCase" bzw. "LCase" werden auf die aktuelle Zeile oder
                        ' den Inhalt der Makierung ausgefuehrt.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                            
                            temp_string_3 = UCase(akt_zeile_mark)
                        
                        Else
                            
                            temp_string_3 = LCase(akt_zeile_mark)
                        
                        End If
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)

                    ElseIf (pFunktion = FKT_CAMEL_CASE) Then 'ISXF
                    
                        '
                        ' Funktion Upper-CamelCase
                        ' Der Inhalt der Variablen "akt_zeile_mark" wird in einen KlarText-String gewandelt.
                        ' Je nach dem Wert der Variable "m_toggle_mr_stringer_fkt" wird dabei als Trennzeichen
                        ' ein Leerstring (= keine Worttrennung) oder ein Unterstrich genutzt.
                        '
                        ' Wenn die Markierung benutzt werden soll, wird der gewandelte Text in die Originalzeile
                        ' eingebaut bzw. dort an den Positionen der Markierung ausgetauscht.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            temp_string_3 = getKlartext(akt_zeile_mark, LEER_STRING, ",.-()[]""")
                        
                        Else
                        
                            temp_string_3 = getKlartext(akt_zeile_mark, UNTER_STRICH, ",.-()[]""")
                        
                        End If
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_long_1 = Len(akt_zeile_mark) - Len(temp_string_3) ' Ergebnis kann kuerzer werden
                            
                            If (temp_long_1 > 0) Then
                            
                                temp_string_2 = String(temp_long_1, " ")
                                
                            End If
                           
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3 & temp_string_2)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)

                    ElseIf (pFunktion = FKT_ERSTELLE_NAMEN) Then 'ISXF
                        '
                        ' Funktion "Erstelle Namen"
                        '
                        ' Erstellt aus der aktuellen Zeile oder der Makierung Variablennamen.
                        '
                        ' Einmal mit keinem Trennzeichen (=Camelcase) und einmal mit einem
                        ' Unterstrich als Trennzeichen.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                            
                            temp_string_3 = getKlartext(akt_zeile_mark, LEER_STRING, LEER_ZEICHEN)
                        
                        Else
                            
                            temp_string_3 = LCase(getKlartext(akt_zeile_mark, UNTER_STRICH, LEER_ZEICHEN))
                        
                        End If
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)
                    
                    ElseIf ((pFunktion = FKT_SET_TRENNZEICHEN) Or (pFunktion = FKT_SET_TRENNZEICHEN_VOR) Or (pFunktion = FKT_SET_TRENNZEICHEN_ZURUECK)) Then
                        '
                        ' Funktion "Trennzeichen setzen"
                        ' Die Funktion "Trennzeichen setzen" hat 3 Unterfunktionen:
                        ' - Setzen an der aktuellen Position
                        ' - Am Wortanfang, ausgehend von der aktuellen Position
                        ' - Am Wortende, ausgehend von der aktuellen Position
                        '
                        ' Die Funktion wird nur dann ausgefuehrt, wenn die aktuelle Zeile keine Leerzeile ist.
                        '
                        If (aktuelle_zeile <> LEER_STRING) Then
                      
                            If (pFunktion = FKT_SET_TRENNZEICHEN_VOR) Then
                            
                                temp_long_2 = getPosWortende(aktuelle_zeile, ab_position)
                                
                                If (temp_long_2 > POSITION_0) Then
                                
                                    temp_long_2 = temp_long_2 + 1

                                End If
                            
                            ElseIf (pFunktion = FKT_SET_TRENNZEICHEN_ZURUECK) Then
                            
                                temp_long_2 = getPosWortAnfang(aktuelle_zeile, ab_position)
                                
                            Else

                                If (Len(aktuelle_zeile) >= ab_position) Then

                                    temp_long_2 = ab_position

                                Else

                                    temp_string_1 = aktuelle_zeile
                                    
                                    temp_string_2 = LEER_STRING
                                    
                                    temp_long_2 = 0
                                
                                End If

                            End If
                            
                            '
                            ' Pruefung: gibt es eine Trennstelle fuer den String?
                            ' Wenn dem so ist, wird in den Hilfsvariablen einmal der Text bis zur
                            ' Trennposition und einmal der Text nach der Trennposition gespeichert.
                            '
                            If (temp_long_2 > POSITION_0) Then
                            
                                temp_string_1 = Left(aktuelle_zeile, temp_long_2 - 1)
                                
                                temp_string_2 = Mid(aktuelle_zeile, temp_long_2, Len(aktuelle_zeile))
                            
                            End If
                                
                            Call cls_string_array.setString(zeilen_zaehler, temp_string_1 & temp_string_3 & temp_string_2)
                        
                        End If
                    
                    ElseIf (pFunktion = FKT_CLIP_ENTFERNE_TEXT) Then
                    
                        '
                        ' Funktion "Clip Entferne Text"
                        '
                        ' Entfernt den selektierten Bereich
                        '
                        Call cls_string_array.setString(zeilen_zaehler, getRemoveAbBis(aktuelle_zeile, ab_position, bis_position))
                    
                    ElseIf (pFunktion = FKT_CLIP_GET_TEXT) Then
                        
                        '
                        ' Funktion "Clip Get Text"
                        '
                        ' Gibt den Text in den Positionen Ab/Bis-Position zurueck.
                        '
                        Call cls_string_array.setString(zeilen_zaehler, getStringAbBis(aktuelle_zeile, ab_position, bis_position))
                    
                    ElseIf (pFunktion = FKT_KONSTANTEN_UEBER_SPLIT) Then
                    
                        '
                        ' Funktion "Erstelle Konstanten ueber Split"
                        '
                        
                        If (aktuelle_zeile <> LEER_STRING) Then
                        
                            If (pSelStart > POSITION_0) Then
    
                                temp_long_1 = InStr(aktuelle_zeile, inhalt_markierung)
                                
                            End If
    
                            If (temp_long_1 > POSITION_0) Then
                                
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    temp_string_2 = Left(aktuelle_zeile, temp_long_1 - 1)
                                    
                                    temp_string_3 = Mid(aktuelle_zeile, temp_long_1 + pSelLength, Len(aktuelle_zeile))
                                
                                Else
                                
                                    temp_string_3 = Left(aktuelle_zeile, temp_long_1 - 1)
                                    
                                    temp_string_2 = Mid(aktuelle_zeile, temp_long_1 + pSelLength, Len(aktuelle_zeile))
                                
                                End If
                                
                                '
                                ' Erstellung Konstanten-Name
                                ' Alle Leerzeichen werden mit einem Unterstrich vertauscht.
                                ' Alle Buchstaben als Grossbuchstaben.
                                '
                                temp_string_2 = UCase(getKlartext(replaceUmlaute(temp_string_2), UNTER_STRICH))
                
                                '
                                ' Erstellung Konstanten-Wert
                                ' Der Konstantenwert darf selber keine Anfuehrungszeichen enthalten.
                                ' Es werden alle Anfuehrungszeichen entfernt und das Ergebnis getrimmt.
                                '
                                temp_string_3 = Trim(Replace(temp_string_3, """", LEER_STRING))
                
                                Call cls_string_array.setString(zeilen_zaehler, MARKIER_STRING_INTERN_1 & temp_string_2 & MARKIER_STRING_INTERN_2 & AUSRICHT_STRING_TEMP_1 & MARKIER_STRING_INTERN_3 & temp_string_3 & MARKIER_STRING_INTERN_4)

                            End If

                        End If
                    
                    ElseIf (pFunktion = FKT_STRING_SPLIT) Then
                        '
                        ' Funktion "Split"
                        '
                        ' Zerteilt die Zeile anhand einer Position oder Markierung.
                        ' Eine Zeile kann nur dann gesplittet werden, wenn diese kein Leerstring ist.
                        ' Die Split-Position in der aktuellen Zeile wird durch den Wert in "temp_long_1" vorgegeben.
                        '
                        ' Ist eine Markierung vorgegeben, wird die Zeichenkette der
                        ' Markierung in der aktuellen Zeile gesucht. Wird die Markierung
                        ' nicht gefunden, wird der Wert in "temp_long_1" zu 0. Bei der
                        ' Verwendung eines Suchwortes, muss die Selektionslaenge groesser
                        ' 0 sein, d.h. es muss auch ein Wort markiert worden sein.
                        '
                        ' Durch eine Pruefung wird der Wert in "temp_long_1" auf groesser 0 geprueft.
                        ' Wenn dem so ist, wird die aktuelle Zeile gesplittet.
                        '
                        ' Soll die Zeile nicht gesplittet werden, wird hier nichts gemacht.
                        ' Die Zeile an der Position des Zeilenzaehlers wird nicht im Stringarray veraendert.
                        '
                        ' Soll nach einer Position gesplittet werden, ist der Wert in "temp_long_1" fest vorgegeben.
                        ' Soll nach einem String gesplittet werden, ist der Wert in "temp_long_1" variabel.
                        '
                        If (aktuelle_zeile <> LEER_STRING) Then
                        
                            If (pSelLength > 0) Then
    
                                temp_long_1 = InStr(aktuelle_zeile, inhalt_markierung)
                                
                            End If
    
                            If (temp_long_1 > POSITION_0) Then

                                If (m_toggle_mr_stringer_fkt) Then
                                    
                                    Call cls_string_array.setString(zeilen_zaehler, Left(aktuelle_zeile, temp_long_1 - 1))
                                
                                Else
                                    
                                    Call cls_string_array.setString(zeilen_zaehler, Mid(aktuelle_zeile, temp_long_1 + pSelLength, Len(aktuelle_zeile)))
                                
                                End If

                            End If

                        End If
                        
                    ElseIf (pFunktion = FKT_GENERATOR_VB_CHECK_LEER_STRING) Then

                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            aktuelle_zeile = Replace(Trim(akt_zeile_mark), """", LEER_STRING)
                            
                        End If
                        
                        '
                        ' Ist die aktuelle Zeile/Markierung kein Leerstring,
                        ' wird eine IF-Abfrage im Funktionsergebnis erstellt.
                        '
                        If (Trim(aktuelle_zeile) <> LEER_STRING) Then
        
                            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "if ( " & aktuelle_zeile & " = """" ) then "
                            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "    'call wl( ""Fehler: " & aktuelle_zeile & " nicht gesetzt"" )"
                            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "    '"
                            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "    'fkt_ergebnis = false"
                            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "End If"
                        
                        End If

                    ElseIf (pFunktion = FKT_GENERATOR_STRING_IT) Then
                        '
                        ' Funktion "String it"
                        ' Generator fuer die Stringerzeugung in Programmiersprachen. Setzt vor und hinter
                        ' dem zu bearbeitenden String die notwendigen Aktionen um einen String aufzubauen.
                        ' Hierzug werden in den beiden Variablen "temp_string_1" und "temp_string_2" die
                        ' Anweisungen der Zielsprache hinterlegt (weiter oben).
                        '
                        ' Beruecksichtigt die Markierung.
                        '
                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            aktuelle_zeile = Replace(Trim(akt_zeile_mark), """", temp_string_3)
                            
                        End If
                    
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_1 & aktuelle_zeile & temp_string_2)
                        
                    ElseIf (pFunktion = FKT_ZEILEN_ADD) Then
                
                        '
                        ' Funktion "Zeilen ADD"
                        '
                        ' Zusammenfassung von X-Zeilen zu einer.
                        '
                        ' Leerzeilen werden ueberlesen.
                        '
                        ' Die Anzahl der Zeilen, welche zusammengefasst werden, ist in
                        ' der Variablen "temp_long_1" gespeichert.
                        '
                        ' Die Anzahl der aktuell zusammengefassten Zeilen ist in der
                        ' Variablen "temp_long_2" gespeichert.
                        '
                        ' Ist "temp_long_2" gleich "temp_long_1" wird ein Zeilenumbruch
                        ' dem Ergebnis hinzugefuegt.
                        '
                
                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            str_fkt_ergebnis = str_fkt_ergebnis & aktuelle_zeile & temp_string_1
                            
                            If (temp_long_2 = temp_long_1) Then
                            
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                
                                temp_long_2 = 1
                                
                            Else
                            
                                temp_long_2 = temp_long_2 + 1
                                
                            End If
                            
                        End If
                    
                    ElseIf (pFunktion = FKT_ERSTELLE_BLOCK) Then
                    
                        '
                        ' Funktion "Erstelle Block"
                        '
                        ' In der Variablen "temp_long_1" ist die Breite des zu erstellen
                        ' Blockes festgelegt.
                        '
                        ' Die aktuelle Leseposition fuer die aktuelle Zeile ist in der
                        ' Variablen "temp_long_2" gespeichert.
                        '
                        temp_long_2 = 1
                        
                        '
                        ' In der Variablen "temp_string_1" wird immer ein Teilstring in
                        ' der Laenge der Breite aus "temp_long_1" gespeichert.
                        '
                        ' Die Leseposition wird immer um die Breite weitergestellt.
                        '
                        ' Ist der herausgeschnittene Teilstring ein Leerstring, ist die
                        ' While-Schleife beendet.
                        '
                        temp_string_1 = Mid(aktuelle_zeile, temp_long_2, temp_long_1)
                        
                        While (temp_string_1 <> LEER_STRING)
                        
                            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & temp_string_1
                             
                            temp_long_2 = temp_long_2 + temp_long_1
                        
                            temp_string_1 = Mid(aktuelle_zeile, temp_long_2, temp_long_1)
                            
                        Wend

                    ElseIf (pFunktion = FKT_GENERATOR_VARIABLEN_DEKLARATION) Then
                        '
                        ' Funktion "Variablen Deklaration"
                        ' Aus der aktuellen Zeile oder der Markierung wird abwechselnd fuer
                        ' Java und Visual-Basic eine Variabelen-Deklaration erstellt.
                        '
                        ' Java =  String xxx = null;
                        ' VB   =  Dim xxx As String
                        '
                        If (Trim(akt_zeile_mark) <> LEER_STRING) Then
                        
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                str_fkt_ergebnis = str_fkt_ergebnis & "String " & akt_zeile_mark & " = null;" & zeichen_zeilenumbruch
                                
                            Else
                            
                                str_fkt_ergebnis = str_fkt_ergebnis & "Dim " & akt_zeile_mark & " As String" & zeichen_zeilenumbruch
                                
                            End If
                            
                        End If
    
                    ElseIf (pFunktion = FKT_GENERATOR_SET_NULL) Then
                        '
                        ' Funktion "Variablen auf null stellen"
                        '
                        ' Aus der aktuellen Zeile oder der Markierung wird abwechselnd fuer
                        ' Java und Visual-Basic eine Anweisung fuer die Null-Setzung erstellt.
                        ' Leerzeilen werden ueberlesen.
                        '
                        ' Java =  xxx = null;
                        ' VB   =  set xxx = nothing
                        '
                        If (Trim(akt_zeile_mark) <> LEER_STRING) Then
                        
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                str_fkt_ergebnis = str_fkt_ergebnis & akt_zeile_mark & " = null;" & zeichen_zeilenumbruch
                                
                            Else
                            
                                str_fkt_ergebnis = str_fkt_ergebnis & "set " & akt_zeile_mark & " = nothing" & zeichen_zeilenumbruch
                                
                            End If
                            
                        End If
                        
                    ElseIf (pFunktion = FKT_GROUP_NACH_STRING) Then
                        '
                        ' Funktion "Group nach String"
                        '
                        ' Diese Funktion fuegt eine Leerstelle in ein, wenn sich
                        ' der Text der Markierung gegenueber den vorhergehenden
                        ' Wert aendert. In der Variablen "temp_string_1" wird der
                        ' letzte Text der Markierung gespeichert.
                        '
                        ' Leerzeilen werden nicht Bestandteil des Ergebnisses.
                        '
                        If (Trim(akt_zeile_mark) <> LEER_STRING) Then
                        
                            If (akt_zeile_mark <> temp_string_1) Then
                            
                                '
                                ' Aendert sich der Text, wird eine Leerzeile dem Ergebnis hinzugefuegt.
                                '
                                str_fkt_ergebnis = str_fkt_ergebnis & vbCrLf
                                
                                '
                                ' Der neue Gruppierungsstring wird in der Variable "temp_string_1" vermerkt.
                                '
                                temp_string_1 = akt_zeile_mark
                            
                            End If
                            
                            str_fkt_ergebnis = str_fkt_ergebnis & vbCrLf & aktuelle_zeile
                            
                        End If
    
                    ElseIf (pFunktion = FKT_CMD_RENAME) Then
                        '
                        ' Funktion "Rename"
                        '
                        ' Erstellung eines neuen Dateinamens ohne Leerzeichen oder Sonderzeichen.
                        '
                        ' Abwechselnd mit dem "Rename"-Befehl fuer BAT-Dateien.
                        '
                        ' Leerzeilen werden ueberlesen.
                        '
                        ' Soll eine Markierung benutzt werden, wird der Dateiname aus der
                        ' Markierung genommen, ansonsten wird die aktuelle Zeile genommen.
                        '
                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            temp_string_3 = Trim(akt_zeile_mark)
                            
                            temp_string_3 = renameYoutube(temp_string_3, True)
                            
                            temp_string_3 = renameDatei(temp_string_3)
                            
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                 temp_string_3 = "rename """ & akt_zeile_mark & """" & TRENN_STRING_6 & " " & TRENN_STRING_7 & """" & temp_string_3 & """" & TRENN_STRING_8
                                
                            End If
    
                            Call cls_string_array.setString(zeilen_zaehler, temp_string_3)
                            
                        End If
                    
                    ElseIf (pFunktion = FKT_GENERATOR_DEBUG_AUSGABE) Then
                        '
                        ' Funktion "Debug-Ausgabe"
                        ' Erstellt fuer die aktuelle Zeile oder die Markierung eine Debug-Ausgabe fuer
                        ' VB, PHP und Java. Leerzeilen werden nicht beruecksichtigt.
                        '
                        akt_zeile_mark = Trim(akt_zeile_mark)
                        
                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            If (m_zaehler_debug_print = 2) Then
                            
                                Call cls_string_array.setString(zeilen_zaehler, "'##sss( """ & Replace(akt_zeile_mark, """", """""") & " =>"" & " & Replace(akt_zeile_mark, """", "") & " & ""<"" )")
                                
                            ElseIf (m_zaehler_debug_print = 3) Then
                            
                                Call cls_string_array.setString(zeilen_zaehler, "temp_str = temp_str & chr(13) & """ & Replace(akt_zeile_mark, """", """""") & " =>"" & " & Replace(akt_zeile_mark, """", "") & " & ""<"" ")
                            
                            ElseIf (m_zaehler_debug_print = 4) Then
                            
                                Call cls_string_array.setString(zeilen_zaehler, "temp_str += ""\n" & Replace(akt_zeile_mark, """", "\""") & " =>"" + " & Replace(akt_zeile_mark, """", "") & " + ""<"";")
                                
                            ElseIf (m_zaehler_debug_print = 5) Then
                            
                                Call cls_string_array.setString(zeilen_zaehler, " =>"" & " & Replace(akt_zeile_mark, """", "") & " & ""< ")

                            ElseIf (m_zaehler_debug_print = 6) Then

                                Call cls_string_array.setString(zeilen_zaehler, " =>"" + " & Replace(akt_zeile_mark, """", LEER_STRING) & " + ""< ")

                            Else

                                Call cls_string_array.setString(zeilen_zaehler, "FkLogger.wl( """ & Replace(akt_zeile_mark, """", "\""") & " =>"" + " & Replace(akt_zeile_mark, """", LEER_STRING) & " + ""<"" );")

                            End If
                            
                        End If
                        
                    ElseIf ((pFunktion = FKT_ERSTELLE_XML) Or (pFunktion = FKT_ERSTELLE_XML_2)) Then
                        '
                        ' Funktion "XML-Erstellung"
                        ' Die aktuelle Zeile oder Markierung wird als Tag-Namen betrachtet.
                        ' Dabei wird der TAG-Name in Grossbuchstaben gewandelt und ein XML-TAG
                        ' erstellt. Dieses einmal in einer Klammer oder mit Start- und End-Tag.
                        '
                        akt_zeile_mark = Trim(akt_zeile_mark)
                        
                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            If (pFunktion = FKT_ERSTELLE_XML_2) Then

                                str_fkt_ergebnis = str_fkt_ergebnis & "<" & UCase(getKlartext(akt_zeile_mark, UNTER_STRICH)) & " x_attribut=""" & akt_zeile_mark & """ /> " & zeichen_zeilenumbruch
                            
                            Else
                            
                                akt_zeile_mark = UCase(getKlartext(akt_zeile_mark, UNTER_STRICH))
                                
                                str_fkt_ergebnis = str_fkt_ergebnis & temp_string_1 & akt_zeile_mark & temp_string_2
                                
                                If (temp_string_3 <> LEER_STRING) Then
                                
                                    str_fkt_ergebnis = str_fkt_ergebnis & akt_zeile_mark & temp_string_3
                                
                                End If
                                
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                
                            End If
                        
                        End If
                        
                    ElseIf ((pFunktion = FKT_SINGLETON_JAVA) Or (pFunktion = FKT_GETTER_SETTER_JAVA) Or (pFunktion = FKT_GETTER_SETTER_VB) Or (pFunktion = FKT_GETTER_SETTER_JAVA_SCRIPT)) Then
                        '
                        ' Funktion "Getter Setter Java" oder "Getter Setter VB"
                        '
                        ' Es wird ein Trim auf die aktuelle Zeile/Markierung gemacht.
                        '
                        akt_zeile_mark = Trim(akt_zeile_mark)
                        
                        '
                        ' Pruefung: aktuelle Zeile ungleich Leerstring ?
                        '
                        ' Leerstrings werden ueberlesen
                        '
                        If (akt_zeile_mark <> LEER_STRING) Then
                        
                            '
                            ' Variablentyp-Trennposition bestimmen
                            '
                            ' Es wird ein Trennstring gesucht, welcher den Namen vom Typ trennt.
                            ' Die Position des Trennstrings wird in "temp_long_1" gespeichert.
                            ' In der Variablen "temp_long_3" wird die Laenge des Trennstrings gespeichert.
                            '
                            ' Es wird zuerst ein Gleichheitszeichen gesucht. Wurde kein Gleichheitszeichen
                            ' gefunden, wird nach dem String " AS " gesucht.
                            '
                            ' OPTIMIERUNGSMOEGLICHKEIT:
                            ' Es koennte auch ein Trennstring von aussen im Parameter "pOptString1" uebergeben werden.
                            '
                            Dim str_var_typ As String
                            
                            temp_long_1 = InStr(akt_zeile_mark, "=")
                            
                            If (temp_long_1 <= 0) Then
                            
                                temp_long_1 = InStr(akt_zeile_mark, " As ")
                                
                                temp_long_3 = 4
                                
                            Else
                            
                                temp_long_3 = 1
                            
                            End If
                            
                            '
                            ' Pruefung: Trennstring gefunden ?
                            '
                            ' Ist eine Position fuer ein Typ-Trennzeichen gefunden worden, wird der
                            ' Variablen-Name aus dem ersten Teilstring, der Typ aus dem Zweiten
                            ' Teilstring gelesen.
                            '
                            ' Der Variablenname wird in der Variablen "temp_string_1" gespeichert.
                            ' Der Variablentyp wird in der Variablen "temp_string_2" gespeichert.
                            '
                            ' Der Variablenname bekommt den Praefix "m_" vorangestellt, wird in
                            ' lower-Case gewandelt und mit Unterstrichen versehen.
                            '
                            ' Der Variablentyp wird so uebernommen, wie dieser gelesen worden ist.
                            ' Ist keine Typinformation in der aktuellen Zeile vorhanden, wird als
                            ' Typ "String" genommen.
                            '
                            If (temp_long_1 > POSITION_0) Then
                            
                                temp_long_2 = temp_long_1 + temp_long_3
                                
                                temp_string_1 = "m_" & LCase(getKlartext(Trim(Left(akt_zeile_mark, temp_long_1 - 1)), UNTER_STRICH))
                                
                                temp_string_2 = getKlartext(Trim(Left(akt_zeile_mark, temp_long_1 - 1)), LEER_STRING)
                                
                                str_var_typ = Trim(Mid(akt_zeile_mark, temp_long_2, Len(akt_zeile_mark)))
                            
                            Else
                            
                                str_var_typ = "String" ' " & str_var_typ  & "
                            
                                temp_string_1 = "m_" & LCase(getKlartext(akt_zeile_mark, UNTER_STRICH))  ' member-Variable
                                
                                temp_string_2 = getKlartext(akt_zeile_mark, LEER_STRING) ' CamelCase-Grundname

                            End If
                            
                            If (pFunktion = FKT_SINGLETON_JAVA) Then

                                temp_string_3 = temp_string_3 & zeichen_zeilenumbruch & IIf(m_toggle_mr_stringer_fkt, "private ", LEER_STRING) & str_var_typ & " " & temp_string_1 & " = "
                                
                                temp_long_1 = 0
                                
                                If ((LCase(str_var_typ) = "boolean") Or (LCase(str_var_typ) = "bool")) Then
                                
                                    temp_string_3 = temp_string_3 & "false; // true;"
                                    
                                    str_var_typ = "boolean"
                                    
                                ElseIf (LCase(str_var_typ) = "long") Then
                                
                                    temp_string_3 = temp_string_3 & "0;"
                                    
                                ElseIf (LCase(str_var_typ) = "double") Then
                                
                                    temp_string_3 = temp_string_3 & "0.0d;"
                                    
                                Else
                                
                                    temp_string_3 = temp_string_3 & "null;"
                                    
                                    temp_long_1 = 1
                                
                                End If
                                
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                
                                If (temp_long_1 = 0) Then
                                    
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "public " & str_var_typ & " get" & temp_string_2 & "() { return " & temp_string_1 & "; }"
                                
                                Else
                                    
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "public " & str_var_typ & " get" & temp_string_2 & "() { if ( " & temp_string_1 & " == null ) { " & temp_string_1 & " = new " & str_var_typ & "(); } return " & temp_string_1 & "; }"
                                
                                End If
                                
                            ElseIf (pFunktion = FKT_GETTER_SETTER_VB) Then
                            
                                temp_string_3 = temp_string_3 & zeichen_zeilenumbruch & IIf(m_toggle_mr_stringer_fkt, "Private ", "Dim ") & temp_string_1 & " As " & str_var_typ & " ' = "
                                
                                If ((LCase(str_var_typ) = "boolean") Or (LCase(str_var_typ) = "bool")) Then
                                
                                    temp_string_3 = temp_string_3 & "false ' true"
                                    
                                    str_var_typ = "boolean"
                                    
                                ElseIf (LCase(str_var_typ) = "bigdecimal") Then
                                
                                    temp_string_3 = temp_string_3 & " 0.0"
                                
                                ElseIf (LCase(str_var_typ) = "long") Then
                                
                                    temp_string_3 = temp_string_3 & "0"
                                
                                ElseIf (LCase(str_var_typ) = "double") Then
                                
                                    temp_string_3 = temp_string_3 & "0.0"
                                    
                                Else
                                
                                    temp_string_3 = temp_string_3 & """"""
                            
                                End If
                                
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "'######## VB #############"
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "'"
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "Public Function get" & temp_string_2 & "() As " & str_var_typ
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & zeichen_zeilenumbruch & "    get" & temp_string_2 & " = " & temp_string_1
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & zeichen_zeilenumbruch & "End Function"
                                
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "'######## VB #############"
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "'"
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "Public Sub set" & temp_string_2 & "( p" & temp_string_2 & " As " & str_var_typ & " )"
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & zeichen_zeilenumbruch & "    " & temp_string_1 & " = p" & temp_string_2
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & zeichen_zeilenumbruch & "End Sub"
                                
                            ElseIf (pFunktion = FKT_GETTER_SETTER_JAVA_SCRIPT) Then
                                
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                    
                                    If (temp_long_1 > 0) Then ' explizite Typangabe --> dann Initialisierung mit undefined
                                    
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  " & temp_string_1 & " : undefined,"
                                        
                                    Else
                                    
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  " & temp_string_1 & " : """","
                                        
                                    End If

                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "get" & temp_string_2 & " : function()"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "{"
                                    
                                    If (temp_long_1 > 0) Then ' explizite Typangabe --> dann Singletonpattern hinzufuegen
                                    
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  if ( this." & temp_string_1 & " == undefined )"
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  {"
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "    this." & temp_string_1 & " = new " & str_var_typ & "();"
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  }"
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                        
                                    End If
                                    
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  return this." & temp_string_1 & ";"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "},"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "set" & temp_string_2 & " : function( p" & temp_string_2 & " )"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "{"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  this." & temp_string_1 & " = p" & temp_string_2 & ";"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "},"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                
                                Else
                                
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                    
                                    If (temp_long_1 > 0) Then ' explizite Typangabe --> dann Initialisierung mit undefined
                                    
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  this." & temp_string_1 & " = undefined;"
                                        
                                    Else
                                    
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  this." & temp_string_1 & " = """";"
                                        
                                    End If
                                    
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "Bean.prototype.get" & temp_string_2 & " = function()"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "{"
                                    
                                    If (temp_long_1 > 0) Then ' explizite Typangabe --> dann Singletonpattern hinzufuegen
                                    
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  if ( this." & temp_string_1 & " == undefined )"
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  {"
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "    this." & temp_string_1 & " = new " & str_var_typ & "();"
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  }"
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                        
                                    End If
                                    
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  return this." & temp_string_1 & ";"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "}"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "Bean.prototype.set" & temp_string_2 & " = function( p" & temp_string_2 & " )"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "{"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "  this." & temp_string_1 & " = p" & temp_string_2 & ";"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "}"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                    
                                End If
                            
                            Else ' = FKT_GETTER_SETTER_JAVA

                                temp_string_3 = temp_string_3 & zeichen_zeilenumbruch & IIf(m_toggle_mr_stringer_fkt, "private ", "") & str_var_typ & " " & temp_string_1 & " = "
                                
                                If (LCase(str_var_typ) = "boolean") Then
                                
                                    temp_string_3 = temp_string_3 & "false; // true;"
                                    
                                ElseIf (LCase(str_var_typ) = "bigdecimal") Then
                                
                                    temp_string_3 = temp_string_3 & "null; // new BigDecimal( ""0.00"" );"
                                    
                                    temp_string_3 = temp_string_3 & " 0.0"
                                
                                ElseIf ((LCase(str_var_typ) = "long") Or (LCase(str_var_typ) = "integer")) Then
                                
                                    temp_string_3 = temp_string_3 & "0;"
                                
                                ElseIf (LCase(str_var_typ) = "double") Then
                                
                                    temp_string_3 = temp_string_3 & "0.0;"
                                Else
                                
                                    temp_string_3 = temp_string_3 & """"";"
                            
                                End If

                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch

                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "public " & str_var_typ & " get" & temp_string_2 & "() { return " & temp_string_1 & "; }"
                                
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "public void set" & temp_string_2 & "( " & str_var_typ & " p" & temp_string_2 & " ) { " & temp_string_1 & " = p" & temp_string_2 & "; }"
                                
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                                
                            End If

                        End If
                    
                    ElseIf (pFunktion = FKT_CALC_SUMME) Then
                        '
                        ' Funktion "Summe"
                        '
                        ' Aus der aktuellen Zeile wird eine Zahl gebildet und diese dann
                        ' auf die Summe in "temp_double_1" hinzugezaehlt. Kann aus der
                        ' aktuellen Zeile keine Zahl erstellt werden, ist die Zeile 0.
                        '
                        akt_zeile_mark = Trim(akt_zeile_mark)
                        
                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            temp_long_1 = temp_long_1 + 1
                            
                            temp_double_1 = Val(getzahl(akt_zeile_mark, 4, False))
                            
                            temp_double_2 = temp_double_2 + temp_double_1
                        
                            str_fkt_ergebnis = str_fkt_ergebnis & "OK     >" & temp_long_1 & "< >" & akt_zeile_mark & "< >" & temp_double_1 & "< >" & temp_double_2 & "<" & zeichen_zeilenumbruch
                        
                        Else
                        
                            temp_double_1 = 0
                            
                            str_fkt_ergebnis = str_fkt_ergebnis & "FEHLER >" & temp_long_1 & "< >" & akt_zeile_mark & "< >" & temp_double_1 & "< >" & temp_double_2 & "<" & zeichen_zeilenumbruch

                        End If
                                                   
                    ElseIf (pFunktion = FKT_MARKIERE_DOPPELT_PLUS_1_ZEILE) Then
                        
                        '
                        ' Funktion "FKT_MARKIERE_DOPPELT_PLUS_1_ZEILE"
                        '
                        ' Ist der String der letzten Zeile gleich dem String der aktuellen Zeile,
                        ' Wird die aktuelle Zeile markiert.
                        '
                        ' Die aktuelle Zeile/Markierung wird in der Variablen "temp_string_1" gespeichert.
                        '
                        ' Die erste Zeile muss nicht gemerkt werden, da es keine Vorgaengerzeile gibt.
                        '
                        If (akt_zeile_mark = temp_string_1) Then
                        
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                Call cls_string_array.setString(zeilen_zaehler, aktuelle_zeile & MARKIERUNG_DOPPELTE_VORKOMMEN)
                                
                            Else
                            
                                Call cls_string_array.setString(zeilen_zaehler, MARKIERUNG_DOPPELTE_VORKOMMEN & aktuelle_zeile)
                                
                            End If
                        
                        End If
                        
                        temp_string_1 = akt_zeile_mark
                    
                    ElseIf (pFunktion = FKT_MARKIERE_DOPPELT_PLUS_1_ZEILE_MINUS) Then
                        '
                        ' Stimmt die aktuelle Zeile mit der vorhergehenden ueberein ?
                        '
                        ' Wenn ja, markiere die Zeile mit dem String MARKIERUNG_DOPPELTE_VORKOMMEN
                        '
                        '
                        If (akt_zeile_mark = temp_string_1) Then
                        
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                Call cls_string_array.setString(zeilen_zaehler, aktuelle_zeile & MARKIERUNG_DOPPELTE_VORKOMMEN)
                                
                            Else
                            
                                Call cls_string_array.setString(zeilen_zaehler, MARKIERUNG_DOPPELTE_VORKOMMEN & aktuelle_zeile)
                            
                            End If
                            
                            '
                            ' Es wird die vorhergehende Zeile aus dem Stringarray geholt.
                            '
                            temp_string_1 = cls_string_array.getString(zeilen_zaehler - 1)
                            
                            '
                            ' Es wird dort auf das Vorhandensein des String MARKIERUNG_DOPPELTE_VORKOMMEN geprueft.
                            ' Ist die Markierung nicht vorhanden, wird die Markierung in der vorhergehenden Zeile
                            ' eingebaut.
                            '
                            If (InStr(temp_string_1, MARKIERUNG_DOPPELTE_VORKOMMEN) = 0) Then
                           
                                If (m_toggle_mr_stringer_fkt) Then
                                    
                                    Call cls_string_array.setString(zeilen_zaehler - 1, temp_string_1 & MARKIERUNG_DOPPELTE_VORKOMMEN)
                                    
                                Else

                                    Call cls_string_array.setString(zeilen_zaehler - 1, MARKIERUNG_DOPPELTE_VORKOMMEN & temp_string_1)
                                    
                                End If
                           
                           End If
                        
                        End If
                        
                        temp_string_1 = akt_zeile_mark
                    
                    ElseIf (pFunktion = FKT_ZEILEN_BOOLEAN) Then
                    
                        '
                        ' Funktion "Zeilen Boolean"
                        ' Es wird der optionale Parameter "pKennzeichen1" negiert.
                        ' Dessen Wert bestimmt den Wert in der Ausgabe.
                        ' Der Wert ist einmal 0 und 1, und einmal True und False.
                        '
                        pKennzeichen1 = Not pKennzeichen1
                        
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            Call cls_string_array.setString(zeilen_zaehler, IIf(pKennzeichen1, "1", "0"))
                            
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, LEER_STRING & pKennzeichen1)
                            
                        End If
                        
                    ElseIf (pFunktion = FKT_ZEILEN_ZAEHLER) Then
                        '
                        ' Funktion "Zeilen Zaehler"
                        ' Zaehlt die Zeilen der Eingabe, bzw. jede Zeile bekommt eine
                        ' Zeilennummer im Ergebnis.
                        '
                        If (temp_long_1 = 0) Then
                        
                            Call cls_string_array.setString(zeilen_zaehler, LEER_STRING & zeilen_zaehler)
                            
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, Right(NULL_ZIFFERN_100 & zeilen_zaehler, temp_long_1))
                            
                        End If
                        
                    ElseIf (pFunktion = FKT_MAKE_LONG_DATUM) Then
                    
                        '
                        ' Funktion "Make Long-Datum"
                        '
                        ' Die Zeichen ab der "Ab-Position" werden so umgestellt, dass ein Long-Datum erstehen kann.
                        ' Es wird angenommen, dass sich an der Stelle "Ab-Position" ein Datum im Format "TT.MM.JJJJ" steht.
                        ' Diese Zeichenfolge wircht nach "JJJJMMTT" umgewandelt und ersetzt den Ursprungsstring
                        '
                        temp_string_1 = Mid(aktuelle_zeile, ab_position + 6, 4) & Mid(aktuelle_zeile, ab_position + 3, 2) & Mid(aktuelle_zeile, ab_position, 2)
                        
                        Call cls_string_array.setString(zeilen_zaehler, replaceSubstringAbBis(aktuelle_zeile, ab_position, ab_position + 9, temp_string_1))
                    
                    ElseIf (pFunktion = FKT_LEERZEILEN_LOESCHEN) Then
                    
                        '
                        ' Funktion "Leerzeilen loeschen"
                        '
                        ' Ist die Laenge der aktuellen Zeile getrimmt groesser als 0,
                        ' wird die aktuelle Zeile fuer das Ergebnis uebernommen.
                        ' (Ist keine Leerzeile)
                        '
                        ' Zeilen mit einer Laenge von 0 Zeichen, werden ueberlesen.
                        '
                    
                        If (Len(Trim(aktuelle_zeile)) > 0) Then
                        
                            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & aktuelle_zeile

                        End If
    
                    ElseIf (pFunktion = FKT_LEERZEILEN_EINFUEGEN) Then
                    
                        '
                        ' Funktion "Leerzeilen einfuegen"
                        '
                        ' In der Variablen "temp_long_2" wird die aktuelle Zeilenlaenge gespeichert.
                        '
                    
                        temp_long_2 = Len(Trim(aktuelle_zeile))
                        '
                        ' Es werden nur gesetzte Zeilen uebernommen, d.h. der Wert in
                        ' der Variablen "temp_long_2" muss groesser 0 sein.
                        '
                        If (temp_long_2 > 0) Then
                        
                            '
                            ' Bei der ersten Zeile wird kein Zeilenumbruchszeichen vorangestellt.
                            ' Ist die Variable "temp_long_1" gleich 1, wird das Zeilenumbruchszeichen
                            ' vorangestellt. Die Variable "temp_long_1" wurde bei der Initialisierung
                            ' dieser Funktion auf den Wert 0 gestellt.
                            '
                            If (temp_long_1 = 1) Then
                            
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch
                            
                            End If
                            
                            '
                            ' Die aktuelle Zeile wird dem Ergebnis hinzugefuegt und bekommt
                            ' selber wieder ein Zeilenumbruchszeichen angehaengt.
                            '
                            str_fkt_ergebnis = str_fkt_ergebnis & aktuelle_zeile & zeichen_zeilenumbruch
                            
                            '
                            ' Nach dem ersten Durchlauf durch diesen Code, wird die Variable "temp_long_1"
                            ' auf den Wert 1 gestellt.
                            '
                            temp_long_1 = 1
                            
                        End If
                        
                    ElseIf ((pFunktion = FKT_GET_UNIQUE) Or (pFunktion = FKT_GET_DOPPELTE_VORKOMMEN) Or (pFunktion = FKT_GET_EINMALIGE_VORKOMMEN)) Then
                        '
                        ' Die aktuelle Zeile/Markierung wird getrimmt in der Variablen "temp_string_1" gespeichert.
                        '
                        temp_string_1 = Trim(akt_zeile_mark)
                        
                        '
                        ' Pruefung: "temp_string_1" ungleich Leerstring ?
                        '
                        ' Die Funktion wird nur auf gesetzte String ausgefuehrt.
                        ' Leerzeilen werden ueberlesen, da eine Verarbeitung keinen Sinn macht.
                        '
                        If (temp_string_1 <> LEER_STRING) Then
                            '
                            ' Vermeidung von zufaelligen Treffern
                            ' Damit durch die aneinanderreihung in "temp_string_2" keine zufaelligen Treffer
                            ' erzeugt werden koennen, wird der zu suchende String um "#1#" am Start und
                            ' "#2#" am Ende erweitert.
                            '
                            ' Da innerhalt dieser 3 Funktionen es nur um die Ausortierung von ganzen Zeilen
                            ' geht, muss der aktuelle Wert aus "temp_string_1" nicht nochmal in einer anderen
                            ' Variable weggespeichert werden. Es geht nur darum, ob die aktuelle Zeile zum
                            ' Ergebnis hinzugefuegt werden kann oder nicht.
                            '
                            temp_string_1 = "#1#" & temp_string_1 & "#2#"
                            
                            If (pFunktion = FKT_GET_UNIQUE) Then
                                '
                                ' Funktion "get Unique"
                                ' Ist der Suchstring aus "temp_string_1" noch nicht in "temp_string_2" enthalten,
                                ' wird die aktuelle Zeile in das Funktionsergebnis aufgenommen.
                                '
                                If (InStr(temp_string_2, temp_string_1) <= POSITION_0) Then
                                
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & aktuelle_zeile
                                    
                                    temp_string_2 = temp_string_2 & temp_string_1
                                
                                End If
                        
                            ElseIf (pFunktion = FKT_GET_DOPPELTE_VORKOMMEN) Then
                                '
                                ' Funktion "get doppelte Vorkommen"
                                '
                                ' In der Variablen "temp_string_2" wird eine Liste der schon einmal
                                ' vorgekommenen Strings erzeugt.
                                '
                                ' In "temp_string_3" wird eine Liste der bereits zweimal vorgekommenen
                                ' Strings erzeugt.
                                '
                                ' Pruefung: Erstes vorkommen des Strings?
                                ' Es wird geprueft, ob der aktuelle String in "temp_string_1" schon in
                                ' "temp_string_2" vorhanden ist. Ist das nicht der Fall, wird
                                ' "temp_string_1" in "temp_string_2" gespeichert.
                                '
                                If (InStr(temp_string_2, temp_string_1) <= 0) Then
                                    
                                    temp_string_2 = temp_string_2 & temp_string_1
                                
                                Else
                                    '
                                    ' weiteres Auftreten des Strings
                                    ' Ist der aktuelle String aus "temp_string_1" nicht in "temp_string_3" vorhanden,
                                    ' wird "temp_string_1" in "temp_string_3" aufgenommen.
                                    '
                                    ' Gleichzeitig wird die aktuelle Zeile in das Funktionsergebnis aufgenommen.
                                    '
                                    If (InStr(temp_string_3, temp_string_1) <= 0) Then
                                                           
                                       temp_string_3 = temp_string_3 & temp_string_1
                                       
                                       str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & aktuelle_zeile
            
                                    End If
                                
                                End If
                            
                            ElseIf (pFunktion = FKT_GET_EINMALIGE_VORKOMMEN) Then
                                
                                '
                                ' Funktion "get einmalige Vorkommen"
                                '
                                ' Pruefung: Ist der String schon mal vorgekommen
                                '
                                If (InStr(temp_string_2, temp_string_1) <= POSITION_0) Then
                                
                                    '
                                    ' erstes auftreten des Strings
                                    '
                                    ' Pruefung: ob String nicht im Speicher fuer doppelte ist
                                    '
                                    If (InStr(temp_string_3, temp_string_1) <= 0) Then
                                    
                                        temp_string_2 = temp_string_2 & temp_string_1
                                        
                                    End If
                                
                                Else
                                    '
                                    ' weiteres Auftreten des Strings
                                    ' ... String im Remove-String vermerken
                                    ' ... String aus dem TempString2 rauswerfen
                                    '
                                    temp_string_3 = temp_string_3 & temp_string_1
                                    
                                    temp_string_2 = Replace(temp_string_2, temp_string_1, "")
                                    
                                End If
                            
                            End If

                        End If
    
                    ElseIf ((pFunktion = FKT_CSV_SWAP) Or (pFunktion = FKT_CSV_JAVA_CASE) Or (pFunktion = FKT_CSV_VB_KONVERTER) Or (pFunktion = FKT_CSV_JAVA_PROP)) Then

                        '
                        ' Es wird die Position des CSV-Strings innerhalb der aktuellen Zeile gesucht.
                        ' Die Position wird in "temp_long_1" gespeichert.
                        '
                        temp_long_1 = InStr(aktuelle_zeile, pOptString1)

                        '
                        ' Pruefung: CSV-String in der aktuellen Zeile gefunden ?
                        '
                        If (temp_long_1 > POSITION_0) Then

                            '
                            ' Bestimmung bis zu welcher Position das CSV-Zeichen geht.
                            '
                            ' OPTIMIERUNG: Die Laenge des OptStrings1 koennte auch in der
                            '              Variablen "temp_long_3" gespeichert werden.
                            '              Das waere dann nur einmal die Laengenermittlung
                            '              von "pOptString1".
                            '
                            temp_long_2 = temp_long_1 + Len(pOptString1)
                            
                            '
                            ' Spaltenswitch
                            ' Bei jedem zweiten Durchgang werden die Positionen der Spalten
                            ' getauscht.
                            '
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                temp_string_1 = Left(aktuelle_zeile, temp_long_1 - 1)
                                
                                temp_string_2 = Mid(aktuelle_zeile, temp_long_2, Len(aktuelle_zeile))
                                
                            Else
                            
                                temp_string_2 = Left(aktuelle_zeile, temp_long_1 - 1)
                                
                                temp_string_1 = Mid(aktuelle_zeile, temp_long_2, Len(aktuelle_zeile))
                                
                            End If
                      
                            If (pFunktion = FKT_CSV_JAVA_CASE) Then
                            
                                Call cls_string_array.setString(zeilen_zaehler, "case " & temp_string_1 & " : { " & temp_string_2 & " break; }")
                      
                            ElseIf (pFunktion = FKT_CSV_JAVA_PROP) Then
                            
                                temp_string_1 = Trim(temp_string_1)
                                temp_string_2 = Trim(temp_string_2)
                                
                                Call cls_string_array.setString(zeilen_zaehler, LEER_STRING & STR_VAR_NAME_PROPERTIES_LOKAL & ".setProperty( " & temp_string_1 & ", " & AUSRICHT_STRING_TEMP_1 & temp_string_2 & AUSRICHT_STRING_TEMP_2 & " );")
                                
                            ElseIf (pFunktion = FKT_CSV_VB_KONVERTER) Then
                            
                                temp_string_1 = Trim(temp_string_1)
                                temp_string_2 = Trim(temp_string_2)
                            
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "    " & temp_string_3 & " ( " & pOptString2 & " = """ & temp_string_1 & """ ) Then "
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & ""
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "        " & pOptString3 & " = """ & temp_string_2 & """ "
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & ""
                                    
                                temp_string_3 = "ElseIf"
                                
                            Else
                            
                                Call cls_string_array.setString(zeilen_zaehler, temp_string_2 & pOptString1 & temp_string_1)
                                
                            End If
                        
                        End If
    
                    ElseIf (pFunktion = FKT_CSV_ERSTELLE_CSV) Then
                        '
                        ' Funktion "Erstelle CSV Liste"
                        ' Erstellt eine Zeile, in welcher die Elemente (aktuelle Zeile/Makrierung)
                        ' durch das Trennzeichen getrennt sind.
                        '
                        If (zeilen_zaehler > 1) Then
                
                            str_fkt_ergebnis = str_fkt_ergebnis & pOptString1
                    
                        End If
                        
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            str_fkt_ergebnis = str_fkt_ergebnis & """" & akt_zeile_mark & """"
                            
                        Else
                        
                            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeile_mark
                        
                        End If
                     
                    ElseIf (pFunktion = FKT_TRIM_AUFEINANDERFOLGENDE_LEERZEICHEN) Then
                        '
                        ' Funktion "Trim X"
                        ' TrimX eliminiert doppelte Leerzeichen durch den gesamten String hindurch.
                        '
                        If (knz_benutze_markierung) Then
                            
                            Call cls_string_array.setString(zeilen_zaehler, replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, trimX(akt_zeile_mark)))
                        
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, trimX(akt_zeile_mark))
                        
                        End If
                        
                    ElseIf (pFunktion = FKT_BLOCK_ZUFALL) Then
                         
                        '
                        ' Funktion "Block Zufall"
                        ' Diese Funktion erstellt aus der Eingabe einen zufaelligen Text.
                        ' Dient dazu, um Texte oder um Klarnamen verfremden zu koennen.
                        '
                        ' Soll die Markierung benutzt werden, wird nur in den Grenzen der
                        ' Markierung die Zufallsfunktion ausgefuehrt.
                        '
                        If (knz_benutze_markierung) Then
                            
                            Call cls_string_array.setString(zeilen_zaehler, getBlockZufall(aktuelle_zeile, ab_position, bis_position))
                        
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, getBlockZufall(aktuelle_zeile, -1, -1))
                        
                        End If
                    
                    ElseIf (pFunktion = FKT_DUPLIZIERUNG) Then
                        '
                        ' Funktion "Duplizierung"
                        '
                        ' Die aktuelle Zeile oder die Markierung wird der aktuellen Zeile nochmal
                        ' vor- oder hintenangestellt.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            Call cls_string_array.setString(zeilen_zaehler, akt_zeile_mark & TRENN_STRING_4 & aktuelle_zeile)
                            
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, aktuelle_zeile & TRENN_STRING_4 & akt_zeile_mark)
                            
                        End If
    
                    ElseIf (pFunktion = FKT_STRING_VERSCHIEBEN) Then
                        '
                        ' Funktion "Verschieben"
                        '
                        ' Verschieben eines Teilbereiches aus der aktuellen Zeile, einmal nach vorne
                        ' und einmal ans Ende der aktuellen Zeile.
                        '
                        temp_string_2 = getRemoveAbBis(aktuelle_zeile, ab_position, bis_position)
                
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            Call cls_string_array.setString(zeilen_zaehler, akt_zeile_mark & TRENN_STRING_5 & temp_string_2)
                            
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, temp_string_2 & TRENN_STRING_5 & akt_zeile_mark)
                            
                        End If
                    
                    ElseIf ((pFunktion = FKT_JSON_LESEN_SCHREIBEN) Or (pFunktion = FKT_NOTES_LESEN_SCHREIBEN) Or (pFunktion = FKT_JAVA_XML_WRITER_STRING) Or (pFunktion = FKT_JAVA_XML_WRITER_NUMMER)) Then
                    
                        '
                        ' Funktionen Lesen oder Schreiben in Json, XML, Notes
                        '
                    
                        temp_string_1 = Trim(akt_zeile_mark)
                        
                        If (temp_string_1 = LEER_STRING) Then
                        
                            '
                            ' Ist die aktuele Zeile/Markierung ein Leerstring, wird der Ergebnisstring
                            ' auf einen Leerstring gesetzt. Im Ergebnis werden dadurch Leerzeichen
                            ' im Ergebnisstring entfernt.
                            '
                                                    
                            temp_string_3 = LEER_STRING
                        
                        Else
                        
                            '
                            ' Ist die aktuelle Zeile kein Leerstring, wird die auszufuehrende
                            ' Generatorfunktion gesucht und in der Variablen "temp_string_3"
                            ' das Ergebnis gespeichert.
                            '

                            If (pFunktion = FKT_JSON_LESEN_SCHREIBEN) Then
                            
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    temp_string_3 = "json_erg += '\n" & temp_string_1 & " >' + AjaxErg." & temp_string_1 & " + '<';"
                                    
                                Else
                                
                                    temp_string_3 = "json_string += FkJson.getStringJson( """ & temp_string_1 & """, " & temp_string_1 & " ) + "","";" ' + "\"","";"
                                    
                                End If
                                
                            ElseIf (pFunktion = FKT_NOTES_LESEN_SCHREIBEN) Then
                            
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    temp_string_3 = "Call notesDokumentStringSet( notes_dokument, """ & temp_string_1 & """, " & temp_string_1 & " )"
                                    
                                Else
                                
                                    temp_string_3 = LCase(getKlartext(Trim(temp_string_1), UNTER_STRICH)) & " = notesDokumentStringGet( notes_dokument, """ & temp_string_1 & """ )"
                                    
                                End If
                                
                            ElseIf (pFunktion = FKT_JAVA_XML_WRITER_STRING) Then
            
                                temp_string_1 = UCase(getKlartext(temp_string_1, UNTER_STRICH))
                                
                                temp_string_2 = "#Xp" + getKlartext(temp_string_1, "")
                            
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    temp_string_3 = "xml_string.append( FkXml.getXmlTag( TAG_" & temp_string_1 & ", " & temp_string_2 & ", TAG_VORGABE_" & temp_string_1 & " );"
                                    
                                Else
                                
                                    temp_string_3 = "pBuffer.append( ""<" & temp_string_1 & ">"" + " & temp_string_2 & " + ""</" & temp_string_1 & ">"" );"
                                    
                                End If
                            
                            ElseIf (pFunktion = FKT_JAVA_XML_WRITER_NUMMER) Then
                            
                                temp_long_1 = temp_long_1 + 1
                            
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    temp_string_3 = temp_string_1 & " " & AUSRICHT_STRING_TEMP_1 & "= FkString.getXmlString( xml_root_x, """ & temp_long_1 & """);"
                                    
                                Else
                                
                                    temp_string_3 = "xml_string += ""<" & temp_long_1 & ">"" + " & temp_string_1 & " " & AUSRICHT_STRING_TEMP_1 & "+ ""</" & temp_long_1 & ">"";"
                                    
                                End If
                            
                            End If
                        
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)
                        
                    ElseIf (pFunktion = FKT_NOTES_DEBUG_FELD_WERTE) Then
                    
                        If (Trim(aktuelle_zeile) <> LEER_STRING) Then
                        
                            temp_string_1 = Trim(akt_zeile_mark)
                            
                            If (temp_string_1 <> LEER_STRING) Then
                                
                                temp_string_2 = "notesDokumentStringGet( notes_dokument, """ & temp_string_1 & """ )"
                                
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    Call cls_string_array.setString(zeilen_zaehler, "'##sss( """ & Replace(temp_string_1, """", """""") & " =>"" & " & temp_string_2 & " & ""<"" )")
                                
                                Else
                                
                                    Call cls_string_array.setString(zeilen_zaehler, "temp_str = temp_str & chr(13) & """ & Replace(temp_string_1, """", """""") & " =>"" & " & temp_string_2 & " & ""<"" ")
                                
                                End If
                            
                            End If

                        End If
                    
                    ElseIf ((pFunktion = FKT_GENERATOR_IF_JAVA_VB) Or (pFunktion = FKT_GENERATOR_IF_JAVA_SCRIPT)) Then
                    
                        If ((aktuelle_zeile <> LEER_STRING) And (ab_position > POSITION_0)) Then
                        
                            If (pSelLength > 0) Then
                            
                                temp_long_1 = InStr(aktuelle_zeile, inhalt_markierung)
                                
                                If (temp_long_1 > POSITION_0) Then
                                
                                    temp_string_1 = Left(aktuelle_zeile, temp_long_1 - 1)
                                    
                                    temp_string_2 = Mid(aktuelle_zeile, temp_long_1 + temp_long_2, Len(aktuelle_zeile))
                                
                                End If
                            
                            Else
                            
                                temp_string_1 = Left(aktuelle_zeile, ab_position - 1)
                                
                                temp_string_2 = Mid(aktuelle_zeile, ab_position, Len(aktuelle_zeile))
                                
                            End If
                            
                        Else
                        
                            temp_string_1 = Trim(aktuelle_zeile)
                            
                            temp_string_2 = aktuelle_zeile
                        
                        End If
                            
                        temp_string_1 = Trim(temp_string_1)
    
                        If (temp_string_1 = LEER_STRING) Then
                            
                            str_fkt_ergebnis = LEER_STRING
                        
                        Else
                        
                            If (pFunktion = FKT_GENERATOR_IF_JAVA_SCRIPT) Then
                                
                                str_fkt_ergebnis = zeichen_zeilenumbruch & temp_string_3 & " ( " & temp_string_1 & " == " & temp_string_2 & " ) " & zeichen_zeilenumbruch & "{" & zeichen_zeilenumbruch & "}"
                                
                                temp_string_3 = "else if"
                      
                            Else
                                
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "JAVA_IF1  " & temp_string_3 & " ( str_parameter_name.equalsIgnoreCase( """ & temp_string_1 & """ ) ) "
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "JAVA_IF2  " & temp_string_3 & " ( str_parameter_name == " & temp_string_1 & " ) "
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "{"
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "JAVA_TH1  " & "  str_lokale_variable  = """ & temp_string_2 & """; "
                                    
                                    If (Right(temp_string_2, 1) = ";") Then
                                        
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "JAVA_TH2  " & "  str_lokale_variable  = " & temp_string_2 & ""
                                    
                                    Else
                                        
                                        str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "JAVA_TH2  " & "  str_lokale_variable  = " & temp_string_2 & ";"
                                    
                                    End If
                                    
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "}"
                                    
                                    temp_string_3 = "else if"
                                
                                Else
                                   
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "VB_IF1" & temp_string_3 & " ( str_parameter_name  = " & temp_string_1 & ") Then"  '#1
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "VB_IF2" & temp_string_3 & " ( str_parameter_name  = """ & temp_string_1 & """) Then"  '#1
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "VB_TH1  " & "      str_lokale_variable  = " & temp_string_2 & " "
                                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & "VB_TH2  " & "      str_lokale_variable  = """ & temp_string_2 & """"
                                    
                                    temp_string_3 = "ElseIf"
        
                                End If
    
                            End If
    
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, str_fkt_ergebnis)
                        
                        str_fkt_ergebnis = LEER_STRING
    
                    ElseIf (pFunktion = FKT_JAVA_GENERATOR) Then
    
                        '
                        ' Anforderungen an Generatorquelltext
                        ' Slasches und Anfuehrungszeichen qualifizieren, da diese nachher in einem String stehen.
                        '
                        aktuelle_zeile = Replace(Replace(trimTail(aktuelle_zeile), "\", "\\"), """", "\""")
                        
                        '
                        ' Leerzeichen Klammern
                        ' Nach jeder oefnenden Klammer wird ein Leerzeichen eingefuegt.
                        ' Vor jeder schliessenden Klammer wird ein Leerzeichen eingefuegt.
                        ' Runde-, Eckige- und geschweifte Klammern.
                        '
                        aktuelle_zeile = Replace(Replace(Replace(Replace(aktuelle_zeile, "(", "( "), ")", " )"), "[", "[ "), "]", " ]")
                        
                        '
                        ' Eliminierung von doppelten Leerzeichen nach Klammern
                        ' Fuer jede oeffnende Klammer mit 2 nachfolgenden Leerzeichen, wird
                        ' eine Klammer mit nur einem nachfolgendem Leerzeichen. Gleiches wird
                        ' fuer die schliessenden Klammern gemacht.
                        '
                        ' Desweiteren werden Klammern fuer Funktionsaufrufe zusammengezogen.
                        ' Aus "(  )" und "( )" wird "()".
                        '
                        aktuelle_zeile = Replace(Replace(Replace(Replace(Replace(Replace(aktuelle_zeile, "(  )", "()"), "[  ]", "[]"), "(  ", "( "), "  )", " )"), "[  ", "[ "), "  ]", " ]")
                        
                        aktuelle_zeile = Replace(Replace(Replace(Replace(Replace(Replace(aktuelle_zeile, "( )", "()"), "[ ]", "[]"), "(  ", "( "), "  )", " )"), "[  ", "[ "), "  ]", " ]")
                        
                        '
                        ' Praefix und Suffix
                        ' Die umgestellte Zeile wird mit dem praefix "pBuffer.append( """ und
                        ' dem Suffix """ );" versehen.
                        '
                    
                        aktuelle_zeile = "pBuffer.append( """ & aktuelle_zeile & """ );"
                        
                        '
                        ' Ersetzungen fuer VB6 rausgenommen
                        'If (m_toggle_mr_stringer_fkt) Then
                        '
                        '    aktuelle_zeile = Replace(aktuelle_zeile, "( LEER_STRING    ", "( TAB_STR + """)
                        '    aktuelle_zeile = Replace(aktuelle_zeile, " TAB_STR + ""    ", " TAB_STR + TAB_STR + """)
                        '    aktuelle_zeile = Replace(aktuelle_zeile, " TAB_STR + ""    ", " TAB_STR + TAB_STR + """)
                        '    aktuelle_zeile = Replace(aktuelle_zeile, " TAB_STR + ""    ", " TAB_STR + TAB_STR + """)
                        '    aktuelle_zeile = Replace(aktuelle_zeile, " TAB_STR + ""    ", " TAB_STR + TAB_STR + """)
                        '
                        '    aktuelle_zeile = Replace(aktuelle_zeile, "End If", """ + STR_VB_END_IF + """)
                        '    aktuelle_zeile = Replace(aktuelle_zeile, "ElseIf", """ + STR_VB_ELSE_IF + """)
                        '    aktuelle_zeile = Replace(aktuelle_zeile, "End Function", """ + STR_VB_END_FUNCTION + """)
                        '    aktuelle_zeile = Replace(aktuelle_zeile, "End Sub", """ + STR_VB_END_SUB + """)
                        '    aktuelle_zeile = Replace(aktuelle_zeile, " As ", """ + LZ + STR_VB_AS_TYPE + LZ + """)
                        '    aktuelle_zeile = Replace(aktuelle_zeile, "Public Function ", """ + STR_VB_PUBLIC + LZ + STR_VB_FUNCTION + LZ + """)
                        '
                        '    aktuelle_zeile = Replace(aktuelle_zeile, " = ", """ + LZ + STR_VB_ZUWEISUNG + LZ + """)
                        '    aktuelle_zeile = Replace(aktuelle_zeile, "+ """" +", "+")
                        '    aktuelle_zeile = Replace(aktuelle_zeile, " + """" );", " );")
                        '    aktuelle_zeile = Replace(aktuelle_zeile, "( """" +", "(")
                        '
                        'End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, aktuelle_zeile)

                    ElseIf (pFunktion = FKT_GET_DIR) Then
                        '
                        ' Funktion "Verzeichnis einlesen"
                        ' Liest das Verzeichis ein, welches in der aktuellen Zeile steht.
                        '
                        aktuelle_zeile = Trim(aktuelle_zeile)
                        
                        If (aktuelle_zeile <> LEER_STRING) Then
                        
                            '
                            ' Sicherstellung, dass der Pfad mit einem Slash endet
                            '
                            If (InStr("\/", Right(aktuelle_zeile, 1)) <= POSITION_0) Then
                            
                                aktuelle_zeile = aktuelle_zeile & "\"
                                
                            End If
                            
                            '
                            ' Bei jedem zweiten Aufruf wird der Verzeichnisname dem Dateienamen vorangestellt.
                            '
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                temp_string_2 = aktuelle_zeile
                                
                            Else
                            
                                temp_string_2 = LEER_STRING
                                
                            End If
                            
                            '
                            ' While-Schleife fuer das Einlesen der Dateien starten.
                            ' Es werden maximal 32123 Dateien gelesen.
                            '
                            temp_long_1 = 0
                            
                            temp_string_1 = Dir(aktuelle_zeile & "*.*")
    
                            While ((temp_string_1 <> LEER_STRING) And (temp_long_1 < 32123))
                            
                                str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & temp_string_2 & temp_string_1
                                
                                temp_string_1 = Dir()
    
                            Wend
                            
                        End If
                 
                    End If
                    
                    '
                    ' Zeilenzaehler erhoehen
                    ' Am Ende der IF-Kaskade wird der Zeilenzaehler fuer den naechsten Durchgang um 1 erhoeht.
                    '
                    zeilen_zaehler = zeilen_zaehler + 1
                
                Wend
                
                '
                ' Abschlussarbeiten nach While-Schleife
                ' Vervollstaendigung der Ausgabe um weitere Praefixe und Suffixe.
                '
                If (pFunktion = FKT_ERSTELLE_XML) Or (pFunktion = FKT_ERSTELLE_XML_2) Then
                
                    '
                    ' Funktion "XML"
                    ' Dem Ergebnis wird noch der Vorlauf und der Nachlauf (abschliessende Klammer) hinzugefuegt.
                    '
                    str_fkt_ergebnis = "<?xml version=""1.0"" encoding=""UTF-8""?>" & zeichen_zeilenumbruch & "<XML_KLAMMER_1>" & zeichen_zeilenumbruch & str_fkt_ergebnis & "</XML_KLAMMER_1>" & zeichen_zeilenumbruch
    
                ElseIf (pFunktion = FKT_CALC_SUMME) Then
                
                    '
                    ' Funktion "Summe"
                    ' Das Summenergebnis wird in die erste Zeile geschrieben.
                    ' Dieses hat den Grund, dass dann nicht gescrollt werden muss und zweitens auch das Summenergebnis
                    ' in der Ausgabebox sichtbar ist (32K Grenze des VB-Anzeigeconotrols).
                    '
                    If (temp_long_1 > 0) Then
                    
                        temp_double_1 = temp_double_2 / temp_long_1
                    
                    Else
                    
                        temp_double_1 = 0#
                        
                    End If
                    
                    str_fkt_ergebnis = "ERGEBNIS >" & temp_long_1 & "<  >" & temp_double_2 & "< Durchschnitt >" & temp_double_1 & "<" & zeichen_zeilenumbruch & str_fkt_ergebnis
                
                ElseIf (pFunktion = FKT_GET_EINMALIGE_VORKOMMEN) Then
                    '
                    ' Funktion "Einmalige Vorkommen"
                    ' Zeilenumbruch fuer das temporaere "#1#", sowie ein Leerstring fuer "#2#".
                    '
                    str_fkt_ergebnis = Replace(Replace(temp_string_2, "#1#", zeichen_zeilenumbruch), "#2#", "")
                 
                ElseIf ((pFunktion = FKT_GETTER_SETTER_JAVA) Or (pFunktion = FKT_GETTER_SETTER_VB) Or (pFunktion = FKT_SINGLETON_JAVA)) Then
                            
                    str_fkt_ergebnis = temp_string_3 & zeichen_zeilenumbruch & str_fkt_ergebnis & zeichen_zeilenumbruch
                
                ElseIf ((pFunktion = FKT_CSV_VB_KONVERTER)) Then
                            
                    str_fkt_ergebnis = zeichen_zeilenumbruch & str_fkt_ergebnis & zeichen_zeilenumbruch & "    EndIf" & zeichen_zeilenumbruch

                ElseIf (Len(str_fkt_ergebnis) = 0) Then
                    '
                    ' Bestimmung Ergebnis aus Zeilenarray
                    ' Ist die Variable "str_fkt_ergebnis" noch nicht mit einem Wert versehen worden,
                    ' bestimmt sich das Funktionsergebnis aus allen Zeilen des Zeilenarrays.
                    ' In diesem Fall veraendern die Funktionen die gespeicherten Zeilen im Objekt.
                    '
                    ' Andere Funktionen benutzten gleich die Variable "str_fkt_ergebnis".
                    '
                    str_fkt_ergebnis = cls_string_array.toString(zeichen_zeilenumbruch)
                    
                End If
                
                If (pFunktion = FKT_GENERATOR_STRING_IT) Then
            
                    If (m_zaehler_string_it = 2) Then
                    
                        str_fkt_ergebnis = "String j_str  = """";" & zeichen_zeilenumbruch & zeichen_zeilenumbruch & str_fkt_ergebnis
                    
                    End If
                    
                ElseIf (pFunktion = FKT_STRING_LAENGE_AUSGEBEN) Then
                
                    str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch & zeichen_zeilenumbruch & "Gesamt " & Len(pString)
                        
                End If

            End If

        End If

    End If

    '
    ' Kennzeichen "Aktiv" wird auf FALSE gestellt
    '
    m_knz_aktiv = False
    
    '
    ' Das Funktionsergebnis wird gesetzt.
    '
    startMrStringer = str_fkt_ergebnis
    
EndFunktion:
    
    '
    ' Die String-Array-Instanz wird auf Nothing gesetzt
    '
    Set cls_string_array = Nothing
    
    '
    ' DoEvents aufrufen
    '
    DoEvents
    
    '
    ' Die Funktion wird verlassen
    '
    Exit Function
    
errStartMrStringer:
    
    startMrStringer = "Fehler: " & Error
    
    Resume EndFunktion

End Function

'################################################################################
'
Public Function startJoin(pStringA As String, pStringB As String, pTrennzeichen As String, Optional knz_restart_zaehler_2 As Boolean = False) As String

On Error GoTo errStartJoin

Dim cls_string_array_1 As clsStringArray
Dim cls_string_array_2 As clsStringArray
Dim str_fkt_ergebnis   As String
Dim str_my_cr          As String
Dim zeilen_anzahl_1    As Long
Dim zeilen_anzahl_2    As Long
Dim zeilen_index_1     As Long
Dim zeilen_index_2     As Long

    m_toggle_mr_stringer_fkt = Not m_toggle_mr_stringer_fkt
    
    '
    ' Join-Reihenfolge
    '
    ' Bei jedem Aufruf aendert sich die Reihenfolge, ob Text A oder Text B zuerst kommt.
    ' Je nach dem Toggle-Kennzeichen werden die Texte unterschiedlich den String-Arrays zugeordnet.
    ' Dieses Vorgehen vereinfacht die Anweisungen in der While-Schleife.
    '
    If (m_toggle_mr_stringer_fkt) Then
                
        Set cls_string_array_1 = startMultiline(pStringA)
        
        Set cls_string_array_2 = startMultiline(pStringB)
            
    Else
                    
        Set cls_string_array_1 = startMultiline(pStringB)
        
        Set cls_string_array_2 = startMultiline(pStringA)
        
    End If
    
    If ((cls_string_array_1 Is Nothing) And (cls_string_array_2 Is Nothing)) Then
    
        '
        ' Sind beide String-Arrays gleich "Nothing", ist das Ergebnis ein Leerstring
        '
        startJoin = LEER_STRING
        
    ElseIf (cls_string_array_2 Is Nothing) Then
    
        '
        ' Ist der zweite String-Array gleich "Nothing", wird pStringA zurueckgegeben
        ' Es muss kein Join gemacht werden.
        '
        startJoin = pStringA
    
    ElseIf (cls_string_array_1 Is Nothing) Then
    
        '
        ' Ist der erste String-Array gleich "Nothing", wird pStringB zurueckgegeben
        ' Es muss kein Join gemacht werden.
        '
        startJoin = pStringB
        
    Else
        '
        ' Sind beide String-Arrays ungleich "Nothing", werden die beiden
        ' String-Arrays zusammengefuehrt.
        '
        
        '
        ' Das Zeilenumbruchszeichen wird auf einen Leerstring gesetzt.
        '
        str_my_cr = LEER_STRING
        
        '
        ' Aus beiden String-Arrays werden die Zeilenanzahlen ermittelt.
        '
        zeilen_anzahl_1 = cls_string_array_1.getAnzahlStrings
        zeilen_anzahl_2 = cls_string_array_2.getAnzahlStrings
        
        '
        ' Startwerte fuer die Zeilen-Index-Zaehler auf die erste Zeile setzen.
        '
        zeilen_index_1 = 1
        zeilen_index_2 = 1
        
        '
        ' While-Schleife
        '
        While (((zeilen_index_1 <= zeilen_anzahl_1) Or (zeilen_index_2 <= zeilen_anzahl_2)) And (zeilen_index_1 < 31213))
        
            '
            ' Ergebnisaufbau ist, dass der erste Text aus Array 1 kommt und der zweite aus Array 2.
            '
            
            str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr & cls_string_array_1.getString(zeilen_index_1) & pTrennzeichen & cls_string_array_2.getString(zeilen_index_2)
            
            '
            ' Zeilenumbruch auf "MY_CHR_13_10" setzen.
            '
            str_my_cr = MY_CHR_13_10
            
            '
            ' Index-Zaehler erhoehen
            '
            zeilen_index_1 = zeilen_index_1 + 1
            zeilen_index_2 = zeilen_index_2 + 1
        
            '
            ' Pruefung: Soll ein Restart von Zaehler 2 gemacht werden ?
            '
            If (knz_restart_zaehler_2) Then
            
                '
                ' Der Zaehler fuer die zweite Spalte wird zurueckgesetzt,
                ' wenn der Zaehler fuer die erste Spalte noch nicht am
                ' Ende der Anzahl Zeilen 1 angekommen ist.
                '
                ' Solange noch Zeilen in der ersten Spalte vorhanden sind,
                ' soll sich die zweite Spalte wiederholen.
                '
                If ((zeilen_index_1 <= zeilen_anzahl_1) And (zeilen_index_2 > zeilen_anzahl_2)) Then
                    
                    zeilen_index_2 = 1
                
                End If
                
            End If
            
        Wend
        
        '
        ' Nach der While-Schleife wird das Funktionsergebnis gesetzt
        '
        startJoin = str_fkt_ergebnis

    End If
    
EndFunktion:

    On Error Resume Next

    '
    ' Die verwendeten Klassen werden auf "Nothing" gesetzt.
    '
    Set cls_string_array_1 = Nothing
    Set cls_string_array_2 = Nothing
    
    '
    ' DoEvents wird aufgerufen.
    '
    DoEvents
    
    Exit Function
    
errStartJoin:

    Resume EndFunktion
    
End Function

'################################################################################
'
' Formatiert den uebergebenen JSON-String.
'
' https://stackoverflow.com/questions/4105795/pretty-print-json-in-java/7310424
'
' PARAMETER: pJsonString der zu formatierende Json-String
'
' RETURN: den formatierten Json-String
'
Public Function formatJsonString(pJsonString As String) As String

Dim str_fkt_ergebnis As String  ' Ergebnisstring fuer die Rueckgabe
Dim str_tab_einzug   As String  ' Der als Tabulator-Einzug zu benutzende String
Dim str_my_cr        As String  ' Das in dieser Funktion verwendete Zeilenumbruchszeichen
Dim akt_zeichen      As String  ' aktuelle Zeichen in der While-Schleife
Dim akt_position     As Integer ' aktuelle Leseposition der While-Schleife
Dim einzug_anzahl    As Integer ' Die Anzahl der TABs fuer den Einzug der Elemente
Dim knz_in_string    As Boolean ' Kennzeichen, ob sich der Leseprozess im String befindet

    '
    ' Das in dieser Funktion verwendete Zeilenumbruchszeichen
    ' wird auf "CR + LF" gestellt.
    '
    str_my_cr = vbCrLf
    
    '
    ' Der hinzuzufuegende Tabulator wird auf 2 Leerzeichen gesetzt
    '
    str_tab_einzug = "  "
    
    '
    ' Die aktuelle Lesepositon wird auf das erste Zeichen gestellt.
    '
    akt_position = 1

    '
    ' While-Schleife ueber alle Zeichen der Eingabe
    '
    While (akt_position <= Len(pJsonString))

        '
        ' Zeichen aus der Eingabe an der aktuellen Lesepositon lesen
        '
        akt_zeichen = Mid(pJsonString, akt_position, 1)
        
        If (akt_zeichen = """") Then
            
            '
            ' Anfuehrungszeichen
            '
            ' Ein Anfuehrungszeichen schaltet das Kennzeichen "knz_in_string" um.
            '
            ' Sollen nur Teil-JSON-Strings formatiert werden, kann dieses zu
            ' Fehlern in der Ausgabe fuehren. Diese Funktion ist so ausgelegt,
            ' dass sie sich selber wieder synchronisiert, indem Klammern das
            ' Kennzeichen "knz_in_string" wieder auf FALSE stellen.
            '
            
            knz_in_string = Not knz_in_string
            
            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
        
        ElseIf (akt_zeichen = vbCr) Then
        
            '
            ' Zeilenumbruch ueberlesen
            '
        
        ElseIf (akt_zeichen = vbLf) Then
        
            '
            ' Linefeed ueberlesen
            '
        
        ElseIf (akt_zeichen = ":") Then
        
            '
            ' Doppelpunkt
            '
            ' Befindet sich der Doppelpunkt in einem String, wird
            ' nur der Doppelpunkt uebernommen.
            '
            ' Befindet sich der Doppelpunkt ausserhalb eines Strings,
            ' wird dieser mit 2 Leerzeichen umschlossen. Diese beiden
            ' Leerzeichen wuerden sonst von dieser Funktion eleminiert
            ' werden. Sie dienen der besseren Uebersichtlichkeit.
            '
            
            If (knz_in_string) Then
            
                str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
                
            Else
            
                str_fkt_ergebnis = str_fkt_ergebnis & " " & akt_zeichen & " "
                
            End If
        
        ElseIf (akt_zeichen = " ") Then
            
            '
            ' Leerzeichen
            '
            ' Leerzeichen in Strings werden uebernommen.
            '
            ' Leerzeichen ausserhalb von Strings werden ueberlesen.
            '
            
            If (knz_in_string) Then
            
                str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
                
            End If
        
        ElseIf ((akt_zeichen = "{") Or (akt_zeichen = "[")) Then
        
            '
            ' Oeffnende Klammern
            '
            ' Eine oeffnende Klammer erstellt ein neues Objekt oder Array.
            ' Nach der Klammer wird ein Zeilenumbruch eingefuegt.
            ' Der Tabulatoreinzug wird erhoeht. Damit im Ergebnisstring
            ' das naechste Zeichen an der korrekten Einzuposition steht,
            ' wird die neue Zeile mit dem Tabulatoreinzug aufgefuellt.
            '
            ' - beendet einen String
            ' - erhoeht die Tab-Einzugsanzahl
            ' - die Einzugsanzahl darf nicht groesser als 200 werden
            '
            ' - Aktuelle Zeichen + CR + TAB-Einzug
            '
            
            knz_in_string = False
            
            einzug_anzahl = einzug_anzahl + 1
            
            If (einzug_anzahl > 200) Then
            
                einzug_anzahl = 200
                
            End If
            
            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
            
            str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr
            
            str_fkt_ergebnis = str_fkt_ergebnis & getStringXmal(str_tab_einzug, einzug_anzahl)

        ElseIf ((akt_zeichen = "}") Or (akt_zeichen = "]")) Then
            
            '
            ' Schliessende Klammern
            '
            ' Eine schliessende Klammer beendet ein Objekt oder Array.
            ' Die aktuelle Zeile wird mit einem Zeilenumbruch abgeschlossen.
            ' Die Tabulatoreinzugsanzahl vermindert sich. Damit die
            ' schliessende Klammer (=aktuelles Zeichen) an der korrekten
            ' Einzugsposition erscheint, wird der neuen Zeile der Einzug
            ' aus den Tabulatoren hinzugefuegt.
            '
            ' - beendet einen String
            ' - vermindert die Tab-Einzugsanzahl
            ' - die Einzugsanzahl darf nicht negativ werden
            '
            ' - CR + TAB-Einzug + Aktuelle Zeichen
            '
            
            knz_in_string = False
            
            einzug_anzahl = einzug_anzahl - 1
            
            If (einzug_anzahl < 0) Then
            
                einzug_anzahl = 0
                
            End If
            
            str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr
            
            str_fkt_ergebnis = str_fkt_ergebnis & getStringXmal(str_tab_einzug, einzug_anzahl)

            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
            
        ElseIf (akt_zeichen = ",") Then
        
            '
            ' Komma
            '
            ' - das Zeichen wird in die aktuelle Zeile uebernommen
            ' - ausserhalb eines Stringes beedet das Komma ein Element.
            '   Es wird ein CR und ein Tab-Einzug hinzugefuegt.
            '

            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
            
            If (knz_in_string = False) Then
            
                str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr
                
                str_fkt_ergebnis = str_fkt_ergebnis & getStringXmal(str_tab_einzug, einzug_anzahl)
                
            End If
        
        Else
        
            '
            ' Zeichen ohne JSON-Sonderfunktion werden in den Ergebnisstring uebernommen
            '
            
            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen

        End If

        '
        ' Es wird die Leseposition der While-Schleife erhoeht und
        ' mit dem naechsten Schleifedurchlauf weitergemacht.
        '
        akt_position = akt_position + 1

    Wend

    '
    ' Der Aufrufer bekommt am Funktionsende den Ergebnisstring zurueck.
    '
    formatJsonString = str_fkt_ergebnis

End Function

'########################################################################################
'
'    Eingabe   <XML_KLAMMER_1><EINS /><ZWEI></ZWEI><DREI><A></A><B/></DREI><!-- Kommentar//--></XML_KLAMMER_1>
'
'    Ausgabe   <XML_KLAMMER_1>
'                <EINS />
'                <ZWEI></ZWEI>
'                <DREI>
'                  <A></A>
'                  <B/>
'                </DREI>
'                <!-- Kommentar//-->
'              </XML_KLAMMER_1>
'
' PARAMETER: pString = Eingabe-XML
'
' RETURN : die formatierte zeilenweise Darstellung des Eingabe-XMLs
'
Public Function formatXML(pString As String) As String

    formatXML = pString
    
    If (pString = LEER_STRING) Then
        
        Exit Function
        
    End If
    
Dim tag_art_start                   As Integer ' Tag-Art fuer ein Start-Tag
Dim tag_art_ende_normal             As Integer ' Tag-Art fuer ein End-Tag <tag />
Dim tag_art_start_ende__kommentar   As Integer ' Tag-Art fuer ein kombiniertes Start-End-Tag (... oder Kommentaren)
Dim knz_tag_art                     As Integer ' Kennzeichen fuer die aktuelle Tag-Art
Dim knz_letzte_tag_art              As Integer ' Kennzeichen fuer die letzte Tag-Art
Dim eingabe_str                     As String  ' der konvertierte Eingabestring (enthaelt keine Zeilenumbrueche)
Dim start_zeichen                   As String  ' Das Startzeichen "<"
Dim end_zeichen                     As String  ' Das Endzeichen ">"
Dim str_fkt_ergebnis                As String  ' Der Stringbuffer fuer die Aufnahme des Ergebnisses
Dim lese_position_start             As Long    ' Die aktuelle Startleseposition
Dim lese_position_ende              As Long    ' Die aktuelle Endleseposition
Dim letztes_end_tag_ende            As Long    ' Die Position wo das letzte Tag endete
Dim pos_akt_xml_start_zeichen       As Long    ' Die aktuelle Position des Startzeichens
Dim pos_akt_xml_end_zeichen         As Long    ' Die aktuelle Position des Endzeichens
Dim zaehler_while_schleife          As Long    ' Endlosschleifenverhinderungszaehler
Dim len_eingabe_str                 As Long    ' Laenge des Eingabestrings
Dim einrueck_zaehler                As Integer ' aktuelle Einrueckbreite
Dim einrueck_breite                 As Integer ' Einrueckeinheitenbreite
Dim knz_trennen                     As Boolean ' Kennzeichen ob getrennt werden soll
Dim str_debug                       As String  ' Ein optionaler Debug-String
Dim akt_teil_string                 As String
Dim pos_next_slash                  As Long

    '
    ' Alle Kombinationen aus Zeilenumbruechen aus dem Eingabestring entfernen
    '
    eingabe_str = Replace(Replace(pString, MY_CHR_13_10, ""), Chr(13), "")
    len_eingabe_str = Len(eingabe_str)
    einrueck_zaehler = 0
    einrueck_breite = 2
    start_zeichen = "<"
    end_zeichen = ">"
    lese_position_start = 1
    lese_position_ende = 1
    letztes_end_tag_ende = -1
    knz_trennen = False
    str_debug = LEER_STRING
    
    '
    ' Initialisierung der Tag-Arten (Werte zuweisen)
    '
    tag_art_start = 1
    tag_art_ende_normal = 2
    tag_art_start_ende__kommentar = 3
    
    '
    ' Position des ersten XML-Startzeichens "<" suchen.
    '
    pos_akt_xml_start_zeichen = InStr(1, eingabe_str, start_zeichen)
    
    '
    ' Pruefung: Konnte ein XML-Startzeichen gefunden werden?
    '
    ' Kein XML-Startzeichen --> Rueckgabe der Eingabe --> Ende
    '
    If (pos_akt_xml_start_zeichen = 0) Then
        
        formatXML = pString
        
        Exit Function
        
    End If
    '
    ' Position des ersten XML-Endzeichens ">" suchen.
    '
    pos_akt_xml_end_zeichen = InStr(1, eingabe_str, end_zeichen)
    
    '
    ' Pruefung: Liegt ein Enzeichen vor dem ersten Startzeichen ?
    ' Diese Bedingung kann nur bei einer unvollstaendig uebergebenen XML-Notation
    ' passieren. Es muss dann geprueft werden, bis zu welcher Position der unvoll-
    ' staendige Teil der ersten Ergebniszeile geht. Kommt nach dem ersten gefundenen
    ' XML-Startzeichen ein "/", muss ab dem Startzeichen noch die Position des
    ' Endzeichens ermittelt werden. Dieses ist z.B. hier der Fall
    '
    ' "ehlt>fd_vorname_as</fehlt><fe"
    ' "ehlt>fd_vorname_as</fehlt></fehlt>"
    '
    ' aber nicht bei:
    '
    ' "ntrag><antrag_aktionskennzeichen></antrag_aktions"
    '
    If (pos_akt_xml_end_zeichen < pos_akt_xml_start_zeichen) Then
        '
        ' Pruefung: Start-Xml-Zeichen gehoert zu Endtag ?
        '
        If (Mid(eingabe_str, pos_akt_xml_start_zeichen + 1, 1) = "/") Then
        
            pos_akt_xml_end_zeichen = InStr(pos_akt_xml_start_zeichen, eingabe_str, end_zeichen)
            
            If (pos_akt_xml_end_zeichen = 0) Then
                
                pos_akt_xml_end_zeichen = len_eingabe_str
            
            End If
            '
            ' Kennzeichen fuer Art des letzten Tags auf "End-Tag-Normal" stellen
            '
            knz_letzte_tag_art = tag_art_ende_normal
            
        Else
        
            knz_letzte_tag_art = tag_art_start
        
        End If
        
        lese_position_ende = pos_akt_xml_end_zeichen
    
        str_fkt_ergebnis = str_fkt_ergebnis & Mid(eingabe_str, lese_position_start, (lese_position_ende - lese_position_start) + 1) & MY_CHR_13_10
        
        lese_position_start = lese_position_ende + 1
        
        If (pos_akt_xml_start_zeichen < len_eingabe_str) Then
            
            pos_akt_xml_start_zeichen = InStr(pos_akt_xml_start_zeichen + 1, eingabe_str, start_zeichen)
            
            If (pos_akt_xml_start_zeichen = 0) Then
            
                str_fkt_ergebnis = str_fkt_ergebnis & Trim(Mid(eingabe_str, lese_position_start, len_eingabe_str)) & MY_CHR_13_10
            
            End If
        
        End If
        
    End If
    
    '
    ' Die Schleife wird ausgefuehrt solange noch das Startzeichen gefunden wird
    ' und der Endllosschleifenverhinderungszaehler noch kleiner als 32000 ist.
    '
    While ((pos_akt_xml_start_zeichen > 0) And (zaehler_while_schleife < 32000))
    
        knz_tag_art = tag_art_start
        '
        ' Pruefung: Startzeichen gefunden?
        '
        If (pos_akt_xml_start_zeichen > 0) Then
        
            '
            ' das Endzeichen ab der Startposition suchen
            '
            pos_akt_xml_end_zeichen = InStr(pos_akt_xml_start_zeichen + 1, eingabe_str, end_zeichen)
            
            '
            ' Pruefung: konnte Endzeichen ab Start gefunden werden ?
            '
            ' Kann das Endzeichen ab dem Start nicht gefunden werden, weil der ....
            ' wird die Position auf das Stringende gesetzt.
            '
            If (pos_akt_xml_end_zeichen = 0) Then
                
                pos_akt_xml_end_zeichen = len_eingabe_str
            
            End If
            
            '
            ' Pruefung:  Endzeichen gefunden?
            ' Bedingung: Endzeichen ">" liegt hinter Startzeichen "<"
            '
            If (pos_akt_xml_end_zeichen > pos_akt_xml_start_zeichen) Then
                '
                ' Pruefung:  End-Tag Variante 1 - Normal "</end_tag>"
                ' Bedingung: Liegt 1 Zeichen nach dem  Startzeiten ein "/"  ?
                '
                ' Das Kennzeichen fuer die Tag-Art wird auf "tag_art_ende_normal" gestellt.
                '
                If (Mid(eingabe_str, pos_akt_xml_start_zeichen + 1, 1) = "/") Then
                    
                    knz_tag_art = tag_art_ende_normal
                    
                Else
                    '
                    ' Fehlertoleranz fuer XML-Fehler:
                    '
                    ' Position des naechsten Slashes suchen.
                    '
                    pos_next_slash = InStr(pos_akt_xml_start_zeichen + 1, eingabe_str, "/")
                    
                    '
                    ' Pruefung: naechster Slash innerhalb aktueller Tag-Grenzen ?
                    '
                    ' Beispiel:
                    ' <XML_KLAMMER_1><b><a></a><a></a><a><<</a><a></a></b></XML_KLAMMER_1>
                    '
                    ' Befindet sich ein Slash in den aktuellen Grenzen des TAGs,
                    ' handelt sich sich um ein abschliessendes TAG.
                    '
                    ' Vermeided fehlerhaftes Trennen und Einruecken. Duerfte fuer die meisten
                    ' Faelle reichen.
                    '
                    If (pos_next_slash > pos_akt_xml_start_zeichen) And (pos_next_slash < pos_akt_xml_end_zeichen) Then
                    
                        knz_tag_art = tag_art_ende_normal
                        
                    End If
                
                End If
                
                '
                ' Pruefung:  End-Tag Variante 2 "<end_tag/>"
                ' Bedingung: Liegt 1 Zeichen vor dem Ende ein "/"  ?
                '
                ' Das Kennzeichen fuer die Tag-Art wird auf "tag_art_start_ende__kommentar" gestellt.
                '
                If (Mid(eingabe_str, pos_akt_xml_end_zeichen - 1, 1) = "/") Then
                
                    knz_tag_art = tag_art_start_ende__kommentar
                    
                End If
                
                '
                ' Pruefung:  Kommentar  "<!-- Kommentar -->"
                ' Bedingung: Liegt 1 Zeichen nach dem  Startzeiten ein "!"  ?
                '
                ' Wird wie ein Start-End-Tag-Behandelt, da zwischen erster und letzter
                ' Klammer der gesamte Inhalt steht.
                '
                If (Mid(eingabe_str, pos_akt_xml_start_zeichen + 1, 1) = "!") Then
                    
                    knz_tag_art = tag_art_start_ende__kommentar
                
                End If
                
            End If
    
        End If
        
        '
        ' Sonderfall:
        ' <?xml version="1.0" encoding="UTF-8"?><XML_KLAMMER_1><EINS /><ZWEI /><a></a><DREI /><VIER /></XML_KLAMMER_1>
        '
        ' Sonderfall ist, dass ein Start-Ende-Tag nach einer oeffnenden Klammer
        ' kommt. Das Start-Tag muss noch in die Ergebnismenge uebertragen werden.
        '
        If ((knz_letzte_tag_art = tag_art_start) And (knz_tag_art = tag_art_start_ende__kommentar)) Then
        
            'str_debug = "A - "
            lese_position_ende = letztes_end_tag_ende
        
            str_fkt_ergebnis = str_fkt_ergebnis & str_debug & String(einrueck_zaehler, " ") & Trim(Mid(eingabe_str, lese_position_start, (lese_position_ende - lese_position_start) + 1)) & MY_CHR_13_10
            
            '
            ' Leseposition-Start aktualisieren
            ' Die naechste Startposition fuer den Leseprozess liegt hinter dem aktuellen Endezeichen.
            ' Das ist genau die Startposition fuer das aktuelle XML-Tag
            '
            lese_position_start = lese_position_ende + 1
            
            '
            ' Einrueckung
            ' Nach einer oeffnenden XML-klammer (=letztes TAG) muss die Einrueckung
            ' fuer die aktuelle XML-Klammer vor dem evtl. naechsten Trennen erhoeht
            ' werden.
            '
            If (einrueck_zaehler < 200) Then
            
                einrueck_zaehler = einrueck_zaehler + einrueck_breite

            End If

        End If
        
        '
        ' Pruefung: Start-End-Tag gefunden?
        '
        If (knz_tag_art = tag_art_start_ende__kommentar) Then
            
            'str_debug = "B - "
            '
            ' End-Tag gefunden
            ' Bei einem End-Tag wird immer getrennt.
            '
            knz_trennen = True
            
            '
            ' Die Endposition ist die aktuelle Position des XML-Endzeichens
            '
            lese_position_ende = pos_akt_xml_end_zeichen
        '
        ' Pruefung: Normales End-Tag gefunden?
        '
        ElseIf (knz_tag_art = tag_art_ende_normal) Then
            
            '
            'str_debug = "C - "
            '
            ' End-Tag gefunden
            ' Bei einem End-Tag wird immer getrennt.
            '
            knz_trennen = True
            
            '
            ' Die Endposition ist die aktuelle Position des XML-Endzeichens
            '
            lese_position_ende = pos_akt_xml_end_zeichen
            
            '
            ' Einrueckung
            ' Bei der Kombination "End-Tag" folgt auf "End-Tag" wird der Einzug um die
            ' eingestellte Breite reduziert. Der Einrueckzaehler darf die Untergrenze
            ' von 0 nicht unterschreiten, welches durch eine Pruefung verhindert wird.
            '
            ' Die Einrueckung wird nur bei normalen End-Tags "</TAG>" gemacht. Bei
            ' Start-End-Tags waere es falsch, da dann zwei aufeinanderfolgende Tags
            ' den Einzug schnell verkleinern wuerden.
            '
            If (knz_letzte_tag_art <> tag_art_start) Then
                
                einrueck_zaehler = einrueck_zaehler - einrueck_breite
                
                If (einrueck_zaehler < 0) Then
                
                    einrueck_zaehler = 0
                    
                End If
            
            End If
        '
        ' Pruefung auf Start-Tag folgt Start-Tag
        '
        ElseIf ((knz_tag_art = tag_art_start) And (knz_letzte_tag_art = tag_art_start)) Then
            '
            'str_debug = "D - "
            '
            '
            ' Bei der Kombination "Start-Tag" folgt "Start-Tag" wird getrennt
            ' Das Ende fuer die Leseposition ist dabei das Ende des letzten
            ' Start-Tags.
            '
            knz_trennen = True
            
            '
            ' Endposition
            '
            ' Ist die Start-Position des XML-Tags kleiner als die Laenge des Eingabestrings,
            ' wird bis zum Ende des letzten Tag-Zeichens abgeschnitten.
            '
            ' Sonderfall, wenn das naechste Starzeichen gleich der Stringlaenge ist.
            ' Dann wird die Endposition auf die Laenge des Eingabestrings gesetzt.
            ' (Korrekter waere: wenn nach dem
            '
            If (pos_akt_xml_start_zeichen < len_eingabe_str) Then
            
                lese_position_ende = letztes_end_tag_ende
                
            Else
            
                lese_position_ende = len_eingabe_str
                
            End If

        End If
        
        '
        ' Pruefung: Soll getrennt werden?
        '
        If (knz_trennen) Then
            
            '
            ' Pruefung: Leseposition Ende gleich -1?
            '
            ' Darf nicht vorkommen, tut es aber leider doch mal
            ' ... dann muss man das eben mal grade ruecken.
            '
            If (lese_position_ende = -1) Then
                
                lese_position_ende = lese_position_start
            
            End If
            
            '
            ' Teilstring aus Eingabe
            ' Ausgehend von der aktuellen Start-Lese-Position wird bis zur End-Lese-Position
            ' der aktuelle Teilstring aus der Eingabe gelesen.
            '
            ' Die Endleseposition ist immer ein XML-End-Zeichen ">", danach kommt kein weiteres Zeichen.
            ' Ab der Startposition koennen jedoch noch Leerzeichen durch Einrueckungen vorhanden sein.
            ' Um die Leerzeichen aus der Startposition zu entfernen wird der Teilstring getrimmt.
            '
            akt_teil_string = Trim(Mid(eingabe_str, lese_position_start, (lese_position_ende - lese_position_start) + 1))
            
            '
            ' Pruefung: akt_teil_string
            ' Nur wenn der aktuell gelesene und getrimmte Teilstring Zeichen enthaelt, wird dieser
            ' dem Funktionsergebnis hinzugefuegt. Ansonsten wuerden/koennten Leerzeilen entstehen.
            '
            If (akt_teil_string <> LEER_STRING) Then
                
                str_fkt_ergebnis = str_fkt_ergebnis & str_debug & String(einrueck_zaehler, " ") & akt_teil_string & MY_CHR_13_10
            
            End If
            
            '
            ' Leseposition-Start aktualisieren
            ' Die naechste Startposition fuer den Leseprozess liegt hinter dem aktuellen Endezeichen.
            '
            lese_position_start = lese_position_ende + 1
            
            '
            ' Flag fuer das Trennen auf FALSE stellen.
            '
            knz_trennen = False
        
        End If
        
        '
        ' Einrueckzaehler erhoehen
        ' Bei der Kombination "Start-Tag" folgt "Start-Tag" wird der Einrueckzaehler
        ' um die eingestellte Breite erhoeht. Die Einrueckung wird nur gemacht, wenn
        ' in der Variablen "akt_teil_string" kein Leerstring enthalten ist. Dieses
        ' soll ein Einruecken bei fehlerhaften XML-Notationen verhindern.
        '
        ' Um ein ausuferndes Einruecken zu verhindern, ist die Einrueckobergrenze
        ' auf 200 eingestellt.
        '
        If ((knz_tag_art = tag_art_start) And (knz_letzte_tag_art = tag_art_start)) Then
            
            If (akt_teil_string <> LEER_STRING) Then
            
                If (einrueck_zaehler < 200) Then
    
                    einrueck_zaehler = einrueck_zaehler + einrueck_breite
    
                End If
                
            End If
        
        End If
        
        '
        ' Kennzeichen fuer die Tag-Art und die Tag-Endposition fuer dne naechsten Durchlauf merken
        '
        knz_letzte_tag_art = knz_tag_art
        
        letztes_end_tag_ende = pos_akt_xml_end_zeichen
        
        '
        ' Pruefung: Leseprozess am Ende angekommen?
        ' Bedingung: Position fuer XML-Startzeichen "<" kleiner der laenge der Eingabe
        '
        If (pos_akt_xml_start_zeichen < len_eingabe_str) Then
            
            '
            ' Position des naechsten XML-Startzeichens ermitteln.
            ' Das naechste Start-Zeichen kann nur hinter der Endposition des akutellen
            ' Endzeichens liegen.
            '
            pos_akt_xml_start_zeichen = InStr(pos_akt_xml_end_zeichen + 1, eingabe_str, start_zeichen)
            
            '
            ' Kein weiteres Startzeichen gefunden
            ' Restzeichen ab der aktuellen Leseposition-Start bis zum Ende dem Ergebnis hinzufuegen.
            '
            If (pos_akt_xml_start_zeichen = 0) Then
            
               str_fkt_ergebnis = str_fkt_ergebnis & String(einrueck_zaehler, " ") & Trim(Mid(eingabe_str, lese_position_start, len_eingabe_str)) & MY_CHR_13_10
            
            End If
            
        Else
            
            pos_akt_xml_start_zeichen = 0
        
        End If
        '
        ' Den Endlossschleifenverhinderungszaehler eins weiterzaehlen
        '
        zaehler_while_schleife = zaehler_while_schleife + 1
        
    Wend
    
    '
    ' Vermeidung von Leerzeilen
    ' Vor der Rueckgabe werden hintereinander kommende Zeilenumbrueche
    ' int einen Zeilenumbruch gewandelt.
    '
    formatXML = Replace(str_fkt_ergebnis, MY_CHR_13_10 & MY_CHR_13_10, MY_CHR_13_10) & MY_CHR_13_10
    
End Function

'################################################################################
'
Public Function startGetHexDump(pString As String, pZahlenJeZeile As Integer, pModus As Integer) As String

On Error GoTo errStartGetHexDump

    m_toggle_mr_stringer_fkt = Not m_toggle_mr_stringer_fkt

Dim akt_position            As Integer ' aktuelle Leseposition der While-Schleife
Dim akt_zeichen             As String  ' aktuelle Zeichen in der While-Schleife
Dim str_my_cr               As String  ' Das in dieser Funktion verwendete Zeilenumbruchszeichen
Dim str_fkt_ergebnis        As String  ' Ergebnisstring fuer die Rueckgabe
Dim str_ausgabe_position    As String  ' String fuer die Zeilen-Nummernausgabe
Dim str_ausgabe_ascii_werte As String  ' String fuer die Hex- bzw. Dezimalangabe der einzelnen Zeichen
Dim str_ausgabe_zeichen     As String  ' String fuer die Zeichenausgabe einer Zeile
Dim str_vorlaufende_nullen  As String  ' String mit 20 Nullen fuer die vorlaufenden 0en.
Dim zahlen_je_zeile_anzahl  As Integer ' Anzahl der Ausgabezeichen je Zeile
Dim zahlen_je_zeile_zaehler As Integer ' Aktuelle Zeichenanzahl der aktuellen Zeile
Dim knz_ausgabe_hexadezimal As Boolean ' Steuert, ob die Ausgabe als Hex- oder Dezimalangabe erstellt wird
    
    str_my_cr = vbCrLf
    
    str_vorlaufende_nullen = "00000000000000000000"
    
    str_fkt_ergebnis = LEER_STRING
    
    zahlen_je_zeile_anzahl = pZahlenJeZeile
    
    zahlen_je_zeile_zaehler = 0
    
    akt_position = 1
    
    str_ausgabe_position = Right(str_vorlaufende_nullen & akt_position, 6) & " "
    
    knz_ausgabe_hexadezimal = m_toggle_mr_stringer_fkt
    
    Dim knz_erstelle_vorlaufende_nullen As Boolean
    
    knz_erstelle_vorlaufende_nullen = pModus = 1
    
    '
    ' Die While-Schleife laeuft ueber die Laenge des Eingabestrings.
    '
    While (akt_position <= Len(pString))
    
        '
        ' Das Zeichen an der aktuellen Leseposition ermitteln
        '
        akt_zeichen = Mid(pString, akt_position, 1)
        
        '
        ' Pruefung: Hex-Version oder Dezimale-Version ?
        '
        ' Je nach Kennzeichenfeld wird eine hexadzimale Zahl oder eine
        ' dezmiale Zahl vom ASCI-Wert erstellt und dem String der Zahlen
        ' hinzugefuegt.
        '
        ' Die hexadezimalen Zahlen bekommen ein vorlaufendes Leerzeichen,
        ' damit die Breite mit der dezimalen Version identisch ist.
        '
        If (knz_ausgabe_hexadezimal) Then
            
            If (knz_erstelle_vorlaufende_nullen) Then
        
                str_ausgabe_ascii_werte = str_ausgabe_ascii_werte & " " & Right(str_vorlaufende_nullen & Hex(Asc(akt_zeichen)), 2) & " "
                
            Else
        
                str_ausgabe_ascii_werte = str_ausgabe_ascii_werte & " " & Right("  " & Hex(Asc(akt_zeichen)), 2) & " "
                
            End If
            
        Else
        
            If (knz_erstelle_vorlaufende_nullen) Then
        
                str_ausgabe_ascii_werte = str_ausgabe_ascii_werte & Right(str_vorlaufende_nullen & Asc(akt_zeichen), 3) & " "
                
            Else
        
                str_ausgabe_ascii_werte = str_ausgabe_ascii_werte & Right("   " & Asc(akt_zeichen), 3) & " "
                
            End If
            
        End If
        
        '
        ' Zeichenstring
        '
        ' Ist der Asci-Wert kleiner 31 wird fuer das Zeichen ein Punkt
        ' dem Zeichenstring hinzugefuegt.
        '
        ' Ist der Asci-Wert groesser 30 wird das Zeichen normal hinzugefuegt.
        '
        If (Asc(akt_zeichen) < 31) Then
        
            str_ausgabe_zeichen = str_ausgabe_zeichen & "."
            
        Else
        
            str_ausgabe_zeichen = str_ausgabe_zeichen & akt_zeichen
            
        End If
        
        '
        ' Zahlenzaehler erhoehen
        '
        zahlen_je_zeile_zaehler = zahlen_je_zeile_zaehler + 1
        
        '
        ' Pruefung: Zahlengrenze je Zeile erreicht ?
        '
        ' Hat der Zahlenzaehler die Grenze der Zahlenanzahl je Zeile erreicht,
        ' wird dem Ergebnisstring eine neue Zeile hinzugefuegt.
        '
        ' Die neue Zeile setzt sich aus dem Positionsstring, dem Zahlenstring und
        ' dem Zeichenstring zusammen.
        '
        If (zahlen_je_zeile_zaehler = zahlen_je_zeile_anzahl) Then
            
            str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr & str_ausgabe_position & str_ausgabe_ascii_werte & str_ausgabe_zeichen
            
            str_ausgabe_position = Right(str_vorlaufende_nullen & (akt_position + 1), 6) & LEER_ZEICHEN
            
            str_ausgabe_ascii_werte = LEER_STRING
            
            str_ausgabe_zeichen = LEER_STRING
            
            zahlen_je_zeile_zaehler = 0
        
        End If
        
        '
        ' Am Ende der While-Schleife wird die Leseposition um 1 erhoeht
        ' und mit dem naechsten Schleifendurchlauf weitergemacht.
        
        akt_position = akt_position + 1
    
    Wend

    '
    ' Ist der Zahlenzaehler nach der While-Schleife groesser als 0, sind
    ' noch Ausgabedaten vorhanden, welche dem Ergebnis noch hinzugefuegt
    ' werden muessen.
    '
    If (zahlen_je_zeile_zaehler > 0) Then
        
        '
        ' Ist die Zahlenanzahl je Zeile nicht erreicht worden, werden
        ' die noch fehlenden Position mit Leerzeichen aufgefuellt. Das
        ' sind 3 Leerzeichen fuer die Maximalausdehung der Zahlen und
        ' ein Leerzeichen fuer die Trennung der Zahlen.
        '
        ' Die Variable "str_ausgabe_zeichen" muss nicht aufgefuellt werden.
        ' Der Inhalt dieser Variablen wird am Ende des Ergebnisstrings
        ' hinzugefuegt.
        '
        While (zahlen_je_zeile_zaehler < zahlen_je_zeile_anzahl)
            
            str_ausgabe_ascii_werte = str_ausgabe_ascii_werte & "   " & LEER_ZEICHEN
            
            zahlen_je_zeile_zaehler = zahlen_je_zeile_zaehler + 1
        
        Wend
        
        str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr & str_ausgabe_position & str_ausgabe_ascii_werte & str_ausgabe_zeichen
        
    End If

EndFunktion:

    On Error Resume Next

    str_ausgabe_ascii_werte = LEER_STRING
    
    str_ausgabe_zeichen = LEER_STRING

    DoEvents

    startGetHexDump = str_fkt_ergebnis

    Exit Function

errStartGetHexDump:

    str_fkt_ergebnis = str_fkt_ergebnis & "Fehler: errStartGetHexDump: " & Err & " " & Error & " " & Erl

    Resume EndFunktion

End Function

'################################################################################
'
Public Function startGetHexJDump2(pString As String, pZahlenJeZeile As Integer, pModus As Integer) As String

On Error GoTo errStartGetHexJDump2

    m_toggle_mr_stringer_fkt = Not m_toggle_mr_stringer_fkt

Dim akt_position            As Integer ' aktuelle Leseposition der While-Schleife
Dim akt_zeichen             As String  ' aktuelle Zeichen in der While-Schleife
Dim str_my_cr               As String  ' Das in dieser Funktion verwendete Zeilenumbruchszeichen
Dim str_fkt_ergebnis        As String  ' Ergebnisstring fuer die Rueckgabe
Dim str_ausgabe_position    As String  ' String fuer die Zeilen-Nummernausgabe
Dim str_ausgabe_ascii_werte As String  ' String fuer die Hex- bzw. Dezimalangabe der einzelnen Zeichen
Dim str_ausgabe_zeichen     As String  ' String fuer die Zeichenausgabe einer Zeile
Dim str_vorlaufende_nullen  As String  ' String mit 20 Nullen fuer die vorlaufenden 0en.
Dim zahlen_je_zeile_anzahl  As Integer ' Anzahl der Ausgabezeichen je Zeile
Dim zahlen_je_zeile_zaehler As Integer ' Aktuelle Zeichenanzahl der aktuellen Zeile
Dim knz_ausgabe_hexadezimal As Boolean ' Steuert, ob die Ausgabe als Hex- oder Dezimalangabe erstellt wird
    
    str_my_cr = vbCrLf
    
    str_vorlaufende_nullen = "00000000000000000000"
    
    str_fkt_ergebnis = LEER_STRING
    
    zahlen_je_zeile_anzahl = pZahlenJeZeile
    
    zahlen_je_zeile_zaehler = 0
    
    akt_position = 1
    
    str_ausgabe_position = Right(str_vorlaufende_nullen & akt_position, 6) & LEER_ZEICHEN
    
    knz_ausgabe_hexadezimal = m_toggle_mr_stringer_fkt
    
    Dim knz_erstelle_vorlaufende_nullen As Boolean
    
    knz_erstelle_vorlaufende_nullen = pModus = 1
    
    '
    ' Die While-Schleife laeuft ueber die Laenge des Eingabestrings.
    '
    While (akt_position <= Len(pString))
    
        '
        ' Das Zeichen an der aktuellen Leseposition ermitteln
        '
        akt_zeichen = Mid(pString, akt_position, 1)
        
        '
        ' Pruefung: Hex-Version oder Dezimale-Version ?
        '
        ' Je nach Kennzeichenfeld wird eine hexadzimale Zahl oder eine
        ' dezmiale Zahl vom ASCI-Wert erstellt und dem String der Zahlen
        ' hinzugefuegt.
        '
        ' Die hexadezimalen Zahlen bekommen ein vorlaufendes Leerzeichen,
        ' damit die Breite mit der dezimalen Version identisch ist.
        '
        If (knz_ausgabe_hexadezimal) Then
            
                str_ausgabe_ascii_werte = str_ausgabe_ascii_werte & " 0x" & Right("00" & Hex(Asc(akt_zeichen)), 2) & ", "
            
        Else
        
                str_ausgabe_ascii_werte = str_ausgabe_ascii_werte & Right("   " & Asc(akt_zeichen), 3) & ", "
            
        End If
        
        '
        ' Zeichenstring
        '
        ' Ist der Asci-Wert kleiner 31 wird fuer das Zeichen ein Punkt
        ' dem Zeichenstring hinzugefuegt.
        '
        ' Ist der Asci-Wert groesser 30 wird das Zeichen normal hinzugefuegt.
        '
        If (Asc(akt_zeichen) < 31) Then
        
            str_ausgabe_zeichen = str_ausgabe_zeichen & "."
            
        Else
        
            str_ausgabe_zeichen = str_ausgabe_zeichen & akt_zeichen
            
        End If
        
        '
        ' Zahlenzaehler erhoehen
        '
        zahlen_je_zeile_zaehler = zahlen_je_zeile_zaehler + 1
        
        '
        ' Pruefung: Zahlengrenze je Zeile erreicht ?
        '
        ' Hat der Zahlenzaehler die Grenze der Zahlenanzahl je Zeile erreicht,
        ' wird dem Ergebnisstring eine neue Zeile hinzugefuegt.
        '
        ' Die neue Zeile setzt sich aus dem Positionsstring, dem Zahlenstring und
        ' dem Zeichenstring zusammen.
        '
        If (zahlen_je_zeile_zaehler = zahlen_je_zeile_anzahl) Then
            
            str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr & str_ausgabe_position & str_ausgabe_ascii_werte & str_ausgabe_zeichen
            
            str_ausgabe_position = Right(str_vorlaufende_nullen & (akt_position + 1), 6) & " "
            
            str_ausgabe_ascii_werte = LEER_STRING
            
            str_ausgabe_zeichen = LEER_STRING
            
            zahlen_je_zeile_zaehler = 0
        
        End If
        
        '
        ' Am Ende der While-Schleife wird die Leseposition um 1 erhoeht
        ' und mit dem naechsten Schleifendurchlauf weitergemacht.
        
        akt_position = akt_position + 1
    
    Wend

    '
    ' Ist der Zahlenzaehler nach der While-Schleife groesser als 0, sind
    ' noch Ausgabedaten vorhanden, welche dem Ergebnis noch hinzugefuegt
    ' werden muessen.
    '
    If (zahlen_je_zeile_zaehler > 0) Then
        
        '
        ' Ist die Zahlenanzahl je Zeile nicht erreicht worden, werden
        ' die noch fehlenden Position mit Leerzeichen aufgefuellt. Das
        ' sind 3 Leerzeichen fuer die Maximalausdehung der Zahlen und
        ' ein Leerzeichen fuer die Trennung der Zahlen.
        '
        ' Die Variable "str_ausgabe_zeichen" muss nicht aufgefuellt werden.
        ' Der Inhalt dieser Variablen wird am Ende des Ergebnisstrings
        ' hinzugefuegt.
        '
        While (zahlen_je_zeile_zaehler < zahlen_je_zeile_anzahl)
            
            str_ausgabe_ascii_werte = str_ausgabe_ascii_werte & "   " & " "
            
            zahlen_je_zeile_zaehler = zahlen_je_zeile_zaehler + 1
        
        Wend
        
        str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr & str_ausgabe_position & str_ausgabe_ascii_werte & str_ausgabe_zeichen
        
    End If

EndFunktion:

    On Error Resume Next

    str_ausgabe_ascii_werte = LEER_STRING
    
    str_ausgabe_zeichen = LEER_STRING

    DoEvents

    startGetHexJDump2 = str_fkt_ergebnis

    Exit Function

errStartGetHexJDump2:

    str_fkt_ergebnis = str_fkt_ergebnis & "Fehler: errStartGetHexJDump2: " & Err & " " & Error & " " & Erl

    Resume EndFunktion

End Function

'################################################################################
'
Public Function startGetAsciiPrint(pString As String) As String

'
' ? startGetAsciiPrint( "!#$%&'*+-/=?^_`{|}~" )
'
On Error GoTo errStartGetAsciiPrint

Dim akt_position           As Integer ' aktuelle Leseposition der While-Schleife
Dim akt_zeichen            As String  ' aktuelle Zeichen in der While-Schleife
Dim str_my_cr              As String  ' Das in dieser Funktion verwendete Zeilenumbruchszeichen
Dim str_fkt_ergebnis       As String  ' Ergebnisstring fuer die Rueckgabe
Dim str_vorlaufende_nullen As String  ' String mit Nullen fuer die vorlaufenden 0en.
Dim akt_ascii_wert         As Integer

    str_my_cr = vbCrLf
    
    str_vorlaufende_nullen = "0000000"
    
    str_fkt_ergebnis = LEER_STRING
    
    akt_position = 1
    
    '
    ' While-Schleife ueber die Laenge der Eingabe.
    ' Es werden 32123 Zeichen beruecksichtigt.
    '
    While ((akt_position <= Len(pString)) And (akt_position <= 32123))
    
        '
        ' Aktuelles Zeichen an der Lesepositon lesen
        '
        akt_zeichen = Mid(pString, akt_position, 1)
        
        '
        ' Ascii-Wert des Zeichens ermitteln
        '
        akt_ascii_wert = Asc(akt_zeichen)
        
        '
        ' Aufbau der Ergebniszeile
        '
        str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr & TRENN_STRING_6
        
        If (akt_ascii_wert < 31) Then
        
            '
            ' Ist der Ascii-Wert kleiner als 31 wird das Zeichen selber nicht in Teil im Ausgabestring.
            '
            str_fkt_ergebnis = str_fkt_ergebnis & " asc(   ) = "
            
        Else
        
            '
            ' Alle Ascii-Werte groesser als 31 werden zusaetzlich als Zeichen im Ergebnisstring ausgegeben
            '
        
            str_fkt_ergebnis = str_fkt_ergebnis & " asc(""" & akt_zeichen & """) = "
            
        End If
        
        str_fkt_ergebnis = str_fkt_ergebnis & Right(str_vorlaufende_nullen & akt_ascii_wert, 3) & " "
        
        '
        ' Leseposition um eine Position erhoehen und mit dem naechsten Zeichen weitermachen.
        '
        akt_position = akt_position + 1
    
    Wend

EndFunktion:

    On Error Resume Next

    DoEvents

    startGetAsciiPrint = str_fkt_ergebnis

    Exit Function

errStartGetAsciiPrint:

    str_fkt_ergebnis = str_fkt_ergebnis & "Fehler: errStartGetAsciiPrint: " & Err & " " & Error & " " & Erl

    Resume EndFunktion

End Function

'################################################################################
'
Public Function startGrepSuchWorte(pSuchWorte As String, pString As String, pKnzArt As Integer) As String

On Error GoTo errStartGrepSuchWorte

Dim cls_string_array   As clsStringArray
Dim str_markierung     As String
Dim aktuelle_zeile     As String
Dim akt_zaehler        As Long
Dim akt_anzahl_zeilen  As Long

    startGrepSuchWorte = LEER_STRING
    
    '
    ' Pruefung: Parameter "pSuchworte" ungleich Leerstring ?
    '
    If (Trim(pSuchWorte) <> LEER_STRING) Then
    
        '
        ' Es werden alle Suchworte aus dem Parameter "pSuchworte" in dem Text "pString" gesucht.
        ' In der ersten Schleife werden alle Vorkommen aller Suchworte mit einem Praefix versehen.
        ' Das jetzt vorangestellte Praefix-Suchzeichen wird von der zweiten Schleife
        ' gesucht. Das Ergebnis sind alle Zeilen in welchem die Markierung vorhanden ist
        '
        
        str_markierung = "#M#A#R#K#S#U#C#H#W#O#R#T#"
        
        Set cls_string_array = startMultiline(pSuchWorte)
        
        If ((cls_string_array Is Nothing) = False) Then
            
            '
            ' Schleife 1:
            ' Stringarray = Suchworte
            ' Ersetzt wird global in "pString".
            '
            akt_anzahl_zeilen = cls_string_array.getAnzahlStrings
            
            akt_zaehler = 0
            
            While (akt_zaehler <= akt_anzahl_zeilen)
            
                aktuelle_zeile = cls_string_array.getString(akt_zaehler)
            
                If (aktuelle_zeile <> LEER_STRING) Then
                
                    pString = Replace(pString, aktuelle_zeile, str_markierung & aktuelle_zeile)
                    
                End If

                akt_zaehler = akt_zaehler + 1
                
            Wend
            
            Set cls_string_array = Nothing
            
            '
            ' Schleife 2:
            ' Stringarray = pString mit den markierten Suchwoertern der ersten Schleife
            ' In jeder Zeile wird nach der Markierung gesucht.
            '
            ' Zeilen fuer das Ergebnis bleiben im Stringarray erhalten.
            ' Nicht Ergebniszeilen werden im Stringarray geloescht, bzw.
            ' auf einen Leerstring gestellt.
            '
            
            Set cls_string_array = startMultiline(pString)
            
            If ((cls_string_array Is Nothing) = False) Then
                
                akt_anzahl_zeilen = cls_string_array.getAnzahlStrings
                
                akt_zaehler = 0
                
                While (akt_zaehler <= akt_anzahl_zeilen)
                
                    aktuelle_zeile = cls_string_array.getString(akt_zaehler)
                
                    If (aktuelle_zeile <> LEER_STRING) Then
                        
                        If (pKnzArt = 1) Then ' 1 = Positiv Grep +
                        
                            If (InStr(aktuelle_zeile, str_markierung) = 0) Then
                            
                                Call cls_string_array.setString(akt_zaehler, LEER_STRING)
                            
                            End If
                            
                        Else ' 0 oder andere Zahl = Negativ Grep -
                            
                            If (InStr(aktuelle_zeile, str_markierung) > 0) Then
                            
                                '
                                ' Alle Zeilen, in denen ein Suchwort gefunden wurde, werden ausgenullt
                                ' Es bleiben am Ende nur diejenigen Zeilen uebrig, bei welchem kein
                                ' Suchwort vorkam.
                                '
                                Call cls_string_array.setString(akt_zaehler, LEER_STRING)
                            
                            End If
                            
                        End If
                        
                    End If
    
                    akt_zaehler = akt_zaehler + 1
                    
                Wend
                
                startGrepSuchWorte = Replace(cls_string_array.toString(MY_CHR_13_10, True), str_markierung, LEER_STRING)
                
                Set cls_string_array = Nothing
            
            End If
        
        End If
    
    End If
    
EndFunktion:
    
    Set cls_string_array = Nothing
    
    DoEvents
    
    Exit Function
    
errStartGrepSuchWorte:

    Resume EndFunktion
    
End Function

'################################################################################
'
Public Function startCsvKonstanten(pString As String, pTrennzeichen As String) As String

On Error GoTo errStartCsvKonstanten

Dim cls_string_array_input        As clsStringArray
Dim cls_string_array_ergebnis     As clsStringArray
Dim zeilen_zaehler                As Long
Dim zeilen_anzahl                 As Long
Dim akt_zeile                     As String
Dim akt_konstanten_name           As String
Dim akt_konstanten_wert           As String
Dim akt_konstanten_deklaration    As String
Dim temp_string_bis_trennzeichen  As String
Dim temp_string_nach_trennzeichen As String
Dim position_start_suche          As Integer
Dim position_trennzeichen         As Integer
    
    startCsvKonstanten = LEER_STRING
    '
    ' Die Eingabe wird in einem String-Array gespeichert.
    '
    Set cls_string_array_input = startMultiline(pString)
    
    '
    ' Pruefung: "cls_string_array_input" gleich Nothing ?
    '
    ' Die Konstantenerstellung wird gestartet, wenn die Funktion "startMultiline"
    ' eine Instanz vom Typ "clsStringArray" zurueckgeliefert hat.
    '
    ' Ist die Eingabe ein Leerstring, ist die Ausgabe ein Leerstring.
    '
    If ((cls_string_array_input Is Nothing) = False) Then
        
        Set cls_string_array_ergebnis = New clsStringArray
        
        zeilen_anzahl = cls_string_array_input.getAnzahlStrings
        
        zeilen_zaehler = 1
        
        While (zeilen_zaehler <= zeilen_anzahl)
            
            akt_zeile = Trim(cls_string_array_input.getString(zeilen_zaehler))
            
            position_trennzeichen = InStr(akt_zeile, pTrennzeichen)
            
            If (position_trennzeichen > 0) Then
                
                position_start_suche = position_trennzeichen + Len(pTrennzeichen)
                
                temp_string_bis_trennzeichen = Left(akt_zeile, position_trennzeichen - 1)
                
                temp_string_nach_trennzeichen = Mid(akt_zeile, position_start_suche, Len(akt_zeile))
                
                akt_konstanten_name = "CONST_" & UCase(getKlartext(temp_string_bis_trennzeichen, UNTER_STRICH))
                
                akt_konstanten_wert = temp_string_nach_trennzeichen
                
                akt_konstanten_deklaration = "Const " & akt_konstanten_name & " = """ & akt_konstanten_wert & """ "
                
                Call cls_string_array_ergebnis.addString(akt_konstanten_deklaration)
            
            End If
            
            zeilen_zaehler = zeilen_zaehler + 1
        
        Wend
        
        Call cls_string_array_ergebnis.addString("")
        Call cls_string_array_ergebnis.addString("")
        
        zeilen_zaehler = 1
        
        While (zeilen_zaehler <= zeilen_anzahl)
            
            akt_zeile = Trim(cls_string_array_input.getString(zeilen_zaehler))
            
            position_trennzeichen = InStr(akt_zeile, pTrennzeichen)
            
            If (position_trennzeichen > 0) Then
                
                position_start_suche = position_trennzeichen + Len(pTrennzeichen)
                
                temp_string_bis_trennzeichen = Left(akt_zeile, position_trennzeichen - 1)
                
                temp_string_nach_trennzeichen = Mid(akt_zeile, position_start_suche, Len(akt_zeile))
                
                akt_konstanten_name = "STR_" & UCase(getKlartext(temp_string_bis_trennzeichen, UNTER_STRICH))
                
                akt_konstanten_wert = temp_string_nach_trennzeichen
                
                akt_konstanten_deklaration = "public static final String " & akt_konstanten_name & " = """ & akt_konstanten_wert & """; "
                
                Call cls_string_array_ergebnis.addString("/** Konstante fuer " & getKlartext(temp_string_bis_trennzeichen, " ") & " """ & akt_konstanten_wert & """ */")
                
                Call cls_string_array_ergebnis.addString(akt_konstanten_deklaration)
            
            End If
            
            zeilen_zaehler = zeilen_zaehler + 1
        
        Wend
        
        Call cls_string_array_ergebnis.addString("")
        Call cls_string_array_ergebnis.addString("")
        
        startCsvKonstanten = cls_string_array_ergebnis.toString(MY_CHR_13_10)

    End If
    
EndFunktion:
    
    Set cls_string_array_input = Nothing
    
    DoEvents
    
    Exit Function
    
errStartCsvKonstanten:
    
    Resume EndFunktion
    
End Function

'################################################################################
'
Public Function startErstelleKonstantenEinfach(pString As String, pFunktion As Integer, pSelStart As Long, pSelLength As Long) As String

On Error GoTo errStartKonstantenEinfach

Dim cls_string_array              As clsStringArray
Dim akt_string                    As String
Dim aktuelle_zeile                As String
Dim ersatz_string_markierung_1    As String
Dim ersatz_string_markierung_2    As String
Dim ersatz_string_markierung_3    As String
Dim ersatz_string_markierung_4    As String
Dim inhalt_spalte_1               As String
Dim inhalt_spalte_2               As String
Dim knz_benutze_markierung        As Boolean
Dim pos_letztes_chr_vor_sel_start As Long
Dim pos_trennzeichen              As String
Dim temp_trennzeichen             As String
Dim zeilen_anzahl                 As Long
Dim zeilen_zaehler                As Long
Dim ab_position                   As Long
Dim bis_position                  As Long

    temp_trennzeichen = "##1##2KO"
    
    '
    ' Stringarray aus pString erstellen
    '
    Set cls_string_array = startMultiline(pString)
    
    '
    ' Pruefung: String-Array Instanz vorhanden?
    '
    If (cls_string_array Is Nothing) Then
    
        '
        ' Ist keine String-Array-Instanz vorhanden, ist das Funktionsergebnis ein Leerstring
        '
        startErstelleKonstantenEinfach = LEER_STRING
        
    Else

        '
        '
        '
        pos_letztes_chr_vor_sel_start = getLetztePositionVorPos(pString, getBenutztesChr13(pString), pSelStart)
        
        If (pos_letztes_chr_vor_sel_start > 0) Then
            
            ab_position = (pSelStart - pos_letztes_chr_vor_sel_start)
        
        Else
            
            ab_position = pSelStart + 1
        
        End If
        
        bis_position = (ab_position + pSelLength) - 1
        
        knz_benutze_markierung = (ab_position >= 0) And (bis_position >= ab_position)
        
        zeilen_anzahl = cls_string_array.getAnzahlStrings
        
        zeilen_zaehler = 1
        '
        ' While-Schleife ueber alle Zeilen des String-Arrays
        '
        While (zeilen_zaehler <= zeilen_anzahl)
        
            '
            ' Aktuelle Zeile holen
            '
            aktuelle_zeile = cls_string_array.getString(zeilen_zaehler)
            
            '
            ' Pruefung: Aktuelle Zeile gesetzt?
            '
            If (aktuelle_zeile <> LEER_STRING) Then
            
                '
                ' Suche das Trennzeichen fuer Variablenname und Variableninhalt
                '
                pos_trennzeichen = InStr(aktuelle_zeile, "=")
                
                '
                ' Preufung: Trennzeichen gefunden?
                '
                If (pos_trennzeichen > 1) Then
                
                    '
                    ' Wurde ein Trennzeichen gefunden, wird
                    ' in inhalt_spalte_1 der Variablenname gespeichert
                    ' in inhalt_spalte_2 der Variablenwert gespeichert
                    '
                    inhalt_spalte_1 = Left(aktuelle_zeile, pos_trennzeichen - 1)
                    
                    inhalt_spalte_2 = Right(aktuelle_zeile, Len(aktuelle_zeile) - pos_trennzeichen)
                
                Else

                    '
                    ' Wurde kein Trennzeichen gefunden, sind beide Strings
                    ' gleich der aktuellen Zeile.
                    '
                    inhalt_spalte_1 = aktuelle_zeile
                    
                    inhalt_spalte_2 = aktuelle_zeile
                
                End If
                
                If (knz_benutze_markierung) Then
                
                    inhalt_spalte_1 = getStringAbBis(aktuelle_zeile, ab_position, bis_position)
                
                End If
                
                '
                ' Erstellung Konstanten-Name
                ' Alle Leerzeichen werden mit einem Unterstrich vertauscht.
                ' Alle Buchstaben als Grossbuchstaben.
                '
                inhalt_spalte_1 = UCase(getKlartext(inhalt_spalte_1, UNTER_STRICH))
                
                '
                ' Erstellung Konstanten-Wert
                ' Der Konstantenwert darf selber keine Anfuehrungszeichen enthalten.
                ' Es werden alle Anfuehrungszeichen entfernt und das Ergebnis getrimmt.
                '
                inhalt_spalte_2 = Trim(Replace(inhalt_spalte_2, """", ""))
                
                Call cls_string_array.setString(zeilen_zaehler, MARKIER_STRING_INTERN_1 & inhalt_spalte_1 & MARKIER_STRING_INTERN_2 & inhalt_spalte_2 & MARKIER_STRING_INTERN_3 & inhalt_spalte_1 & MARKIER_STRING_INTERN_4)
            
            End If
            
            zeilen_zaehler = zeilen_zaehler + 1
        
        Wend

        If (pFunktion = 1) Then
        
            ersatz_string_markierung_1 = "public static final String "
            ersatz_string_markierung_2 = " " & AUSRICHT_STRING_TEMP_1 & "= """
            ersatz_string_markierung_3 = """; // "" + "
            ersatz_string_markierung_4 = " + """

        ElseIf (pFunktion = 2) Then
        
            ersatz_string_markierung_1 = "public const "
            ersatz_string_markierung_2 = " " & AUSRICHT_STRING_TEMP_1 & "= """
            ersatz_string_markierung_3 = """ ' "" & "
            ersatz_string_markierung_4 = " & """
        
        ElseIf (pFunktion = 3) Then
        
            ersatz_string_markierung_1 = "prop_inst.setProperty( """
            ersatz_string_markierung_2 = """, " & AUSRICHT_STRING_TEMP_1 & " """
            ersatz_string_markierung_3 = """ );"
        
        End If
        
        akt_string = cls_string_array.toString(MY_CHR_13_10)
        
        akt_string = Replace(akt_string, MARKIER_STRING_INTERN_1, ersatz_string_markierung_1)
        
        akt_string = Replace(akt_string, MARKIER_STRING_INTERN_2, ersatz_string_markierung_2)
        
        akt_string = Replace(akt_string, MARKIER_STRING_INTERN_3, ersatz_string_markierung_3)
        
        akt_string = Replace(akt_string, MARKIER_STRING_INTERN_4, ersatz_string_markierung_4)
          
        startErstelleKonstantenEinfach = akt_string

    End If
    
EndFunktion:
    
    Set cls_string_array = Nothing
    
    DoEvents
    
    Exit Function
    
errStartKonstantenEinfach:
    
    Resume EndFunktion
    
End Function

'################################################################################
'
Public Function startMultiline(pString As String) As clsStringArray

    If (pString = LEER_STRING) Then

        Set startMultiline = Nothing

        Exit Function

    End If

Dim cls_string_array      As clsStringArray
Dim knz_umbruch_vorhanden As Boolean
Dim zeichen_zeilenumbruch As String
Dim akt_zeile             As String
Dim akt_position          As Long
Dim letzte_position       As Long
Dim zeilen_zaehler        As Long

    Set cls_string_array = New clsStringArray

    '
    ' Ermittlung welches Zeilenumbruchzeichen verwendet in der Eingabe verwendet wird
    '
    zeichen_zeilenumbruch = MY_CHR_13_10

    knz_umbruch_vorhanden = (InStr(1, pString, zeichen_zeilenumbruch, vbBinaryCompare) > 0)

    If (knz_umbruch_vorhanden = False) Then

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

            Call cls_string_array.addString(akt_zeile)

            letzte_position = akt_position + Len(zeichen_zeilenumbruch)

            akt_position = InStr(letzte_position, pString, zeichen_zeilenumbruch, vbBinaryCompare)

            zeilen_zaehler = zeilen_zaehler + 1
            
        Wend

        If (letzte_position <= Len(pString)) Then

            akt_zeile = Mid(pString, letzte_position, (Len(pString) - letzte_position) + 1)

            Call cls_string_array.addString(akt_zeile)

        End If

    Else

        Call cls_string_array.addString(pString)

    End If

    Set startMultiline = cls_string_array

End Function

'################################################################################
'
' FkString.trimX( "    A  B    C  " ) = "A B C"
' FkString.trimX( "    A  B    C"   ) = "A B C"
' FkString.trimX( "ABC"             ) = "ABC"
' FkString.trimX( "      "          ) = ""
' FkString.trimX( ""                ) = ""
' FkString.trimX( null              ) = null
'
' @param pString die Eingabezeichenfolge
' @return einen String ohne vorlaufende und nachlaufende Leerzeichen und keinen doppelten Leerzeichen zwischen den Zeichen, null wenn die Eingabe selbst null ist.
'
Public Function trimX(pString As String) As String

Dim str_fkt_ergebnis As String
Dim letztes_zeichen  As String
Dim akt_zeichen      As String
Dim akt_position     As Long

    trimX = LEER_STRING
    
    If (pString <> LEER_STRING) Then
        
        '
        ' Variable "letztes_zeichen"
        ' Speichert das zuletzt hinzugefuegt Zeichen im Ergebnis. Der Startwert
        ' eines Leerzeichens sorgt dafuer, dass die fuehrenden Leerzeichen entfernt
        ' werden.
        '
        letztes_zeichen = " "
        '
        ' Variable "akt_zeichen"
        ' Speichert das aktuelle Zeichen aus der Eingabezeichenfolge. Der Startwert
        ' ist nur wegen der Initialisierung der Variable vorhanden.
        '
        akt_zeichen = "!"
        
        ' Schleife Zeichenpruefung
        ' Ueber eine For-Schleife wird jedes Zeichen der Eingabe geprueft.
        '
        For akt_position = 1 To Len(pString)
        
            akt_zeichen = Mid(pString, akt_position, 1)
        
            '
            ' Pruefung: aktuelles Zeichen ist Leerzeichen
            '
            If (akt_zeichen = " ") Then
            
                ' Leerzeichen gefunden
                ' Ist das aktuelle Zeichen ein Leerzeichen darf dieses nur dann dem
                ' Ergebnis hinzugefuegt werden, wenn das zuletzt hinzugefuegte Zeichen
                ' ungleich einem Leerzeichen war.
                '
                ' War das letzte Zeichen ein Leerzeichen, wird das neue Leerzeichen ueberlesen.
                '
                If (letztes_zeichen <> " ") Then
                
                    str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
                    
                    letztes_zeichen = akt_zeichen
                 
                End If
            
            Else
            
                '
                ' Zeichen ungleich Leerzeichen
                ' Alle anderen Zeichen werden dem Ergebnisstring hinzugefuegt. Das aktuelle
                ' Zeichen wird in der Variablen "letztes_zeichen" gespeichert.
                '
                str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
                
                letztes_zeichen = akt_zeichen
            
            End If
        
        Next
        
        '
        ' Abschlusspruefung
        ' Pruefung, ob das Ergebnis auf ein Leerzeichen endet.
        ' Das Ergebnis endet auf ein Leerzeichen, wenn das letzte hinzugefuegte Zeichen ein Leerzeichen war.
        '
        If (letztes_zeichen = LEER_ZEICHEN) Then
        
            If (Len(str_fkt_ergebnis) < 2) Then
          
                str_fkt_ergebnis = LEER_STRING
            
            Else
            
                str_fkt_ergebnis = Left(str_fkt_ergebnis, Len(str_fkt_ergebnis) - 1)
            
            End If
          
        End If
        
        trimX = str_fkt_ergebnis
        
    End If
    
End Function

'################################################################################
'
' Zieht einen Camel-Case-Text auseinander und setzt das Trennzeichen vor den GBuchstaben
'
' WICHTIG: diese Funktion funktioniert unter Lotus-Script nicht
'
' ? getKlartext( "EinCamelCaseText"            , "_" ) = Ein_Camel_Case_Text
' ? getKlartext( "EinWeitererTextMitZahlen123" , " " ) = Ein Weiterer Text Mit Zahlen 123
' ? getKlartext( "Eins2Drei4Fuenf67AchtNeun10" , " " ) = Eins 2 Drei 4 Fuenf 67 Acht Neun 10
' ? getKlartext( "MEIN_XML_TAG"                , " " ) = Mein Xml Tag
'
' PARAMETER: pString        = die Ausgangszeichenfolge
'
' RETURN : einen CamelCase-String
'
Public Function getKlartext(pString As String, pTrennzeichen As String, Optional pErhalteZeichen As String = "") As String
    
    '
    ' Parameterpruefung
    ' pString ==> wenn null oder keine Zeichen vorhanden sind, wird null zurueckgegeben
    '
    If (Len(Trim(pString)) = 0) Then

        getKlartext = LEER_STRING

        Exit Function

    End If

Dim str_fkt_ergebnis                       As String
Dim akt_zeichen                            As String
Dim letztes_zeichen                        As String
Dim str_trennzeichen                       As String
Dim knz_letztes_zeichen_war_grossbuchstabe As Boolean
Dim knz_forciere_kleinbuchstabe            As Boolean
Dim knz_trennzeichen_einfuegen             As Boolean
Dim knz_trennzeichen_erlaubt               As Boolean
Dim knz_next_zeichen_gross                 As Boolean
Dim knz_hinzfuegen                         As Boolean
Dim zaehler_schleife                       As Integer

    akt_zeichen = " "
    letztes_zeichen = " "
    str_trennzeichen = pTrennzeichen
    zaehler_schleife = 1
    knz_hinzfuegen = False
    knz_letztes_zeichen_war_grossbuchstabe = False
    knz_forciere_kleinbuchstabe = False
    knz_trennzeichen_einfuegen = False
    knz_trennzeichen_erlaubt = False
    knz_next_zeichen_gross = True

    '
    ' Schleife ueber die gesamte Laenge des Eingabestrings
    '
    While (zaehler_schleife <= Len(pString))

        '
        ' Aktuelles Zeichen aus der Eingabe bestimmen (Zaehlerposition)
        '
        akt_zeichen = Mid(pString, zaehler_schleife, 1)

        '
        ' Kleinbuchstaben
        ' Flag zum hinzufuegen setzen und forcierung von Kleinbuchstaben aufheben
        '
        If ((akt_zeichen >= "a" And akt_zeichen <= "z")) Then
        'If ( ( akt_zeichen_ascii >= 97 And akt_zeichen_ascii <= 122 ) ) Then

            knz_hinzfuegen = True

            knz_forciere_kleinbuchstabe = False

        '
        ' Grossbuchstaben
        ' Flag zum hinzufuegen und Flag fuer Grossbuchstabe setzen
        '
        ' War das letzte hinzugefuegte Zeichen selber schon ein Grossbuchstabe,
        ' wird das Flag fuer das Forcieren von Kleinbuchstaben gesetzt.
        '
        ElseIf (akt_zeichen >= "A" And akt_zeichen <= "Z") Then
        'ElseIf ( akt_zeichen_ascii >= 65 And akt_zeichen_ascii <= 90 ) Then

            knz_hinzfuegen = True

            knz_next_zeichen_gross = True

            If (knz_letztes_zeichen_war_grossbuchstabe) Then

                knz_forciere_kleinbuchstabe = True

            End If

        '
        ' Zahlen
        ' Nur Flag zum hinzufuegen setzen
        '
        ElseIf ((akt_zeichen >= "0" And akt_zeichen <= "9")) Then
        'ElseIf ( ( akt_zeichen_ascii >= 48 And akt_zeichen_ascii <= 57 ) ) Then

            knz_hinzfuegen = True

            '
            ' Trennzeichen vor Zahlen
            ' Vor einer Zahl wird nur dann ein Trennzeichen eingefuegt, wenn das letzte
            ' Zeichen KEINE Zahl war. Somit wird verhindert, dass vor jeder Zahl ein
            ' Trennzeichen steht. Zahlen sollen zusammenbleiben, aber fuer sich getrennt
            ' stehen.
            '
            knz_trennzeichen_einfuegen = Not ((letztes_zeichen >= "0") And (letztes_zeichen <= "9"))
            ' knz_trennzeichen_einfuegen = Not ( ( letztes_zeichen_ascii >= 48 ) And ( letztes_zeichen_ascii <= 57 ) )

        ElseIf (InStr(pErhalteZeichen, akt_zeichen) > 0) Then

            knz_hinzfuegen = True
        '
        ' Sonstige Zeichen
        ' Bei allen sonstigen Zeichen wird das Hinzufuegen-Flag auf False gesetzt.
        '
        Else

            knz_hinzfuegen = False

            '
            ' Definierte Sonderzeichen
            ' Bei diesen Sonderzeichen wird der nachfolgend aufgenommene Buchstabe ein Grossbuchstabe.
            ' Damit dieses klappt muss erstens die Variable "knz_next_zeichen_gross" auf True gestellt werden.
            ' Die Forcierung von Kleinbuchstaben wird aufgehoben (= Flag auf FALSE gesetzt).
            '
            If ((akt_zeichen = UNTER_STRICH) Or (akt_zeichen = " ") Or (akt_zeichen = "-") Or (akt_zeichen = "( ") Or (akt_zeichen = " )")) Then

                knz_next_zeichen_gross = True

                knz_forciere_kleinbuchstabe = False
                
                '
                ' Einzelzeichen
                ' Folgt nach einem Einzelzeichen ein Worttrenner, wird hier ein Trennzeichen hinzugefuegt.
                '
                ' Die Flagvariable "knz_letztes_zeichen_war_grossbuchstabe" wird auf FALSE gestellt, damit
                ' dass naechste Zeichen wieder zu einem Grossbuchstaben wird.
                '
                ' Die Flagvariable "knz_trennzeichen_erlaubt" wird auf FALSE gestellt, weil ansonsten mit
                ' dem naechsten Zeichen abermals ein Trennzeichen eingefuegt werden wuerde.
                '
                If ((knz_letztes_zeichen_war_grossbuchstabe) And ((akt_zeichen = UNTER_STRICH) Or (akt_zeichen = " ") Or (akt_zeichen = "-"))) Then
                
                     str_fkt_ergebnis = str_fkt_ergebnis & str_trennzeichen
                     
                     knz_letztes_zeichen_war_grossbuchstabe = False
                     
                     knz_trennzeichen_erlaubt = False
                
                End If

'        ElseIf (akt_zeichen = "‰") Then
'
'            akt_zeichen = "ae"
'
'        ElseIf (akt_zeichen = "ƒ") Then
'
'            akt_zeichen = "Ae"
'
'        ElseIf (akt_zeichen = "ˆ") Then
'
'            akt_zeichen = "oe"
'
'        ElseIf (akt_zeichen = "÷") Then
'
'            akt_zeichen = "Oe"
'
'        ElseIf (akt_zeichen = "¸") Then
'
'            akt_zeichen = "ue"
'
'        ElseIf (akt_zeichen = "‹") Then
'
'            akt_zeichen = "Ue"
'
'        ElseIf (akt_zeichen = "ﬂ") Then
'
'            akt_zeichen = "ss"
'
'        ElseIf (akt_zeichen = "È") Then
'
'            akt_zeichen = "e"
'
'        ElseIf (akt_zeichen = "Ë") Then
'
'            akt_zeichen = "e"
'
'        ElseIf (akt_zeichen = "Ä") Then
'
'            akt_zeichen = "eur"
'
'        ElseIf (akt_zeichen = "$") Then
'
'            akt_zeichen = "dollar"
'
'        End If

            End If

        End If

        '
        ' Pruefung: aktuelles Zeichen hinzufuegen
        '
        If (knz_hinzfuegen) Then

            '
            ' Pruefung: Kleinbuchstaben forcieren
            ' Wenn dem so ist, wird das aktuelle Zeichen in einen Kleinbuchstaben umgewandelt.
            ' Desweiteren wird die Flagvariable "knz_next_zeichen_gross" auf FALSE gestellt.
            ' Solange Kleinbuchstaben forciert werden sollen, wird damit ein evtl. falsches
            ' Kennzeichen ausgenullt. Wuerde es nicht gemacht werden, gaebe es einen Fehler
            ' bei z.B. TTest ==> TtEst.
            '
            If (knz_forciere_kleinbuchstabe) Then

                akt_zeichen = LCase(akt_zeichen)

                knz_next_zeichen_gross = False

            '
            ' Pruefung: naechstes Zeichen als Grossbuchstabe
            '
            ElseIf (knz_next_zeichen_gross) Then

                '
                ' Weitere Bedingung fuer einen Grossbuchstaben ist, dass das zuletzt hinzugefuegte
                ' Zeichen kein Grossbuchstabe war (sonst stehen 2 Grossbuchstaben hintereinander).
                ' Ist diese Bedingung erfuellt, wird das aktuelle Zeichen in einen Grossbuchstaben
                ' umgewandelt.
                '
                If (knz_letztes_zeichen_war_grossbuchstabe = False) Then

                    akt_zeichen = UCase(akt_zeichen)

                    knz_trennzeichen_einfuegen = True

                End If

                '
                ' Flagvariable "knz_next_zeichen_gross" selber wird auf FALSE gestellt. Dieses
                ' unabhaengig davon, ob das aktuelle Zeichen in einen Grossbuchstaben gewandelt wurde.
                '
                knz_next_zeichen_gross = False

            End If

            '
            ' Aufbau Ergebnis
            ' Pruefung, ob vor dem aktuellen Zeichen ein Trennzeichen eingefuegt werden soll.
            '
            ' Bei dieser Pruefung wird zuerst das Flag fuer die Freischaltung von Trennzeichen
            ' geprueft. Dieses verhindert ein erstes Leerzeichen am Start. Das Flag wird nach
            ' dem ersten hinzugefuegten Zeichen im Ergebnisstring auf TRUE gestellt (=erlaubt)
            '
            ' Die zweite Pruefung greift auf das eigentliche Steuerflag zu. Dieses wird immer
            ' dann auf TRUE gestellt, wenn ein Grossbuchstabe oder eine Zahl hinzugefuegt
            ' werden soll.
            '
            If ((knz_trennzeichen_erlaubt) And (knz_trennzeichen_einfuegen)) Then

                str_fkt_ergebnis = str_fkt_ergebnis & str_trennzeichen

            End If

            '
            ' Aufbau Ergebnis
            ' Das Zeichen aus der Variablen "akt_zeichen" wird dem Ergebnis-String hinzugefuegt.
            '
            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen

            '
            ' Das schlussendlich hinzugefuegte Zeichen, bestimmt den Wert fuer
            ' die Variable "knz_letztes_zeichen_war_grossbuchstabe"
            '
            knz_letztes_zeichen_war_grossbuchstabe = (akt_zeichen >= "A" And akt_zeichen <= "Z")
            'knz_letztes_zeichen_war_grossbuchstabe = ( akt_zeichen_ascii >= 65 And akt_zeichen_ascii <= 90 )

            '
            ' Das aktuelle Zeichen wird fuer die weitere Verwendung in der Variablen
            ' "letztes_zeichen" gespeichert (fuer Zahlen).
            '
            letztes_zeichen = akt_zeichen

            '
            ' Flagvariable fuer das Zulassen von Trennzeichen auf TRUE stellen, da
            ' jetzt mindestens schon ein Zeichen im Ergebnis steht.
            '
            ' Die Hinzufuege-Flagvariable "knz_trennzeichen_einfuegen" wird fuer den
            ' naechsten Schleifendurchlauf auf FALSE gestellt.
            '
            knz_trennzeichen_erlaubt = True

            knz_trennzeichen_einfuegen = False

        End If

        '
        ' Zaehler um eine Position weiterstellen und mit dem naechsten Zeichen
        ' weiter machen.
        '
        zaehler_schleife = zaehler_schleife + 1
    Wend

    getKlartext = str_fkt_ergebnis

End Function

'################################################################################
'
Public Function getStringLit(pString As String) As String

Dim str_fkt_ergebnis As String
Dim akt_zeichen      As String
Dim akt_position     As Long
Dim knz_in_string    As Boolean

    getStringLit = LEER_STRING
    
    If (pString <> LEER_STRING) Then
        
        knz_in_string = False

        akt_position = 1
        
        '
        ' Schleife Zeichenpruefung
        ' Ueber eine For-Schleife wird jedes Zeichen der Eingabe geprueft.
        '
        While (akt_position <= Len(pString))
        
            akt_zeichen = Mid(pString, akt_position, 1)
                
            If ((knz_in_string) And (akt_zeichen = """")) Then
            
                '
                ' Anfuehrungszeichen gefunden
                '
                ' naechstes Zeichen betrachen
                '
                If (akt_position + 1 < Len(pString)) Then    ' Stringlaenge
                
                    '
                    ' Pruefung: ist das naechste Zeichen ein Anfuehrungszeichen
                    '
                    If (Mid(pString, akt_position + 1, 1) = """") Then
                    
                        '
                        ' JA = Maskiertes Anfuehrungszeichen gefunden beide Zeichen in String uebernehmen
                        '
                        
                        ' akt_position     = Anfuehrungszeichen == geprueft
                        ' akt_position + 1 = Anfuehrungszeichen == geprueft
                        '
                        ' ... Leseposition auf "akt_pos +1" stellen
                        '
                        akt_position = akt_position + 1
                        '
                        ' ... das aktuelle Zeichen mit zwei Anfuehrungszeichen bestuecken
                        '
                        akt_zeichen = """"""
                    End If
                
                End If
                
            End If
                
            If (akt_zeichen = """") Then
            
                If (knz_in_string = False) Then
                
                    str_fkt_ergebnis = str_fkt_ergebnis & "|"
                
                Else
                    
                    'str_fkt_ergebnis = str_fkt_ergebnis & ">"
                
                End If
                
                knz_in_string = Not knz_in_string
                
                akt_zeichen = LEER_STRING
            
            End If
            
            If ((knz_in_string) And (akt_zeichen <> LEER_STRING)) Then
                
                str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
                
            End If
            
            akt_position = akt_position + 1
            
        Wend
    
        getStringLit = str_fkt_ergebnis
        
    End If
    
End Function

'################################################################################
'
Private Function getStringLitKonst(pString As String) As String

Dim str_fkt_ergebnis        As String
Dim akt_zeichen             As String
Dim akt_position            As Long
Dim knz_in_string           As Boolean
Dim akt_literal             As String
Dim str_trennzeichen        As String
Dim str_trennzeichen_to_set As String
Dim zeichen_anf_maskierung  As String
Dim letztes_zeichen         As String
    
    zeichen_anf_maskierung = """"

    getStringLitKonst = ""
    
    '
    ' Pruefung: ist pString gesetzt?
    '
    If (pString <> LEER_STRING) Then
        
        knz_in_string = False
        
        letztes_zeichen = "!"
        
        str_trennzeichen = LEER_STRING
        
        str_trennzeichen_to_set = getBenutztesChr13(pString)
        
        akt_position = 1
        
        '
        ' Schleife Zeichenpruefung
        ' Ueber eine For-Schleife wird jedes Zeichen der Eingabe geprueft.
        '
        While (akt_position <= Len(pString))
        
            akt_zeichen = Mid(pString, akt_position, 1)
            
            '
            ' Pruefung: Maskiertes Anfuehrungszeichen gefunden?
            '
            If ((knz_in_string) And (akt_zeichen = zeichen_anf_maskierung)) Then
                '
                ' Anfuehrungszeichen gefunden
                '
                ' naechstes Zeichen betrachen
                '
                If (akt_position + 1 < Len(pString)) Then    ' Stringlaenge
                
                    If (Mid(pString, akt_position + 1, 1) = """") Then ' ist das naechste Zeichen ein Anfuehrungszeichen
                    
                        '
                        ' JA = Maskiertes Anfuehrungszeichen gefunden beide Zeichen in String uebernehmen
                        '
                        ' akt_position     = Maskierungszeichen fuer das Anfuehrungszeichen
                        '
                        ' akt_position + 1 = Anfuehrungszeichen selber
                        '
                        ' ... Leseposition auf "akt_pos + 1" stellen
                        '
                        akt_position = akt_position + 1
                        
                        '
                        ' ... das aktuelle Zeichen mit der Maskierung uebernehmen
                        '
                        akt_zeichen = zeichen_anf_maskierung & """"
                        
                    End If
                    
                End If
                
            End If
            
            '
            ' Pruefung: Aktuelles Zeichen ein Anfuehrungszeichen?
            '
            If (akt_zeichen = """") Then
                 
                '
                ' Pruefung: Befindet sich der Leseprozess in einem Stringliteral?
                '
                ' Bei Nein, wird der Aufnahmestring mit einem Leerstring initialisiert
                '
                ' Bei Ja, wurde das Endzeichen fuer das Stringliteral gefunden.
                '
                If (knz_in_string = False) Then
                
                    akt_literal = LEER_STRING
                    
                Else
                    
                    str_fkt_ergebnis = str_fkt_ergebnis & str_trennzeichen & akt_literal
                    
                    str_trennzeichen = str_trennzeichen_to_set
                    
                    akt_literal = LEER_STRING
                    
                End If
                
                knz_in_string = Not knz_in_string
                
                akt_zeichen = LEER_STRING
            
            End If
            
            '
            ' Aktuelles Literal fortfuehren
            ' Aktuelle Zeichen dem Aufnahmestring hinzufuegen.
            '
            ' Das kann nur passieren, wenn
            '
            ' ... sich der Leseprozess in einem Literal befindet
            ' ... das aktuelle Zeichen gesetzt ist (und nicht ausgenullt worden ist)
            '
            If ((knz_in_string) And (akt_zeichen <> LEER_STRING)) Then
                
                akt_literal = akt_literal & akt_zeichen
                
            End If
            
            '
            ' Am Ende der While-Schleife wird der Leseprozess um eine Position weitergeschaltet
            akt_position = akt_position + 1
        Wend
    
        getStringLitKonst = str_fkt_ergebnis
        
    End If
    
End Function

'##############################################################################
'
Public Function getBlockZufall(pString As String, pIndexAb As Long, pIndexBis As Long) As String

On Error GoTo errGetBlockZufall

Dim str_fkt_ergebnis       As String  ' Speichert das Funktionsergebnis
Dim akt_zeichen_str        As String  ' das aktuelle Zeichen an der Leseposition
Dim akt_zeichen_ascii_wert As Long    ' ASCII-Wert des Zeichens an der Leseposition
Dim akt_index              As Long    ' Aktuelle Leseposition der Eingabe
Dim index_ab               As Long    ' Startindex fuer Umtauschvorgaenge
Dim index_bis              As Long    ' Endindex fuer Umtauschvorgaenge
Dim index_zufall           As Long    ' Zufallsindex fuer das neue Zeichen
Dim laenge_eingabe         As Long    ' Laenge der Eingabe

'Dim grundmenge_gross       As String ' Ein String mit den Buchstaben des ABC in Gross
'Dim grundmenge_klein       As String ' Ein String mit den Buchstaben des ABC in klein
'Dim grundmenge_zahlen      As String ' Ein String mit den Zahlen

    'grundmenge_zahlen = "0123456789"
    'grundmenge_gross = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    'grundmenge_klein = "abcdefghijklmnopqrstuvwxyzss"

    Randomize

'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB
'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB
'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB
'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB
'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB
'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB
'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB
'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB
'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB
'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB
'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB
'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB
'00000000000000000000000000000000000000000000000B
'00000000000000000000000000000000000000000000000B
'00000000000000000000000000000000000000000000000B
'00000000000000000000000000000000000000000000000B
'00000000000000000000000000000000000000000000000B
'00000000000000000000000000000000000000000000000B

    '
    ' Laenge des Eingabestringsermitteln
    '
    laenge_eingabe = Len(pString)
    
    '
    ' Pruefung: Laenge groesser 0 ?
    '
    ' Es wird nur dann eine Verarbeitung angestossen, wenn die
    ' Laenge der Eingabe groesser als 0 ist.
    '
    If (laenge_eingabe > 0) Then
        
        '
        ' Pruefung: Parameter "pIndexAb" groesser 0 ?
        '
        ' Wurde eine Ab-Position angegeben, wird die Index-Ab-Position
        ' auf den Parameterwert gestellt.
        '
        ' Ist der Parameter kleiner gleich 0, startet die Verarbeitung
        ' am bei Position 1.
        '
        If (pIndexAb > 0) Then
        
            index_ab = pIndexAb
            
        Else
        
            index_ab = 1
            
        End If
        
        '
        ' Pruefung: Parameter "pIndexBis" groesser 0 ?
        '
        ' Wurde eine Bis-Position angegeben, wird die Index-Bis-Position
        ' auf den Parameterwert gestellt.
        '
        ' Ist der Parameter kleiner gleich 0, endet die Verarbeitung
        ' am Ende der Eingabe.
        '
        If (pIndexBis > 0) Then
        
            index_bis = pIndexBis
            
        Else
        
            index_bis = laenge_eingabe

        End If
   
        '
        ' Leseprozess auf das erste Zeichen stellen
        '
        akt_index = 1
    
        '
        ' While-Schleife ueber die gesamte Stringlaenge
        '
        While (akt_index <= laenge_eingabe)
        
            '
            ' Es wird das aktuelle Zeichen an der Leseposition ermittelt
            '
            akt_zeichen_str = Mid(pString, akt_index, 1)
            
            '
            ' Pruefung: Start Verarbeitung ?
            '
            ' Die Verarbeitung wird gestartet, wenn der Leseprozessindex groesser gleich
            ' dem Start-Index aber kleiner gleich dem Endindex ist. Desweiteren muss
            ' der Endindex groesser als der Startindex sein.
            '
            ' Ist das nicht der Fall, wird das aktuelle Zeichen unbehandelt dem
            ' Ergebnisstring zugewiesen.
            '
            If (((akt_index >= index_ab) And (akt_index <= index_bis)) And (index_ab <= index_bis)) Then
                
                '
                ' Ermittlung ASCII-Wert
                '
                akt_zeichen_ascii_wert = Asc(akt_zeichen_str)

                If ((akt_zeichen_ascii_wert >= 97) And (akt_zeichen_ascii_wert <= 122)) Then

                    akt_zeichen_str = Chr(Int(26 * Rnd) + 97)

                    'index_zufall = Int(26 * Rnd) + 1
                    'akt_zeichen_str = Mid(grundmenge_klein, index_zufall, 1)
                
                ElseIf ((akt_zeichen_ascii_wert >= 65) And (akt_zeichen_ascii_wert <= 90)) Then
                
                    akt_zeichen_str = Chr(Int(26 * Rnd) + 65)

                    'index_zufall = Int(26 * Rnd) + 1
                    'akt_zeichen_str = Mid(grundmenge_gross, index_zufall, 1)

                ElseIf ((akt_zeichen_ascii_wert >= 48) And (akt_zeichen_ascii_wert <= 57)) Then

                    akt_zeichen_str = Chr(Int(8 * Rnd) + 48)
                    
                    'index_zufall = Int(8 * Rnd) + 1
                    'akt_zeichen_str = Mid(grundmenge_zahlen, index_zufall, 1)

                End If
            
            End If
        
            '
            ' Das aktuelle Zeichen wird dem Ergebnisstring zugewiesen
            '
            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen_str
        
            akt_index = akt_index + 1
        
        Wend
    
    End If

EndFunktion:

    On Error Resume Next

    DoEvents

    getBlockZufall = str_fkt_ergebnis

    Exit Function

errGetBlockZufall:

    Debug.Print ("Fehler: errGetBlockZufall: " & Err & " " & Error & " " & Erl)

    Resume EndFunktion

End Function

'################################################################################
'
Public Function getSwitchCase(pString As String, pIndexAb As Long, pIndexBis As Long) As String

On Error GoTo errGetSwitchCase

Dim str_fkt_ergebnis       As String  ' Speichert das Funktionsergebnis
Dim akt_zeichen_str        As String  ' das aktuelle Zeichen an der Leseposition
Dim akt_zeichen_ascii_wert As Long    ' ASCII-Wert des Zeichens an der Leseposition
Dim akt_index              As Long    ' Aktuelle Leseposition der Eingabe
Dim index_ab               As Long    ' Startindex fuer Umtauschvorgaenge
Dim index_bis              As Long    ' Endindex fuer Umtauschvorgaenge
Dim laenge_eingabe         As Long    ' Laenge der Eingabe

    '
    ' Laenge des Eingabestringsermitteln
    '
    laenge_eingabe = Len(pString)
    
    '
    ' Pruefung: Laenge groesser 0 ?
    '
    ' Es wird nur dann eine Verarbeitung angestossen, wenn die
    ' Laenge der Eingabe groesser als 0 ist.
    '
    If (laenge_eingabe > 0) Then
        
        '
        ' Pruefung: Parameter "pIndexAb" groesser 0 ?
        '
        ' Wurde eine Ab-Position angegeben, wird die Index-Ab-Position
        ' auf den Parameterwert gestellt.
        '
        ' Ist der Parameter kleiner gleich 0, startet die Verarbeitung
        ' am bei Position 1.
        '
        If (pIndexAb > 0) Then
        
            index_ab = pIndexAb
            
        Else
        
            index_ab = 1
            
        End If
        
        '
        ' Pruefung: Parameter "pIndexBis" groesser 0 ?
        '
        ' Wurde eine Bis-Position angegeben, wird die Index-Bis-Position
        ' auf den Parameterwert gestellt.
        '
        ' Ist der Parameter kleiner gleich 0, endet die Verarbeitung
        ' am Ende der Eingabe.
        '
        If (pIndexBis > 0) Then
        
            index_bis = pIndexBis
            
        Else
        
            index_bis = laenge_eingabe

        End If
   
        '
        ' Leseprozess auf das erste Zeichen stellen
        '
        akt_index = 1
    
        '
        ' While-Schleife ueber die gesamte Stringlaenge
        '
        While (akt_index <= laenge_eingabe)
        
            '
            ' Es wird das aktuelle Zeichen an der Leseposition ermittelt
            '
            akt_zeichen_str = Mid(pString, akt_index, 1)
            
            '
            ' Pruefung: Start Verarbeitung ?
            '
            ' Die Verarbeitung wird gestartet, wenn der Leseprozessindex groesser gleich
            ' dem Start-Index aber kleiner gleich dem Endindex ist. Desweiteren muss
            ' der Endindex groesser als der Startindex sein.
            '
            ' Ist das nicht der Fall, wird das aktuelle Zeichen unbehandelt dem
            ' Ergebnisstring zugewiesen.
            '
            If (((akt_index >= index_ab) And (akt_index <= index_bis)) And (index_ab <= index_bis)) Then
                
                '
                ' Ermittlung ASCII-Wert
                '
                akt_zeichen_ascii_wert = Asc(akt_zeichen_str)
                
                If ((akt_zeichen_ascii_wert >= 97) And (akt_zeichen_ascii_wert <= 122)) Then
                
                    '
                    ' Ist das aktuelle Zeichen ein Grossbuchstabe, wird dieses zu einem Kleinbuchstaben
                    '
                    akt_zeichen_str = UCase(akt_zeichen_str)
                
                ElseIf ((akt_zeichen_ascii_wert >= 65) And (akt_zeichen_ascii_wert <= 90)) Then
                
                    '
                    ' Ist das aktuelle Zeichen ein Kleinbuchstabe, wird dieses zu einem Grossbuchstaben
                    '
                    akt_zeichen_str = LCase(akt_zeichen_str)
                    
                End If
            
            End If
        
            '
            ' Das aktuelle Zeichen wird dem Ergebnisstring zugewiesen
            '
            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen_str
        
            akt_index = akt_index + 1
        
        Wend
    
    End If

EndFunktion:

    On Error Resume Next

    DoEvents

    getSwitchCase = str_fkt_ergebnis

    Exit Function

errGetSwitchCase:

    Debug.Print ("Fehler: errGetSwitchCase: " & Err & " " & Error & " " & Erl)

    Resume EndFunktion

End Function

'##############################################################################
'
' Sucht in "pEingabeString" nach Vorkommen von "pSuchString" und gibt die .
' gefundenen Worte separiert nach "pTrennzeichen" zurueck.
'
' ? fkString.getGrepSuchwort( "instanz.getWertA(), instanz.getWertB(), instanz.getWertC()", "get", ", " ) = getWertA, getWertB, getWertC
'
' PARAMETER: pEingabeString = der zu durchsuchende String
' PARAMETER: pSuchString    = die Zeichenfolge, welche gesucht werden soll
' PARAMETER: pTrennzeichen  = das zu setzende Trennzeichen bei Mehrfachfunden
'
' RETURN  : Einen String mit den gefundenen Worten, welche mit der Suchzeichenfolge beginnen.
'
Public Function getGrepSuchwort(ByVal pEingabeString As String, ByVal pSuchString As String, ByVal pTrennzeichen As String) As String

Dim akt_position_wort_start   As Long   ' aktuelle Startposition (gleichzeitig aktuelle Position des Leseprozesses)
Dim akt_position_wort_ende    As Long   ' aktuelle Endposition
Dim laenge_eingabe            As Long   ' Laenge des Eingabestrings
Dim laenge_such_string        As Long   ' Laenge des Suchstrings
Dim str_akt_trennzeichen      As String
Dim str_fkt_ergebnis          As String ' String fuer das Funktionsergebnis
Dim zeichen_wortbestandteil   As String ' Gueltige Zeichen fuer ein Wort
Dim knz_wortende_gefunden     As Integer

'
' Einzig denkbarer Fehler ist, wenn der Zaehler fuer die aktuelle Position groesser
' als der Bereich Long wird. In einem solchen Fall wird zum Funktionsende
' verzweigt, welcher dann den aktuellen Zaehler zurueck gibt.
'
On Error GoTo endGetGrepSuchwort

    zeichen_wortbestandteil = "enirstaudhgolcmfbkVvwz1paeSDA0E2RBGueMIPKF9UNW3L78oeH4T5CZJy6ssxjOueYXqQae_"

    '
    ' Aktuelles Trennzeichen ist bei der ersten Fundstelle ein Leerstring
    '
    str_akt_trennzeichen = LEER_STRING
    
    '
    ' Ergebnisstring mit einem Leerstring vorbelegen
    '
    str_fkt_ergebnis = LEER_STRING
    
    '
    ' Laenge der Eingabezeichenfolge ermitteln
    '
    laenge_eingabe = Len(pEingabeString)
    
    '
    ' Die Laenge der Such-Zeichenfolge wird der aktuellen Startposition hinzugezaehlt.
    '
    laenge_such_string = Len(pSuchString)
    
    '
    ' Bei einer Suche nach einem Leerstring wuerde es zu einer Endlosschleife kommen.
    ' Um das zu verhindern, darf die Schleife nur bei einem Suchstring mit mehr
    ' als 0 Zeichen gestartet werden.
    '
    If (laenge_such_string > 0) Then
        '
        ' Damit die Startposition fuer den ersten Aufruf der Funktion "InStr"
        ' wieder 1 ergibt, wird die Laenge des Suchstrings von 1 abgezogen.
        '
        ' (Das hinzuaddieren der Suchstringlaenge koennte auch seperat ausserhalb
        ' der Funktion "InStr" gemacht werden.)
        '
        akt_position_wort_start = 1

        Do
            '
            ' Schritt 1: Startposition ermitteln
            ' Die (naechste) Startposition fuer das Suchwort wird ab der aktuellen
            ' Leseposition im Eingabestring ermittelt.
            '
            akt_position_wort_start = InStr(akt_position_wort_start, pEingabeString, pSuchString)
        
            '
            ' Pruefung: Suchwort gefunden ?
            '
            If (akt_position_wort_start <= 0) Then
        
                '
                ' Suchwort nicht gefunden dann Suchschleife fuer Suchwort beenden
                '
                Exit Do
        
            Else
                '
                ' Wurde eine Startposition gefunden, wird ab dieser Position das Wortende gesucht.
                ' Die Wortendeposition liegt hinter der Startposition zuzueglich der Laenge des Suchwortes.
                '
                akt_position_wort_ende = akt_position_wort_start + laenge_such_string
                
                '
                ' Kennzeichen fuer "Wortende gefunden" auf 0 stellen (= noch nicht erreicht)
                '
                knz_wortende_gefunden = 0
                
                '
                ' Die Suchschleife wird solange durchlaufen wie
                ' ... die aktuelle Wortendeposition noch kleiner als die Laenge der Eingabe ist
                ' ... die Flagvariable einen weiteren Schleifendurchlauf erzwingt.
                '
                While ((akt_position_wort_ende <= laenge_eingabe) And (knz_wortende_gefunden = 0))
        
                    '
                    ' Pruefung: gueltiges Wortzeichen an Leseposition?
                    ' Solange das Zeichen an der aktuellen Leseposition noch im String der gueltigen
                    ' Zeichen gefunden werden kann, wird die Leseposition um eins erhoeht.
                    '
                    ' Ist das Zeichen nicht in den gueltigen Zeichen enthalten, wird die Flagvariable
                    ' fuer die Schleifensteuerung auf 1 gestellt. Das Wortende ist gefunden.
                    '
                    If (InStr(zeichen_wortbestandteil, Mid(pEingabeString, akt_position_wort_ende, 1)) > 0) Then
        
                        akt_position_wort_ende = akt_position_wort_ende + 1
        
                    Else
        
                        knz_wortende_gefunden = 1
        
                    End If
        
                Wend
        
                If (akt_position_wort_ende <> akt_position_wort_start) Then
                
                    akt_position_wort_ende = akt_position_wort_ende - 1
                
                End If
                
                '
                ' Ergebnisstring aufbauen
                ' Zuerst kommt das aktuelle Trennzeichen und dann das gefundene Wort ab der
                ' Startposition bis zur Endposition.
                '
                str_fkt_ergebnis = str_fkt_ergebnis & str_akt_trennzeichen & Mid(pEingabeString, akt_position_wort_start, (akt_position_wort_ende - akt_position_wort_start) + 1)
                
                '
                ' Erst nach dem ersten ermittelten Wort wird das aktuelle Trennzeichen auf
                ' das uebergebene Trennzeichen gesetzt. Dieses verhindert ein Trennzeichen
                ' vor dem ersten Wort.
                '
                str_akt_trennzeichen = pTrennzeichen
        
                '
                ' Position Start berechnen
                ' Der neue Suchprozess fuer das Suchwort beginnt ab der Position des Wortendes plus 1.
                '
                akt_position_wort_start = akt_position_wort_ende + 1
        
            End If

        Loop

    End If

endGetGrepSuchwort:
    '
    ' Funktionsende
    ' Der aufgebaute Ergebnis-String wird zurueckgegeben.
    ' Eine explizite Fehlerbehandlung wird nicht gemacht.
    '
    getGrepSuchwort = str_fkt_ergebnis

End Function

'################################################################################
'
' http://de.wikipedia.org/wiki/ROT13
'
' ROT13 (engl. rotate by 13 places, zu Deutsch in etwa "rotiere um 13 Stellen")
' ROT13 ist ein Verschiebechiffre, mit der auf einfache Weise Texte verschluesselt werden koennen.
' Dies geschieht durch Ersetzung von Buchstaben. Bei ROT13 im Speziellen wird jeder Buchstabe
' des lateinischen Alphabets durch den im Alphabet um 13 Stellen davor bzw. dahinter liegenden
' Buchstaben ersetzt.
'
' ROT13 ist nicht zur sicheren Verschluesselung gedacht. Vielmehr dient ROT13 dazu, einen Text
' unlesbar zu machen, also zu verschleiern, so dass eine Handlung des Lesers erforderlich ist,
' um den urspruenglichen Text lesen zu koennen.
'
' ROT13 selbst benutzt nur die 26 Buchstaben des lateinischen Alphabets.
'
' FkString.rot13( "Das ist ein Test" ) = "Qnf vfg rva Grfg"
' FkString.rot13( "Qnf vfg rva Grfg" ) = "Das ist ein Test"
'
' PARAMETER: pString        = der zu ver- oder entschluesselnde String
'
' RETURN  : Den ent- oder verschluesselten String der Eingabe
'
Public Function rot13(ByRef pString As String) As String

Dim akt_index        As Long
Dim str_fkt_ergebnis As String

    str_fkt_ergebnis = LEER_STRING

    For akt_index = 1 To Len(pString)

        Select Case UCase(Mid(pString, akt_index, 1))

            Case "A" To "M"
                
                str_fkt_ergebnis = str_fkt_ergebnis & Chr(Asc(Mid(pString, akt_index, 1)) + 13)

            Case "N" To "Z"
                
                str_fkt_ergebnis = str_fkt_ergebnis & Chr(Asc(Mid(pString, akt_index, 1)) - 13)

            Case Else
                
                str_fkt_ergebnis = str_fkt_ergebnis & Mid(pString, akt_index, 1)

        End Select

    Next

    rot13 = str_fkt_ergebnis

End Function

'################################################################################
'
' Ersetzt Suchfolgen durch Ersatzfolgen im Parameter pString1
' Die Such- und Ersatzfolgen werden im Parameter "pSuchErsetzWorte" zeilenweise uebergeben.
'
' Generelle Funktionsweise:
' Die Such-/Ersetzbegriffe werden zeilenweise extrahiert und auf den Parameter pString1 angewand.
'
' ? startReplaceSuchWorte("A=c" & chr(13) & "B=d", "AABB")  = ccdd
'
Public Function startReplaceSuchWorte(pSuchErsetzWorte As String, ByVal pString1 As String) As String

    startReplaceSuchWorte = LEER_STRING

    '
    ' Pruefung: pSuchErsetzWorte und pSuchString ungleich Leerstring ?
    '
    If ((pSuchErsetzWorte = LEER_STRING) Or (pString1 = LEER_STRING)) Then

        Exit Function

    End If

Dim zeichen_zeilenumbruch           As String  ' das ermittelte Trennzeichen (bzw. eben Zeilenumbruchzeichen)
Dim zeilen_zaehler                  As Long    ' Zaehler fuer Vermeidung von Endlosschleifen
Dim aktuelle_zeile                  As String  ' die aktuell gefundene Zeile aus der Eingabe
Dim aktuelle_startposition          As Long    ' die akutelle Start-Leseposition
Dim naechste_position               As Long    ' Position des naechsen gefundenen Trennzeichens
Dim knz_weiterer_schleifendurchlauf As Boolean ' Kennzeichen ob ein weiterer Schleifendurchlauf notwendig ist
Dim aktuelles_such_wort             As String  ' der aktuell zu suchende String in der Eingabe
Dim aktuelles_ersetz_wort           As String  ' der aktuelle Ersatzstring
Dim pos_trennzeichen                As Long

    '
    ' Ermittlung welches Zeilenumbruchzeichen in der Eingabe verwendet wird
    '
    zeichen_zeilenumbruch = MY_CHR_13_10

    If (InStr(1, pSuchErsetzWorte, zeichen_zeilenumbruch, vbBinaryCompare) <= 0) Then

        zeichen_zeilenumbruch = Chr(13)

    End If
    
    '
    ' Startwerte fuer die Schleifenvariablen setzen
    '
    knz_weiterer_schleifendurchlauf = True

    aktuelle_startposition = 1
    
    '
    ' Die Suchschleife laeuft solange wie...
    ' ... die Variable "knz_weiterer_schleifendurchlauf" auf TRUE steht
    ' ... der Zeilenzaehler noch unter 32200 ist
    '
    While ((knz_weiterer_schleifendurchlauf) And (zeilen_zaehler < 32220))
        
        '
        ' Naechster Zeilenumbruch
        ' Ab der aktuellen Startposition wird die naechste Position des
        ' Zeilenumbruchzeichens gesucht.
        '
        naechste_position = InStr(aktuelle_startposition, pSuchErsetzWorte, zeichen_zeilenumbruch, vbBinaryCompare)
        
        '
        ' Pruefung: Zeilenumbruchzeichen gefunden ?
        '
        If (naechste_position > 0) Then
            '
            ' Wurde eine naechste Position gefunden, wird die naechste aktuelle Zeile bestimmt.
            ' Dafuer wird der Teilstring ab der aktuellen Startposition bis zur Fundstelle des
            ' naechsten Trennzeichens bestimmt.
            '
            aktuelle_zeile = Mid(pSuchErsetzWorte, aktuelle_startposition, naechste_position - aktuelle_startposition)
            
            '
            ' Die naechste aktuelle Startposition liegt ab der Fundstelle zuzueglich der
            ' Laenge des Trennzeichens (hier = Zeilenumbruchzeichen)
            '
            aktuelle_startposition = naechste_position + Len(zeichen_zeilenumbruch)

        Else
            '
            ' Wurde kein weiteres Zeilenumbruchzeichen gefunden, ist in der
            ' Variablen "naechste_position" ein Wert von -1 enthalten.
            '
            ' Dieses ist der Hinweis, dass die While-Schleife nicht nochmal durchlaufen werden
            ' muss. Die Variable "knz_weiterer_schleifendurchlauf" wird auf FALSE gestellt.
            '
            knz_weiterer_schleifendurchlauf = False

            '
            ' Pruefung: Noch ungelesender Teilstring vorhanden ?
            '
            ' Das ist der Fall, wenn die letzte Leseposition kleiner gleich
            ' der Stringlaenge ist. Von der letzten Leseposition wird bis
            ' zum Strinende die aktuelle Zeile ermittelt.
            '
            ' Ist das nicht der Fall, wird die letzte Startposition auf -1 gestellt.
            ' Dieses ist der Hinweis, dass es keine aktuelle Zeile gibt und der
            ' Inhalt der Variablen "aktuelle_zeile" nicht verarbeitet werden darf.
            '
            If (aktuelle_startposition <= Len(pSuchErsetzWorte)) Then

                aktuelle_zeile = Mid(pSuchErsetzWorte, aktuelle_startposition, (Len(pSuchErsetzWorte) - aktuelle_startposition) + 1)

            Else

                aktuelle_startposition = -1

            End If

        End If
        
        '
        ' Pruefung: aktuelle_startposition >= 0 ?
        '
        ' Nur wenn das der Fall ist, darf der Inhalt der Variablen "aktuelle_zeile" verwendet werden.
        '
        If (aktuelle_startposition >= 0) Then
            
            pos_trennzeichen = InStr(aktuelle_zeile, "=")
            
            If (pos_trennzeichen > 0) Then
            
                '
                ' Ermittlung Such- Ersatzwort
                ' Je nach dem Wert aus "m_toggle_mr_stringer_fkt" wird einmal das Wort vor dem
                ' Gleichheitsseichen oder das nach dem Gleichheitszeichen gesucht. Das Ersatzwort
                ' wird entsprechend gesetzt.
                '
                If (m_toggle_mr_stringer_fkt) Then
                
                    aktuelles_ersetz_wort = Right(aktuelle_zeile, Len(aktuelle_zeile) - pos_trennzeichen)
                    
                    aktuelles_such_wort = Left(aktuelle_zeile, pos_trennzeichen - 1)
                    
                Else
                
                    aktuelles_such_wort = Right(aktuelle_zeile, Len(aktuelle_zeile) - pos_trennzeichen)
                    
                    aktuelles_ersetz_wort = Left(aktuelle_zeile, pos_trennzeichen - 1)
                    
                End If
                
                '
                ' Pruefung: Suchwort vorhanden?
                ' Die Suche ist nur dann Sinvoll, wenn ein Suchwort vorhanden ist. Das
                ' Ersatzwort darf ein Leerstring sein.
                '
                If (aktuelles_such_wort <> LEER_STRING) Then
                    
                    pString1 = Replace(pString1, aktuelles_such_wort, aktuelles_ersetz_wort)
                
                End If
                
            End If

        End If

        '
        ' Endlosschleifenverhinderungszaehler 1 weiterstellen
        '
        zeilen_zaehler = zeilen_zaehler + 1

    Wend

    m_toggle_mr_stringer_fkt = Not (m_toggle_mr_stringer_fkt)

    '
    ' Dem Aufrufer das Ergebnis zurueckgeben
    '
    startReplaceSuchWorte = pString1

End Function

'################################################################################
'
Public Function placeStringX(pStringA As String, pStringB As String, pFunktion As Integer, pSelStart As Long, pSelLength As Long) As String
                
    placeStringX = LEER_STRING
                
On Error GoTo errStartPlaceStringX
         
Dim ab_position             As Long
Dim aktuelle_zeile          As String
Dim cls_string_array_a      As clsStringArray
Dim cls_string_array_b      As clsStringArray
Dim str_fkt_ergebnis        As String
Dim knz_benutze_markierung  As Boolean
Dim knz_schleifen_durchlauf As Boolean
Dim temp_long_1             As Long
Dim temp_string_1           As String
Dim temp_string_2           As String
Dim temp_string_3           As String
Dim zeichen_zeilenumbruch_b As String
Dim zeichen_zeilenumbruch_a As String
Dim zeilen_anzahl_a         As Long
Dim zeilen_anzahl_b         As Long
Dim zeilen_zaehler_a        As Long
Dim zeilen_zaehler_b        As Long

    Set cls_string_array_a = startMultiline(pStringA)
    
    Set cls_string_array_b = startMultiline(pStringB)
    
    If ((cls_string_array_a Is Nothing) Or (cls_string_array_b Is Nothing)) Then
    
    ' keine Aktionen machen
    
    ElseIf (cls_string_array_a Is Nothing) Then

        placeStringX = pStringB

    ElseIf (cls_string_array_b Is Nothing) Then

        placeStringX = pStringA
    
    Else
    
        '
        ' Es wird das benutzte Zeilenumbruchszeichen aus "pStringA" ermittelt.
        ' Dieses Zeichen wird bei der Ergebnisstringerstellung verwendet.
        '
        zeichen_zeilenumbruch_a = getBenutztesChr13(pStringA)

        '
        ' Pruefung: Gibt es eine Startposition einer Markierung ?
        '
        If (pSelStart >= 0) Then

            temp_long_1 = getLetztePositionVorPos(pStringA, zeichen_zeilenumbruch_b, pSelStart)
            
            If (temp_long_1 > 0) Then
                
                ab_position = (pSelStart - temp_long_1)
            
            Else
                
                ab_position = pSelStart + 1
            
            End If
            
        Else
        
            ab_position = 0
        
        End If
        
        zeilen_anzahl_a = cls_string_array_a.getAnzahlStrings

        zeilen_zaehler_a = 1

        zeilen_anzahl_b = cls_string_array_b.getAnzahlStrings

        zeilen_zaehler_b = 1
        
        While ((zeilen_zaehler_a <= zeilen_anzahl_a) And (zeilen_zaehler_b <= zeilen_anzahl_b))
        
            '
            ' Die aktuelle Zeile wird aus String A genommen.
            '
            aktuelle_zeile = cls_string_array_a.getString(zeilen_zaehler_a)
            
            '
            ' Die hinzuzufuegende Zeile kommt aus String B.
            '
            temp_string_3 = cls_string_array_b.getString(zeilen_zaehler_b)

            If (aktuelle_zeile = LEER_STRING) Then
            
                temp_string_1 = LEER_STRING
                
                temp_string_2 = LEER_STRING
            
            Else
    
                If (Len(aktuelle_zeile) >= ab_position) Then
    
                    temp_string_1 = Left(aktuelle_zeile, ab_position - 1)
                    
                    temp_string_2 = Mid(aktuelle_zeile, ab_position, Len(aktuelle_zeile))
    
                Else
    
                    temp_string_1 = aktuelle_zeile
                    
                    temp_string_2 = LEER_STRING
    
                End If
    
            End If
            
            '
            ' Ergebnisstring
            ' Dem Ergebnisstring wird eine neue Zeile hinzugefuegt.
            '
            str_fkt_ergebnis = str_fkt_ergebnis & zeichen_zeilenumbruch_b & temp_string_1 & temp_string_3 & temp_string_2
            
            zeichen_zeilenumbruch_b = zeichen_zeilenumbruch_a
            
            '
            ' Beide Zeilenzaehler werden um 1 erhoeht
            '
            zeilen_zaehler_a = zeilen_zaehler_a + 1
            
            zeilen_zaehler_b = zeilen_zaehler_b + 1
        
        Wend
            
    End If
     
    placeStringX = str_fkt_ergebnis

EndFunktion:

    On Error Resume Next
    
    Set cls_string_array_a = Nothing
    Set cls_string_array_b = Nothing
    
    DoEvents
    
    Exit Function
    
errStartPlaceStringX:
     
     placeStringX = "Fehler: " & Error
    
    Resume EndFunktion

End Function

'################################################################################
'
Public Function generatorVbNachJava(pString As String) As String

    generatorVbNachJava = LEER_STRING
    
    If (Trim(pString) = LEER_STRING) Then
        
        Exit Function
    
    End If

' VB   aktuelle_zeile = Mid( pString, aktuelle_startposition, ( FkString.len( pString ) - aktuelle_startposition ) + 1 );
' Java aktuelle_zeile = pString.substring( aktuelle_startposition, pString.length() );

Dim cls_string_array       As clsStringArray
Dim akt_index              As Long
Dim akt_zeile              As String
Dim anzahl_zeilen          As Long
Dim pos_kommentar          As Long
Dim str_place_var_dekl     As String
Dim such_string_var_dekl   As String
Dim temp_pos_1             As Long
Dim temp_pos_2             As Long
Dim temp_pos_3             As Long
Dim temp_pos_4             As Long
Dim temp_string_1          As String
Dim temp_string_2          As String
Dim temp_string_3          As String
Dim temp_string_4          As String
Dim temp_string_5          As String
Dim append_to_akt_zeile    As String
Dim var_deklaration        As String
Dim vb_str                 As String
Dim zaehler_kommentar      As Integer
Dim zeichen_zeilenumbruch  As String
Dim ausgabe_anweisung      As String
Dim knz_bearbeiten         As Boolean

On Error Resume Next

    such_string_var_dekl = "place_here_the_var_dekl_mach_it"
    str_place_var_dekl = such_string_var_dekl
    ausgabe_anweisung = "System.out.println"
    ausgabe_anweisung = "DrLogger.wl"

    zeichen_zeilenumbruch = getBenutztesChr13(pString)

    vb_str = pString
    vb_str = Replace(vb_str, "False", "false")
    vb_str = Replace(vb_str, "True", "true")
    vb_str = Replace(vb_str, "chr(13)", """\n""")
    vb_str = Replace(vb_str, "Trim(", "FkString.trim(")
    vb_str = Replace(vb_str, "Len(", "FkString.len(")
    vb_str = Replace(vb_str, "Left(", "FkString.left(")
    vb_str = Replace(vb_str, "Right(", "FkString.right(")
    vb_str = Replace(vb_str, "chr( 13 )", """\n""")
    vb_str = Replace(vb_str, "End If", "}")
    vb_str = Replace(vb_str, "end if", "}")
    vb_str = Replace(vb_str, "ElseIf", "} else if ")
    vb_str = Replace(vb_str, "Else", "} else {")
    vb_str = Replace(vb_str, "Then", "{")
    vb_str = Replace(vb_str, "Wend", "}")
    vb_str = Replace(vb_str, "Loop", "}")
    vb_str = Replace(vb_str, "Exit Do", "break;")
    vb_str = Replace(vb_str, "On Error Resume Next", "//On Error R esume Next")
    vb_str = Replace(vb_str, "On Error GoTo", "//On Error GoTo")
    vb_str = Replace(vb_str, " Resume ", " //Resume ")
    vb_str = Replace(vb_str, " If ", " if ")
    vb_str = Replace(vb_str, "Public Sub", "public static void")
    vb_str = Replace(vb_str, "Private Sub", "private static void")
    vb_str = Replace(vb_str, "Public Const ", "public static String ")
    vb_str = Replace(vb_str, "Private Const ", "private static String ")
    vb_str = Replace(vb_str, "Const ", "public static String ")
    vb_str = Replace(vb_str, "End Function", "}")
    vb_str = Replace(vb_str, "End Sub", "}")
    vb_str = Replace(vb_str, "Exit Function", "return;")
    vb_str = Replace(vb_str, "Exit Sub", "return;")
    vb_str = Replace(vb_str, "'""", "#-1-#")
    vb_str = Replace(vb_str, """'", "#-2-#")
    vb_str = Replace(vb_str, "= '", "#-3-#")
    vb_str = Replace(vb_str, "' )", "#-4-#")
    vb_str = Replace(vb_str, "'", "//")
    vb_str = Replace(vb_str, "#-1-#", "'""")
    vb_str = Replace(vb_str, "#-2-#", """'")
    vb_str = Replace(vb_str, "#-3-#", "= '")
    vb_str = Replace(vb_str, "#-4-#", "' )")
    
    
    vb_str = Replace(vb_str, " ByVal ", " ")
    vb_str = Replace(vb_str, "&", "+")
    vb_str = Replace(vb_str, """"""" +", "\"""" +")
    vb_str = Replace(vb_str, "+ """"""", "+ ""\""")
    ' """ &
    vb_str = Replace(vb_str, "Public Function", "public static")
    vb_str = Replace(vb_str, "Private Function", "private static")

    Set cls_string_array = startMultiline(vb_str)

    pos_kommentar = -1
    
    akt_index = 1
    
    anzahl_zeilen = cls_string_array.getAnzahlStrings()
    
    zaehler_kommentar = 0
    
    While (akt_index <= anzahl_zeilen)
        
        append_to_akt_zeile = LEER_STRING
        
        knz_bearbeiten = True
        
        akt_zeile = cls_string_array.getString(akt_index)
        
        If (InStr(akt_zeile, "str_debug = ""A - ") > 0) Then
        
            Debug.Print akt_zeile
        
        End If
        
        If (Trim(akt_zeile) = LEER_STRING) Then
        
            akt_zeile = LEER_STRING
            
            knz_bearbeiten = False
        
        End If

'#################################################################################
' Kommentarbehandlung
'
       If (knz_bearbeiten) Then
            
            temp_pos_4 = pos_kommentar ' letzte Positon aus letztem Durchgang

            pos_kommentar = InStr(akt_zeile, "//")

            temp_string_5 = Mid(akt_zeile, 1, pos_kommentar - 1)

            If (Len(Trim(temp_string_5)) > 0) Then

                ' vorm komma steht was

                akt_zeile = temp_string_5 & "; " & Mid(akt_zeile, pos_kommentar, Len(akt_zeile))

                pos_kommentar = InStr(akt_zeile, "//")

            End If

            If (pos_kommentar > 0) Then

                temp_string_1 = cls_string_array.getString(akt_index + 1)

                temp_pos_1 = InStr(temp_string_1, "//")

                If (temp_pos_1 = pos_kommentar) Then

                    zaehler_kommentar = zaehler_kommentar + 1

                    If (zaehler_kommentar = 1) Then

                        akt_zeile = Replace(akt_zeile, "//", "//S0")

                    Else

                        akt_zeile = Replace(akt_zeile, "//", "//S1")

                    End If

                Else

                    If (zaehler_kommentar > 0) Then

                        If (Trim(Mid(akt_zeile, pos_kommentar + 2, Len(akt_zeile))) = LEER_STRING) Then
    
                            akt_zeile = Replace(akt_zeile, "//", "//S2")
    
                        Else
    
                            akt_zeile = Replace(akt_zeile, "//", "//S1")
    
                            append_to_akt_zeile = zeichen_zeilenumbruch & String(pos_kommentar, " ") & "//S2"
    
                        End If

                    End If

                     zaehler_kommentar = 0

                End If

            Else

                zaehler_kommentar = 0

            End If

        End If
        
'#################################################################################
' For var_name = 1 to 100 step 2
'
        If (knz_bearbeiten) Then
        
            temp_pos_1 = InStr(1, akt_zeile, "For ")

            temp_pos_2 = InStr(1, akt_zeile, " To ")

            If ((temp_pos_1 > 0) And (temp_pos_2 > 0) And (temp_pos_2 > temp_pos_1)) Then

                temp_pos_3 = InStr(temp_pos_1 + 4, akt_zeile, "=")

                temp_pos_4 = InStr(temp_pos_2 + 4, akt_zeile, " Step ")

                If (temp_pos_4 > 0) Then

                    temp_string_4 = " + " + Trim(Mid(akt_zeile, temp_pos_4 + 6, Len(akt_zeile)))

                Else

                    temp_pos_4 = Len(akt_zeile) + 1

                    temp_string_4 = "++"

                End If

                temp_string_1 = Trim(Mid(akt_zeile, temp_pos_1 + 4, temp_pos_3 - (temp_pos_1 + 4)))

                temp_string_2 = Trim(Mid(akt_zeile, temp_pos_3 + 1, temp_pos_2 - (temp_pos_3 + 1)))

                temp_string_3 = Trim(Mid(akt_zeile, temp_pos_2 + 4, temp_pos_4 - (temp_pos_2 + 4)))

                akt_zeile = "for ( " & temp_string_1 & " = " & temp_string_2 & ", " & temp_string_1 & " <= " & temp_string_3 & ", " & temp_string_1 & temp_string_4 & " ) {"

                knz_bearbeiten = False

            End If
        
        End If
    
'#################################################################################
' IF-Zeile
'
        If (knz_bearbeiten) Then
            
            temp_pos_1 = InStr(1, akt_zeile, " if ")
    
            If (temp_pos_1 > 0) Then
    
                akt_zeile = Replace(akt_zeile, " = ", " == ")
    
                akt_zeile = Replace(akt_zeile, " And ", " && ")
    
                akt_zeile = Replace(akt_zeile, " Or ", " || ")
    
                akt_zeile = Replace(akt_zeile, " Not ", " ! ")
    
                akt_zeile = Replace(akt_zeile, " and ", " && ")
    
                akt_zeile = Replace(akt_zeile, " or ", " || ")
    
                akt_zeile = Replace(akt_zeile, " not ", " ! ")
    
                akt_zeile = Replace(akt_zeile, " Then ", " { ")
    
                akt_zeile = Replace(akt_zeile, " then ", " { ")
    
                knz_bearbeiten = False
    
            End If
        
        End If
        
'#################################################################################
' While-Zeile
'
        If (knz_bearbeiten) Then
            
            temp_pos_1 = InStr(1, akt_zeile, " While ")
            
            If (temp_pos_1 > 0) Then
            
                akt_zeile = Replace(akt_zeile, " = ", " == ")
    
                akt_zeile = Replace(akt_zeile, " While ", " while ( ")
    
                akt_zeile = Replace(akt_zeile, " And ", " && ")
    
                akt_zeile = Replace(akt_zeile, " Or ", " || ")
    
                akt_zeile = Replace(akt_zeile, " Not ", " ! ")
    
                akt_zeile = Replace(akt_zeile, " and ", " && ")
    
                akt_zeile = Replace(akt_zeile, " or ", " || ")
    
                akt_zeile = Replace(akt_zeile, " not ", " ! ")
    
                akt_zeile = akt_zeile & " ) {"
    
                knz_bearbeiten = False
                
            End If
        
        End If
    
'#################################################################################
' Semikolon ans Ende bei Auftreten von Call oder bei Auftreten von Zuweisung
'
        If (knz_bearbeiten) Then
    
            temp_string_4 = ";"
    
            If (InStr(akt_zeile, "Call ") > 0) Then
    
                akt_zeile = Replace(akt_zeile, "Call ", "")
    
                akt_zeile = akt_zeile & ";"
    
                temp_string_4 = LEER_STRING
    
                knz_bearbeiten = False
    
            End If
    
            temp_pos_1 = InStr(1, akt_zeile, "=")
    
            If (temp_pos_1 > 0) Then
    
                akt_zeile = Replace(akt_zeile, " & ", " + ")
    
                If (pos_kommentar > 0) Then
    
                    If (pos_kommentar < temp_pos_1) Then
    
                        temp_string_4 = LEER_STRING ' kein Semikolon, wenn Gleichheitszeichen Bestandteil eines Kommentares ist
    
                    End If
    
                End If
    
                akt_zeile = akt_zeile & temp_string_4
    
                knz_bearbeiten = False
    
            End If
    
        End If
    
'#################################################################################
' Variablendeklaration mit DIM var_name As Typ // Kommentar
'
        If (knz_bearbeiten) Then
    
            temp_pos_1 = InStr(1, akt_zeile, "Dim")
    
            temp_pos_2 = InStr(1, akt_zeile, " As ")
    
            If (temp_pos_1 > 0) And (temp_pos_2 > 0) Then
    
                append_to_akt_zeile = LEER_STRING
    
                temp_pos_3 = InStr(1, akt_zeile, "//")
    
                If (temp_pos_3 > 0) Then
    
                    temp_string_2 = Trim(Mid(akt_zeile, temp_pos_2 + 4, temp_pos_3 - (temp_pos_2 + 4)))
    
                    temp_string_4 = " " & Trim(Mid(akt_zeile, temp_pos_3, Len(akt_zeile)))
    
                Else
    
                    temp_string_2 = Trim(Mid(akt_zeile, temp_pos_2 + 4, Len(akt_zeile)))
    
                    temp_pos_3 = Len(akt_zeile)
    
                    temp_string_4 = LEER_STRING
    
                End If
    
                If (temp_string_2 = "String") Then
    
                    temp_string_3 = """"""
    
                ElseIf (temp_string_2 = "Double") Then
    
                    temp_string_3 = "0.00"
    
                    temp_string_2 = "double"
    
                ElseIf (temp_string_2 = "Float") Then
    
                    temp_string_3 = "0.00"
    
                    temp_string_2 = "double"
    
                ElseIf (temp_string_2 = "Long") Then
    
                    temp_string_3 = "0"
    
                    temp_string_2 = "long"
    
                ElseIf (temp_string_2 = "Integer") Then
    
                    temp_string_3 = "0"
    
                    temp_string_2 = "int"
    
                ElseIf (temp_string_2 = "Boolean") Then
    
                    temp_string_3 = "false"
    
                    temp_string_2 = "boolean"
    
                Else
    
                    temp_string_3 = "null"
    
                    temp_pos_4 = InStr(1, temp_string_2, ";")
    
                    If (temp_pos_4 > 0) Then
    
                        temp_string_2 = Trim(Mid(temp_string_2, 1, temp_pos_4 - 1))
    
                    End If
    
                End If
    
                temp_string_1 = Trim(Mid(akt_zeile, temp_pos_1 + 3, temp_pos_2 - (temp_pos_1 + 3)))
    
                var_deklaration = var_deklaration & zeichen_zeilenumbruch & Left(temp_string_2 & "                 ", 14) & " " & temp_string_1 & " = " & temp_string_3 & ";" & temp_string_4
    
                akt_zeile = str_place_var_dekl
    
                str_place_var_dekl = LEER_STRING
    
            End If
    
        End If

'#################################################################################
' Variablendeklaration mit Private/Public var_name As Typ // Kommentar
'
        If (knz_bearbeiten) Then
        
            temp_string_5 = "Private "
    
            temp_pos_1 = InStr(1, akt_zeile, temp_string_5)
    
            If (temp_pos_1 < 1) Then
    
               temp_string_5 = "Public "
    
               temp_pos_1 = InStr(1, akt_zeile, temp_string_5)
    
            End If
    
             temp_pos_2 = InStr(1, akt_zeile, " As ")
    
            If (temp_pos_1 > 0) And (temp_pos_2 > 0) Then
    
            temp_string_5 = LCase(temp_string_5)
    
                append_to_akt_zeile = LEER_STRING
    
                temp_pos_3 = InStr(1, akt_zeile, "//")
    
                If (temp_pos_3 > 0) Then
    
                    temp_string_2 = Trim(Mid(akt_zeile, temp_pos_2 + 4, temp_pos_3 - (temp_pos_2 + 4)))
    
                    temp_string_4 = " " & Trim(Mid(akt_zeile, temp_pos_3, Len(akt_zeile)))
    
                Else
    
                    temp_string_2 = Trim(Mid(akt_zeile, temp_pos_2 + 4, Len(akt_zeile)))
    
                    temp_pos_3 = Len(akt_zeile)
    
                    temp_string_4 = LEER_STRING
    
                End If
    
                If (temp_string_2 = "String") Then
    
                    temp_string_3 = """"""
    
                ElseIf (temp_string_2 = "Double") Then
    
                    temp_string_3 = "0.00"
    
                    temp_string_2 = "double"
    
                ElseIf (temp_string_2 = "Float") Then
    
                    temp_string_3 = "0.00"
    
                    temp_string_2 = "double"
    
                ElseIf (temp_string_2 = "Long") Then
    
                    temp_string_3 = "0"
    
                    temp_string_2 = "long"
    
                ElseIf (temp_string_2 = "Integer") Then
    
                    temp_string_3 = "0"
    
                    temp_string_2 = "int"
    
                ElseIf (temp_string_2 = "Boolean") Then
    
                    temp_string_3 = "false"
    
                    temp_string_2 = "boolean"
    
                End If
    
                temp_string_1 = Trim(Mid(akt_zeile, temp_pos_1 + Len(temp_string_5), temp_pos_2 - (temp_pos_1 + Len(temp_string_5))))
    
                akt_zeile = temp_string_5 & Left(temp_string_2 & "                 ", 14) & " " & temp_string_1 & " = " & temp_string_3 & ";" & temp_string_4
    
            End If
        
        End If

'#################################################################################
' Set var_name = nothing
'
        If (knz_bearbeiten) Then
    
            temp_pos_1 = InStr(1, akt_zeile, "Set ")
    
            temp_pos_2 = InStr(1, akt_zeile, "= Nothing")
    
            If (temp_pos_1 > 0) And (temp_pos_2 > 0) Then
    
                temp_string_1 = Trim(Mid(akt_zeile, temp_pos_1 + 3, temp_pos_2 - (temp_pos_1 + 3)))
    
                temp_pos_3 = InStr(1, akt_zeile, "//")
    
                If (temp_pos_3 > 0) Then
    
                    temp_string_4 = " " & Trim(Mid(akt_zeile, temp_pos_3 + 2, Len(akt_zeile)))
    
                Else
    
                    temp_string_4 = LEER_STRING
    
                End If
    
                akt_zeile = String(temp_pos_1, " ") & temp_string_1 & " = null;" & temp_string_4
    
                knz_bearbeiten = False
    
            End If
    
        End If

'#################################################################################
' Debug.Print Anweisungen konvertieren
'
        If (knz_bearbeiten) Then
    
            temp_pos_1 = InStr(1, akt_zeile, "Debug.Print ")
    
            If (temp_pos_1 > 0) Then
    
                temp_string_1 = Trim(Mid(akt_zeile, temp_pos_1 + 12, Len(akt_zeile)))
    
                akt_zeile = ausgabe_anweisung & "( " & temp_string_1 & " );"
    
                knz_bearbeiten = False
    
            End If
    
         End If
    
        Call cls_string_array.setString(akt_index, akt_zeile & append_to_akt_zeile)
        
        akt_index = akt_index + 1
    
    Wend
    
    vb_str = cls_string_array.toString(zeichen_zeilenumbruch)
    
    Set cls_string_array = Nothing
    
    var_deklaration = Replace(var_deklaration, "//S0", "//")
    var_deklaration = Replace(var_deklaration, "//S1//S2", "//")
    var_deklaration = Replace(var_deklaration, "//S2", "//")
    var_deklaration = Replace(var_deklaration, "//S1", "//")
    
    vb_str = Replace(vb_str, such_string_var_dekl, var_deklaration)
    vb_str = Replace(vb_str, "//S0", "/* ")
    vb_str = Replace(vb_str, "//S1//S2", " */")
    vb_str = Replace(vb_str, "//S2", " */")
    vb_str = Replace(vb_str, "//S1", " * ")
    vb_str = Replace(vb_str, " *  ", " * ")

    vb_str = Replace(vb_str, "Next", "} //")
    vb_str = Replace(vb_str, " <> ", " != ")
    vb_str = Replace(vb_str, " Set ", " ")
    vb_str = Replace(vb_str, " = Nothing", " = null")
    vb_str = Replace(vb_str, "} else {if", "} else if")
    vb_str = Replace(vb_str, "{;", "{")
    
    '
    ' Leerzeichen vor und nach Klammern
    '
    vb_str = Replace(Replace(Replace(Replace(vb_str, "(", "( "), ")", " )"), "[", "[ "), "]", " ]")
    
    '
    ' Eliminierung von doppelten Leerzeichen nach Klammern
    '
    vb_str = Replace(Replace(Replace(Replace(Replace(Replace(vb_str, "(   )", "()"), "[   ]", "[]"), "(   ", "( "), "   )", " )"), "[   ", "[ "), "   ]", " ]")
    vb_str = Replace(Replace(Replace(Replace(Replace(Replace(vb_str, "(  )", "()"), "[  ]", "[]"), "(  ", "( "), "  )", " )"), "[  ", "[ "), "  ]", " ]")
    vb_str = Replace(Replace(vb_str, "( )", "()"), "[ ]", "[]")
 
    generatorVbNachJava = vb_str

End Function

'##############################################################################
'
'      FkZahl.getZahl("+150.000,123456 Euro"         , 2    ) =  "150000.12"
'      FkZahl.getZahl("+150.000,123456 Euro"         , 0    ) =  "0"
'      FkZahl.getZahl("150.000,123456 DM"            , -1   ) =  "150000.123456"
'      FkZahl.getZahl("DM 150.000,123456-"           , 2    ) =  "-150000.12"
'      FkZahl.getZahl("null"                         , 2    ) =  "0.00"
'      FkZahl.getZahl("DM,Euro,Reichsmark"           , 2    ) =  "0.00"
'      FkZahl.getZahl("DM 15,0.0,00,12,34,56-"       , 2    ) =  "-15.00"
'      FkZahl.getZahl("100.12"                       , 2    ) =  "100.12"
'      FkZahl.getZahl("100.1234-"                    , 3    ) =  "-100.123"
'      FkZahl.getZahl("100,-"                        , 3    ) =  "-100.000"
'
' @param pString
' @param pAnzahlNachkommaStellen
' @return
'
Public Function getzahl(pString As String, pAnzahlNachkommaStellen As Integer, Optional pKnzFallbackTrennzeichenEin As Boolean = False) As String

Dim str_fkt_ergebnis  As String
Dim aktuelles_zeichen As String
Dim trennzeichen_nk   As String
Dim knz_negativ       As Boolean
Dim knz_nk_aktiv      As Integer
Dim zaehler           As Integer
Dim zaehler_nk        As Integer
Dim ziffern_zaehler   As Integer
    
    trennzeichen_nk = ","
    knz_negativ = False
    knz_nk_aktiv = 0
    zaehler = 1
    zaehler_nk = 0
    ziffern_zaehler = 0
    str_fkt_ergebnis = LEER_STRING

    If (Len(pString) > 0) Then
        '
        ' Hier wird ermittelt, ob das Nachkommatrennzeichen auf einen Punkt
        ' geaendert werden muss. Per Vorgabe wird das Komma als Trennzeichen
        ' genommen. Wird in der Eingabe kein Komma gefunden, wird der Punkt
        ' als Trennzeichen genommen.
        '
        ' Die Notwendigkeit ergab sich, da die Eingabe ja auch schon korrekt
        ' formatiert uebergeben werden kannm, z.B. aus den Werten einer DB.
        '
        ' Da dieses Vorgehen aber auch unerwartete Seiteneffekte haben kann,
        ' kann dieses Vorgehen von aussen mit einer boolschen Variable gesteuert
        ' werden.
        '
        If ((pKnzFallbackTrennzeichenEin) And (InStr(pString, ",") = 0)) Then
        
            trennzeichen_nk = "."
        
        End If

        While (zaehler <= Len(pString))
        
            aktuelles_zeichen = Mid(pString, zaehler, 1)
            
            If (IsNumeric(aktuelles_zeichen)) Then
                '
                ' Ist der Zaehler fuer die Nachkommastellen kleiner als die vorgegebene
                ' Anzahl der Nachkommastellen ist, wird die aktuelle Zahl dem Ergebnis
                ' hinzugefuegt.
                '
                ' Dieses wird auch dann gemacht, wenn die Anzahl der gewuenschten
                ' Nachkommastellen 0 ist. In einem solchen Fall wird dem Aufrufer
                ' die Eingabe nur in eine Zahl konvertiert.
                '
                If ((zaehler_nk < pAnzahlNachkommaStellen) Or (pAnzahlNachkommaStellen < 0)) Then
                
                    str_fkt_ergebnis = str_fkt_ergebnis & aktuelles_zeichen
                
                    zaehler_nk = zaehler_nk + knz_nk_aktiv ' knz_nk_aktiv = 1 wenn Leseprozess in Nachkommastellen
                
                    ziffern_zaehler = ziffern_zaehler + 1
                    
                End If
                
            End If
             
            If (aktuelles_zeichen = trennzeichen_nk) Then
                '
                ' Wenn das aktuelle Zeichen ein Komma ist, wird dieses beim
                ' ersten Auftretetn in einen Punkt gewandelt.
                '
                ' Das Umwandeln darf nicht doppelt gemacht werden.
                '
                If (knz_nk_aktiv = 0) Then
                
                    str_fkt_ergebnis = str_fkt_ergebnis & "."
                    
                End If
            
                '
                ' Kennzeichen wird auf 1 gesetzt
                '
                knz_nk_aktiv = 1
            
            End If
        
            If (aktuelles_zeichen = "-") Then
            
                knz_negativ = True
            
            End If

            zaehler = zaehler + 1
            
        Wend
        
        '
        ' In der Eingabe waren keine Zahlen vorhanden z.B. " Test,Test".
        ' In diesem Fall wuerde das Komma durch einen Punkt ersetzt werden.
        ' Der StringBuffer muss neu initialisiert werden, damit z.B. 0.00
        ' aus dem Rest der Routine hergestellt werden kann.
        '
        If (ziffern_zaehler = 0) Then
        
            str_fkt_ergebnis = LEER_STRING
            
        End If

    End If
    '
    ' Wenn die Eingabe null, ein Leerstring, oder durch vorangegangene
    ' Abfragen wieder ausgenullt worden ist, ist die Laenge des String-
    ' Buffers 0. Damit eine korrekte Zahl erstellt werden kann, wird
    ' eine fuehrende 0 hinzugefuegt.
    '
    If (Len(str_fkt_ergebnis) = 0) Then
    
        str_fkt_ergebnis = "0"
        
    End If

    '
    ' Hier werden die gewuenschten Anzahl der Nachkommastellen hinzugefuegt.
    '
    ' Wenn der Nachkommastellenzaehler noch 0 ist, muss noch ein Punkt hinzugefuegt werden
    '
    While (zaehler_nk < pAnzahlNachkommaStellen)
    
        If ((zaehler_nk = 0) And (knz_nk_aktiv = 0)) Then
        
            str_fkt_ergebnis = str_fkt_ergebnis & "."
            
        End If
        
        str_fkt_ergebnis = str_fkt_ergebnis & "0"
        
        zaehler_nk = zaehler_nk + 1
        
    Wend

    '
    ' Bei der Rueckgabe wird noch das Kennzeichen fuer einen negativen Betrag
    ' ausgewertet und gegebenenfalls ein Bindestrich dem Ergebnis hinzugefuegt.
    '
    getzahl = IIf(knz_negativ, "-", "") + str_fkt_ergebnis

End Function

'################################################################################
'
Private Function getStringMaxCols(pEingabe As String, pMaxAnzahlSpalten As Long, pEinzug As String, pNewLineZeichen As String) As String

Dim str_fkt_ergebnis     As String
Dim str_my_cr            As String
Dim neue_zeile           As String
Dim trenn_position_ab    As Long
Dim trenn_position_bis   As Long
Dim trenn_position_temp  As Long
Dim laenge_eingabe       As Long
Dim zaehler_while        As Long

    '
    ' Pruefung: Parameter pMaxAnzahlSpalten kleiner gleich 30?
    ' Ist der Parameter kleiner der Mindesspaltenanzahl von 30, wird die
    ' Anzahl der der Spalten auf die Vorgabe von 80 Stellen gesetzt.
    '
    If (pMaxAnzahlSpalten <= 30) Then

        pMaxAnzahlSpalten = 80

    End If

    '
    ' Pruefung: Laenge Eingabe kleiner als Max-Spaltenanzahl ?
    '
    ' Ist die Eingabe kuerzer als die maximale Spaltennazahl, ist das Ergebnis
    ' gleich der Eingabestring, da dieser nicht ueber die Max-Spalten hinaus geht.
    '
    If (Len(pEingabe) <= pMaxAnzahlSpalten) Then

        str_fkt_ergebnis = pEingabe

    Else
        '
        ' Ist die Eingabe laenger als die maximale Spaltenanzahl wird die
        ' Verkleinerungsschleife gestartet.
        '
        str_fkt_ergebnis = LEER_STRING

        laenge_eingabe = Len(pEingabe)
        trenn_position_ab = 0
        trenn_position_bis = 0
        zaehler_while = 0

        '
        ' Die Schleife laeuft solange wie
        ' ... die aktuelle Bis-Position noch kleiner der Laenge der Eingabe ist.
        ' ... der Endlosschleifenverhinderungszaehler kleiner 32123 ist.
        '
        While ((trenn_position_bis < laenge_eingabe) And (zaehler_while < 32123))
        
            '
            ' Trennposition Ab
            ' Die aktuelle Position ab welcher die Eingabe herausgeschnitten
            ' wird liegt 1 Zeichen hinter der letzten Bis-Position.
            '
            trenn_position_ab = trenn_position_bis + 1
            
            '
            ' Trennposition Bis
            ' Die naechste Trennposition-Bis liegt "pMaxAnzahlSpalten" hinter
            ' der aktuellen Startposition und dort dann beim ersten Leerzeichen.
            '
            trenn_position_bis = InStr(trenn_position_bis + pMaxAnzahlSpalten, pEingabe, " ")
            
            '
            ' Pruefung: Leerzeichen gefunden?
            ' Konnte kein Leerzeichen mehr gefunden werden, ist die naechste
            ' Bis-Trennposition gleich der Laenge der Eingabe.
            '
            If (trenn_position_bis = 0) Then

                trenn_position_bis = laenge_eingabe
            '
            ' Pruefung: Liegt die Bis-Trennposition 10 Zeichen vor Stringende?
            ' Beim letzten Durchgang kann es passieren, dass von der
            ' Eingabe nur noch weniger als 10 Zeichen vorhanden sind.
            ' Damit jetzt keine weitere kurze Zeile mehr erstellt wird,
            ' wird die Bis-Trennposition auf das Stringende gestellt.
            '
            ElseIf (trenn_position_bis + 10 >= laenge_eingabe) Then

                trenn_position_bis = laenge_eingabe

            '
            ' Pruefung: Ueberschreitung der Spaltenanzahl um 10 Zeichen?
            '
            ' Ziel soll es sein, dass kein "Flattersatz" entsteht. Es kann sein, das
            ' die vorgeschriebene Breite um mehr als 10 Zeichen ueberschritten wird.
            '
            ' Es kann sein, dass eventuell X-Zeichen (hier 3 Zeichen) vor
            ' dem eigentlichen Abschneide-Ende ein Leerzeichen liegt, welches
            ' dazu fuehren wuerde, dass weniger "Flattersatz" entstehen wuerde.
            '
            ' Wenn ein solches Leerzeichen gefunden wird, wird jenes genommen.
            '
            ElseIf (((trenn_position_bis - trenn_position_ab) - pMaxAnzahlSpalten) > 10) Then

                trenn_position_temp = InStr((trenn_position_ab + pMaxAnzahlSpalten) - 3, pEingabe, " ")

                If ((trenn_position_temp > 0) And (trenn_position_temp < trenn_position_bis)) Then

                    trenn_position_bis = trenn_position_temp

                End If

            End If
            
            '
            ' Pruefung: Ueberschreitung um mehr als 10 Zeichen?
            '
            ' Wenn dem so ist, wird eine harte Trennung vorgenommen.
            ' Dieses kann z.B. BASE64 kodierten Zeichenketten vorkommen.
            '
            If (((trenn_position_bis - trenn_position_ab) - pMaxAnzahlSpalten) > 10) Then
                
                trenn_position_bis = trenn_position_ab + pMaxAnzahlSpalten
                
            End If

            '
            ' Pruefung: Zeilenumbruch ?
            ' Ab der aktuellen Trennstartposition wird die Position des naechsten
            ' Zeilenumbruchszeichen gesucht.
            '
            ' Liegt dieses Zeichen zwischen den aktuellen Grenzen Von und Ab,
            ' wird die Bis-Position auf die Position des Zeilenumbruches gesetzt.
            '
            trenn_position_temp = InStr(trenn_position_ab, pEingabe, pNewLineZeichen)
            
            If ((trenn_position_temp > trenn_position_ab) And (trenn_position_temp < trenn_position_bis)) Then
            
                trenn_position_bis = trenn_position_temp
            
            End If

            '
            ' Bestimmung neue Zeile
            ' Aus dem Eingabestring wird der ermittelte Teilstring herausgetrennt und getrimmt.
            ' Ist der Teilstring kein Leerstring, wird dieser dem Ergebnis zugewiesen.
            '
            neue_zeile = Trim(Mid(pEingabe, trenn_position_ab, (trenn_position_bis - trenn_position_ab) + 1))

            If (neue_zeile <> LEER_STRING) Then

                str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr & neue_zeile

            End If

            str_my_cr = pNewLineZeichen & pEinzug

            zaehler_while = zaehler_while + 1

        Wend

    End If

    getStringMaxCols = str_fkt_ergebnis

End Function

'################################################################################
'
Private Function extrahiereWoerter(pText As String, pTrennzeichen As String, pMaxErgebnisZeilenlaenge As Long) As String

On Error GoTo errExtrahiereWoerter

    extrahiereWoerter = LEER_STRING

Dim ergebnis_string_gesamt     As String
Dim ergebnis_string_zeile      As String
Dim ergebnis_wort_trennzeichen As String
Dim ergebnis_max_zeilenlaenge  As Long
Dim ergebnis_wort_min_laenge   As Long

Dim parser_akt_position        As Long
Dim parser_akt_zeichen         As String
Dim parser_temp_wort           As String
Dim parser_wort_trennzeichen   As String

    If (Trim(pTrennzeichen) = LEER_STRING) Then
    
        ergebnis_wort_trennzeichen = ergebnis_wort_trennzeichen
    
    Else
     
        ergebnis_wort_trennzeichen = pTrennzeichen
        
    End If
    
    ergebnis_string_gesamt = LEER_STRING
    ergebnis_string_zeile = LEER_STRING
    ergebnis_wort_min_laenge = 3
    
    If (pMaxErgebnisZeilenlaenge < 0) Then
    
        ergebnis_max_zeilenlaenge = 150
    
    Else
    
        ergebnis_max_zeilenlaenge = pMaxErgebnisZeilenlaenge
        
    End If
    
    '
    ' Unterstrich kein Trennzeichen, da in Variablennamen benutzt
    '
    parser_wort_trennzeichen = "[]{}()-+:=\/:*?!#<> |.,;&""" & vbCr & vbTab & vbLf
    
    parser_akt_position = 1
    
    '
    ' While-Schleife ueber alle Zeichen der Eingabe
    '
    While (parser_akt_position <= Len(pText))
    
        '
        ' Zeichen an der aktuellen Leseposition ermitteln
        '
        parser_akt_zeichen = Mid(pText, parser_akt_position, 1)
        
        If ((InStr(parser_wort_trennzeichen, parser_akt_zeichen) = 0) And (Asc(parser_akt_zeichen) > 30)) Then
        
            '
            ' Ist das aktuelle Zeichen nicht im String der Worttrennzeichen vorhanden
            ' und das Zeichen groesser ASCII 30 ist, wird das Zeichen dem aktuellem
            ' Wort hinzuaddiert.
            '
            
            parser_temp_wort = parser_temp_wort + parser_akt_zeichen
        
        Else
        
            '
            ' Ist das aktuelle Zeichen ein Worttrennzeichen, wird das Wort dem
            ' Ergebnis hinzugefuegt. Es gibt 2 Bedingungen fuer die Aufnahme:
            '
            ' 1. Das aktuelle Wort muss eine Mindestlaenge erfuellen
            ' 2. Das aktuelle Wort darf nicht schon einmal aufgenommen worden sein
            '
            ' Sind die beiden Bedingungen erfuellt, wird das Wort dem Ergebnisstring
            ' hinzugefuegt. Dabei wird auf die Mindestlaenge einer Ergebniszeile
            ' geprueft und gegebenenfalls wird ein Zeilenumbruch hinzugefuegt.
        
            If (Len((parser_temp_wort)) >= ergebnis_wort_min_laenge) Then
                
                If (InStr(ergebnis_string_gesamt, parser_temp_wort & ergebnis_wort_trennzeichen) = 0) Then
                
                    If (InStr(ergebnis_string_zeile, parser_temp_wort & ergebnis_wort_trennzeichen) = 0) Then
                    
                        ergebnis_string_zeile = ergebnis_string_zeile & " " & parser_temp_wort & ergebnis_wort_trennzeichen
                        
                        If (Len(ergebnis_string_zeile) > ergebnis_max_zeilenlaenge) Then
                        
                            ergebnis_string_gesamt = ergebnis_string_gesamt & vbCrLf & ergebnis_string_zeile
                        
                            ergebnis_string_zeile = LEER_STRING
                            
                        End If
                    
                    End If
                
                End If
                
            End If
        
            parser_temp_wort = LEER_STRING
        
        End If
        
        parser_akt_position = parser_akt_position + 1
    
    Wend

    extrahiereWoerter = ergebnis_string_gesamt

EndFunktion:

    On Error Resume Next

    DoEvents

    Exit Function

errExtrahiereWoerter:

    extrahiereWoerter = ("Fehler: errExtrahiereWoerter: " & Err & " " & Error & " " & Erl)

    Resume EndFunktion

End Function

'################################################################################
'
Public Function getUrlDecoded(pString As String) As String

On Error GoTo errGetUrlDecoded

Dim fkt_ergebnis As String

    fkt_ergebnis = Replace(pString, "%20", " ")
    
    fkt_ergebnis = Replace(fkt_ergebnis, "%2d", "-")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3d", "=")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3a", ":")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2c", ",")
    fkt_ergebnis = Replace(fkt_ergebnis, "%40", "@")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2F", "/")
    fkt_ergebnis = Replace(fkt_ergebnis, "%5C", "\")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2e", ".")
    fkt_ergebnis = Replace(fkt_ergebnis, "%26", "&")
    fkt_ergebnis = Replace(fkt_ergebnis, "%26", "&")
    fkt_ergebnis = Replace(fkt_ergebnis, "%22", """")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3C", "<")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3E", ">")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DF", "ﬂ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E4", "‰")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F6", "ˆ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FC", "¸")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C4", "ƒ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D6", "÷")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DC", "‹")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3d", "=")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3a", ":")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2c", ",")
    fkt_ergebnis = Replace(fkt_ergebnis, "%40", "@")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2F", "/")
    fkt_ergebnis = Replace(fkt_ergebnis, "%5C", "\")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2e", ".")
    fkt_ergebnis = Replace(fkt_ergebnis, "%26", "&")
    fkt_ergebnis = Replace(fkt_ergebnis, "%20", " ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%21", "!")
    fkt_ergebnis = Replace(fkt_ergebnis, "%23", "#")
    fkt_ergebnis = Replace(fkt_ergebnis, "%24", "$")
    fkt_ergebnis = Replace(fkt_ergebnis, "%25", "%")
    fkt_ergebnis = Replace(fkt_ergebnis, "%26", "&")
    fkt_ergebnis = Replace(fkt_ergebnis, "%27", "'")
    fkt_ergebnis = Replace(fkt_ergebnis, "%28", "(")
    fkt_ergebnis = Replace(fkt_ergebnis, "%29", ")")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2A", "*")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2B", "+")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2C", ",")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2D", "-")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2E", ".")
    fkt_ergebnis = Replace(fkt_ergebnis, "%2F", "/")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3A", ":")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3B", ";")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3C", "<")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3D", "=")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3E", ">")
    fkt_ergebnis = Replace(fkt_ergebnis, "%3F", "?")
    fkt_ergebnis = Replace(fkt_ergebnis, "%40", "@")
    fkt_ergebnis = Replace(fkt_ergebnis, "%5B", "[")
    fkt_ergebnis = Replace(fkt_ergebnis, "%5C", "\")
    fkt_ergebnis = Replace(fkt_ergebnis, "%5D", "]")
    fkt_ergebnis = Replace(fkt_ergebnis, "%5E", "^")
    fkt_ergebnis = Replace(fkt_ergebnis, "%5F", UNTER_STRICH)
    fkt_ergebnis = Replace(fkt_ergebnis, "%60", "`")
    fkt_ergebnis = Replace(fkt_ergebnis, "%7B", "{")
    fkt_ergebnis = Replace(fkt_ergebnis, "%7C", "|")
    fkt_ergebnis = Replace(fkt_ergebnis, "%7D", "}")
    fkt_ergebnis = Replace(fkt_ergebnis, "%7E", "~")
    fkt_ergebnis = Replace(fkt_ergebnis, "%7F", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "%80", "`")
    fkt_ergebnis = Replace(fkt_ergebnis, "%81", "Å")
    fkt_ergebnis = Replace(fkt_ergebnis, "%82", "Ç")
    fkt_ergebnis = Replace(fkt_ergebnis, "%83", "É")
    fkt_ergebnis = Replace(fkt_ergebnis, "%84", "Ñ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%85", "Ö")
    fkt_ergebnis = Replace(fkt_ergebnis, "%86", "Ü")
    fkt_ergebnis = Replace(fkt_ergebnis, "%87", "á")
    fkt_ergebnis = Replace(fkt_ergebnis, "%88", "à")
    fkt_ergebnis = Replace(fkt_ergebnis, "%89", "â")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8A", "ä")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8B", "ã")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8C", "å")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8D", "ç")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8E", "é")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8F", "è")
    fkt_ergebnis = Replace(fkt_ergebnis, "%90", "ê")
    fkt_ergebnis = Replace(fkt_ergebnis, "%91", "ë")
    fkt_ergebnis = Replace(fkt_ergebnis, "%92", "í")
    fkt_ergebnis = Replace(fkt_ergebnis, "%93", "ì")
    fkt_ergebnis = Replace(fkt_ergebnis, "%94", "î")
    fkt_ergebnis = Replace(fkt_ergebnis, "%95", "ï")
    fkt_ergebnis = Replace(fkt_ergebnis, "%96", "ñ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%97", "ó")
    fkt_ergebnis = Replace(fkt_ergebnis, "%98", "ò")
    fkt_ergebnis = Replace(fkt_ergebnis, "%99", "ô")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9A", "ö")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9B", "õ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9C", "ú")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9D", "ù")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9E", "û")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9F", "ü")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A0", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A1", "°")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A2", "¢")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A3", "£")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A4", "§")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A5", "•")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A6", "¶")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A7", "ß")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A8", "®")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A9", "©")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AA", "™")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AB", "´")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AC", "¨")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AD", "≠")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AE", "Æ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AF", "Ø")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B0", "∞")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B1", "±")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B2", "≤")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B3", "≥")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B4", "¥")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B5", "µ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B6", "∂")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B7", "∑")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B8", "∏")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B9", "π")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BA", "∫")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BB", "ª")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BC", "º")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BD", "Ω")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BE", "æ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BF", "ø")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C0", "¿")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C1", "¡")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C2", "¬")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C3", "√")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C4", "ƒ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C5", "≈")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C6", "∆")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C7", "«")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C8", "»")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C9", "…")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CA", " ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CB", "À")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CC", "Ã")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CD", "Õ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CE", "Œ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CF", "œ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D0", "–")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D1", "—")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D2", "“")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D3", "”")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D4", "‘")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D5", "’")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D6", "÷")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D7", "◊")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D8", "ÿ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D9", "Ÿ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DA", "⁄")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DB", "€")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DC", "‹")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DD", "›")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DE", "ﬁ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DF", "ﬂ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E0", "‡")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E1", "·")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E2", "‚")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E3", "„")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E4", "‰")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E5", "Â")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E6", "Ê")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E7", "Á")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E8", "Ë")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E9", "È")
    fkt_ergebnis = Replace(fkt_ergebnis, "%EA", "Í")
    fkt_ergebnis = Replace(fkt_ergebnis, "%EB", "Î")
    fkt_ergebnis = Replace(fkt_ergebnis, "%EC", "Ï")
    fkt_ergebnis = Replace(fkt_ergebnis, "%ED", "Ì")
    fkt_ergebnis = Replace(fkt_ergebnis, "%EE", "Ó")
    fkt_ergebnis = Replace(fkt_ergebnis, "%EF", "Ô")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F0", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F1", "Ò")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F2", "Ú")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F3", "Û")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F4", "Ù")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F5", "ı")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F6", "ˆ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F7", "˜")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F8", "¯")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F9", "˘")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FA", "˙")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FB", "˚")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FC", "¸")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FD", "˝")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FE", "˛")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FF", "ˇ")

EndFunktion:

    On Error Resume Next

    '
    ' DoEvents aufrufen
    '
    DoEvents

    '
    ' Funktionsergebnis setzen
    '
    getUrlDecoded = fkt_ergebnis

    '
    ' Funktion verlassen
    '
    Exit Function

errGetUrlDecoded:

    Debug.Print ("Fehler: errGetUrlDecoded: " & Err & " " & Error & " " & Erl)

    Resume EndFunktion

End Function

'################################################################################
'
' Erstellt aus der Eingabe einen URL-Codierten String.
'
' https://www.w3schools.com/tags/ref_urlencode.asp
'
' PARAMETER: pString        = der zu behandelnde Eingabestring
'
' RETURN : Einen String mit den konvertieren Zeichen der Url-codierung
'
Public Function getUrlEncoded(ByVal pString As String, pKnzZUmbruch As Boolean) As String

Dim str_fkt_ergebnis      As String  ' Ergebnisstring fuer die Rueckgabe
Dim str_gueltige_zeichen  As String  ' Liste der gueltigen Zeichen
Dim akt_zeichen           As String  ' aktuelle Zeichen in der While-Schleife
Dim akt_position          As Integer ' aktuelle Leseposition der While-Schleife

    '
    ' Initialisierung des Strings mit den gueltigen Zeichen
    '
    str_gueltige_zeichen = "enirstaudhgolcmfbkVvwz1paeSDA0E2RBGueMIPKF9UNW3L78oeH4T5CZJy6xjOUeYXq" & vbCr & vbLf

    akt_position = 1

    '
    ' While-Schleife ueber alle Zeichen der Eingabe
    '
    While (akt_position <= Len(pString))

        '
        ' Zeichen aus der Eingabe an der aktuellen Lesepositon lesen
        '
        akt_zeichen = Mid(pString, akt_position, 1)
        
        '
        ' Pruefung auf Umwandlung des aktuellen Zeichens in die
        ' Codierung fuer URL-Adressangaben
        '
        If (akt_zeichen = " ") Then
        
            str_fkt_ergebnis = str_fkt_ergebnis & "%20"
        
        ElseIf ((pKnzZUmbruch) And (akt_zeichen = vbCr)) Then
        
            str_fkt_ergebnis = str_fkt_ergebnis & "%0d"
        
        ElseIf ((pKnzZUmbruch) And (akt_zeichen = vbLf)) Then
        
            str_fkt_ergebnis = str_fkt_ergebnis & "%0a"
        
        ElseIf (akt_zeichen = "-") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2d"

        ElseIf (akt_zeichen = "=") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3d"

        ElseIf (akt_zeichen = ":") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3a"

        ElseIf (akt_zeichen = ",") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2c"

        ElseIf (akt_zeichen = "@") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%40"

        ElseIf (akt_zeichen = "/") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2F"

        ElseIf (akt_zeichen = "\") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%5C"

        ElseIf (akt_zeichen = ".") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2e"

        ElseIf (akt_zeichen = "&") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%26"
        
        ElseIf (akt_zeichen = "&") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%26"

        ElseIf (akt_zeichen = """") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%22"

        ElseIf (akt_zeichen = "<") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3C"

        ElseIf (akt_zeichen = ">") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3E"

        ElseIf (akt_zeichen = "ﬂ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DF"

        ElseIf (akt_zeichen = "‰") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E4"

        ElseIf (akt_zeichen = "ˆ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F6"

        ElseIf (akt_zeichen = "¸") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FC"

        ElseIf (akt_zeichen = "ƒ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C4"

        ElseIf (akt_zeichen = "÷") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D6"

        ElseIf (akt_zeichen = "‹") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DC"

        ElseIf (akt_zeichen = "=") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3d"

        ElseIf (akt_zeichen = ":") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3a"

        ElseIf (akt_zeichen = ",") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2c"

        ElseIf (akt_zeichen = "@") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%40"

        ElseIf (akt_zeichen = "/") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2F"

        ElseIf (akt_zeichen = "\") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%5C"

        ElseIf (akt_zeichen = ".") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2e"

        ElseIf (akt_zeichen = "&") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%26"
        
        ElseIf (akt_zeichen = " ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%20"

        ElseIf (akt_zeichen = "!") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%21"

        ElseIf (akt_zeichen = "#") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%23"

        ElseIf (akt_zeichen = "$") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%24"

        ElseIf (akt_zeichen = "%") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%25"

        ElseIf (akt_zeichen = "&") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%26"

        ElseIf (akt_zeichen = "'") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%27"

        ElseIf (akt_zeichen = "(") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%28"

        ElseIf (akt_zeichen = ")") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%29"

        ElseIf (akt_zeichen = "*") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2A"

        ElseIf (akt_zeichen = "+") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2B"

        ElseIf (akt_zeichen = ",") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2C"

        ElseIf (akt_zeichen = "-") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2D"

        ElseIf (akt_zeichen = ".") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2E"

        ElseIf (akt_zeichen = "/") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%2F"

        ElseIf (akt_zeichen = ":") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3A"

        ElseIf (akt_zeichen = ";") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3B"

        ElseIf (akt_zeichen = "<") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3C"

        ElseIf (akt_zeichen = "=") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3D"

        ElseIf (akt_zeichen = ">") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3E"

        ElseIf (akt_zeichen = "?") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%3F"

        ElseIf (akt_zeichen = "@") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%40"

        ElseIf (akt_zeichen = "[") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%5B"

        ElseIf (akt_zeichen = "\") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%5C"

        ElseIf (akt_zeichen = "]") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%5D"

        ElseIf (akt_zeichen = "^") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%5E"

        ElseIf (akt_zeichen = UNTER_STRICH) Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%5F"

        ElseIf (akt_zeichen = "`") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%60"

        ElseIf (akt_zeichen = "{") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%7B"

        ElseIf (akt_zeichen = "|") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%7C"

        ElseIf (akt_zeichen = "}") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%7D"

        ElseIf (akt_zeichen = "~") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%7E"

        ElseIf (akt_zeichen = "") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%7F"

        ElseIf (akt_zeichen = "`") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%80"

        ElseIf (akt_zeichen = "Å") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%81"

        ElseIf (akt_zeichen = "Ç") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%82"

        ElseIf (akt_zeichen = "É") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%83"

        ElseIf (akt_zeichen = "Ñ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%84"

        ElseIf (akt_zeichen = "Ö") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%85"

        ElseIf (akt_zeichen = "Ü") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%86"

        ElseIf (akt_zeichen = "á") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%87"

        ElseIf (akt_zeichen = "à") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%88"

        ElseIf (akt_zeichen = "â") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%89"

        ElseIf (akt_zeichen = "ä") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8A"

        ElseIf (akt_zeichen = "ã") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8B"

        ElseIf (akt_zeichen = "å") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8C"

        ElseIf (akt_zeichen = "ç") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8D"

        ElseIf (akt_zeichen = "é") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8E"

        ElseIf (akt_zeichen = "è") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8F"

        ElseIf (akt_zeichen = "ê") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%90"

        ElseIf (akt_zeichen = "ë") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%91"

        ElseIf (akt_zeichen = "í") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%92"

        ElseIf (akt_zeichen = "ì") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%93"

        ElseIf (akt_zeichen = "î") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%94"

        ElseIf (akt_zeichen = "ï") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%95"

        ElseIf (akt_zeichen = "ñ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%96"

        ElseIf (akt_zeichen = "ó") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%97"

        ElseIf (akt_zeichen = "ò") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%98"

        ElseIf (akt_zeichen = "ô") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%99"

        ElseIf (akt_zeichen = "ö") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9A"

        ElseIf (akt_zeichen = "õ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9B"

        ElseIf (akt_zeichen = "ú") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9C"

        ElseIf (akt_zeichen = "ù") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9D"

        ElseIf (akt_zeichen = "û") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9E"

        ElseIf (akt_zeichen = "ü") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9F"

        ElseIf (akt_zeichen = "") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A0"

        ElseIf (akt_zeichen = "°") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A1"

        ElseIf (akt_zeichen = "¢") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A2"

        ElseIf (akt_zeichen = "£") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A3"

        ElseIf (akt_zeichen = "§") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A4"

        ElseIf (akt_zeichen = "•") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A5"

        ElseIf (akt_zeichen = "¶") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A6"

        ElseIf (akt_zeichen = "ß") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A7"

        ElseIf (akt_zeichen = "®") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A8"

        ElseIf (akt_zeichen = "©") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A9"

        ElseIf (akt_zeichen = "™") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AA"

        ElseIf (akt_zeichen = "´") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AB"

        ElseIf (akt_zeichen = "¨") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AC"

        ElseIf (akt_zeichen = "≠") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AD"

        ElseIf (akt_zeichen = "Æ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AE"

        ElseIf (akt_zeichen = "Ø") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AF"

        ElseIf (akt_zeichen = "∞") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B0"

        ElseIf (akt_zeichen = "±") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B1"

        ElseIf (akt_zeichen = "≤") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B2"

        ElseIf (akt_zeichen = "≥") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B3"

        ElseIf (akt_zeichen = "¥") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B4"

        ElseIf (akt_zeichen = "µ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B5"

        ElseIf (akt_zeichen = "∂") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B6"

        ElseIf (akt_zeichen = "∑") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B7"

        ElseIf (akt_zeichen = "∏") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B8"

        ElseIf (akt_zeichen = "π") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B9"

        ElseIf (akt_zeichen = "∫") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BA"

        ElseIf (akt_zeichen = "ª") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BB"

        ElseIf (akt_zeichen = "º") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BC"

        ElseIf (akt_zeichen = "Ω") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BD"

        ElseIf (akt_zeichen = "æ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BE"

        ElseIf (akt_zeichen = "ø") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BF"

        ElseIf (akt_zeichen = "¿") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C0"

        ElseIf (akt_zeichen = "¡") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C1"

        ElseIf (akt_zeichen = "¬") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C2"

        ElseIf (akt_zeichen = "√") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C3"

        ElseIf (akt_zeichen = "ƒ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C4"

        ElseIf (akt_zeichen = "≈") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C5"

        ElseIf (akt_zeichen = "∆") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C6"

        ElseIf (akt_zeichen = "«") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C7"

        ElseIf (akt_zeichen = "»") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C8"

        ElseIf (akt_zeichen = "…") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C9"

        ElseIf (akt_zeichen = " ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CA"

        ElseIf (akt_zeichen = "À") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CB"

        ElseIf (akt_zeichen = "Ã") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CC"

        ElseIf (akt_zeichen = "Õ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CD"

        ElseIf (akt_zeichen = "Œ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CE"

        ElseIf (akt_zeichen = "œ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CF"

        ElseIf (akt_zeichen = "–") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D0"

        ElseIf (akt_zeichen = "—") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D1"

        ElseIf (akt_zeichen = "“") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D2"

        ElseIf (akt_zeichen = "”") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D3"

        ElseIf (akt_zeichen = "‘") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D4"

        ElseIf (akt_zeichen = "’") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D5"

        ElseIf (akt_zeichen = "÷") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D6"

        ElseIf (akt_zeichen = "◊") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D7"

        ElseIf (akt_zeichen = "ÿ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D8"

        ElseIf (akt_zeichen = "Ÿ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D9"

        ElseIf (akt_zeichen = "⁄") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DA"

        ElseIf (akt_zeichen = "€") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DB"

        ElseIf (akt_zeichen = "‹") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DC"

        ElseIf (akt_zeichen = "›") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DD"

        ElseIf (akt_zeichen = "ﬁ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DE"

        ElseIf (akt_zeichen = "ﬂ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DF"

        ElseIf (akt_zeichen = "‡") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E0"

        ElseIf (akt_zeichen = "·") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E1"

        ElseIf (akt_zeichen = "‚") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E2"

        ElseIf (akt_zeichen = "„") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E3"

        ElseIf (akt_zeichen = "‰") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E4"

        ElseIf (akt_zeichen = "Â") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E5"

        ElseIf (akt_zeichen = "Ê") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E6"

        ElseIf (akt_zeichen = "Á") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E7"

        ElseIf (akt_zeichen = "Ë") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E8"

        ElseIf (akt_zeichen = "È") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E9"

        ElseIf (akt_zeichen = "Í") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%EA"

        ElseIf (akt_zeichen = "Î") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%EB"

        ElseIf (akt_zeichen = "Ï") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%EC"

        ElseIf (akt_zeichen = "Ì") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%ED"

        ElseIf (akt_zeichen = "Ó") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%EE"

        ElseIf (akt_zeichen = "Ô") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%EF"

        ElseIf (akt_zeichen = "") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F0"

        ElseIf (akt_zeichen = "Ò") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F1"

        ElseIf (akt_zeichen = "Ú") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F2"

        ElseIf (akt_zeichen = "Û") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F3"

        ElseIf (akt_zeichen = "Ù") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F4"

        ElseIf (akt_zeichen = "ı") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F5"

        ElseIf (akt_zeichen = "ˆ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F6"

        ElseIf (akt_zeichen = "˜") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F7"

        ElseIf (akt_zeichen = "¯") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F8"

        ElseIf (akt_zeichen = "˘") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F9"

        ElseIf (akt_zeichen = "˙") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FA"

        ElseIf (akt_zeichen = "˚") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FB"

        ElseIf (akt_zeichen = "¸") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FC"

        ElseIf (akt_zeichen = "˝") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FD"

        ElseIf (akt_zeichen = "˛") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FE"

        ElseIf (akt_zeichen = "ˇ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FF"

        Else
        
            '
            ' Pruefung: Gueltiges Zeichen
            ' Ist das aktuelle Zeichen in dem String der gueltigen Zeichen enthalten,
            ' wird das Zeichen in den Ergebnisstring uebernommen.
            '
            ' Ist das Zeichen nicht in den gueltigen Zeichen enthalten, wird es ueberlesen
            '
            If (InStr(str_gueltige_zeichen, akt_zeichen) > 0) Then
                
                str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
                
            End If

        End If

        '
        ' Es wird die Leseposition der While-Schleife erhoeht und
        ' mit dem naechsten Schleifedurchlauf weitergemacht.
        '
        akt_position = akt_position + 1

    Wend

    '
    ' Der Aufrufer bekommt den Ergebnisstring zurueck.
    '
    getUrlEncoded = str_fkt_ergebnis

End Function

'################################################################################
'
' ? quoteHtmlCharacter( "<tag_name> """ ) = &lt;tag_name&gt;
'
Public Function quoteHtmlCharacter(pHtmlString As String) As String

On Error GoTo errQuoteHtmlCharacter

Dim anzahl_zeichen       As Integer
Dim akt_index            As Integer
Dim akt_zeichen          As String
Dim html_ergebnis_string As String

    If (pHtmlString <> LEER_STRING) Then
    
        '
        ' Anzahl der Zeichen aus der Eingabe lesen
        '
        anzahl_zeichen = Len(pHtmlString)
        
        m_toggle_mr_stringer_fkt = Not m_toggle_mr_stringer_fkt
        
        '
        ' Startindex auf das erste Zeichen stellen
        '
        akt_index = 1
        
        '
        ' While-Schleife Konvertierung
        '
        While (akt_index <= anzahl_zeichen)
        
            '
            ' Es wird das Zeichen an der aktuellen Leseposition gelesen
            '
            akt_zeichen = Mid(pHtmlString, akt_index, 1)
            
            '
            ' Zeichen A-Z, a-z und 0-9 werden nicht ueber die If-Kaskade geprueft
            '
            If ((((Asc(akt_zeichen) >= 65) And (Asc(akt_zeichen) <= 90)) Or ((Asc(akt_zeichen) >= 97) And (Asc(akt_zeichen) <= 122)) Or ((Asc(akt_zeichen) >= 48) And (Asc(akt_zeichen) <= 57)) Or (akt_zeichen = " ")) = False) Then

                '
                ' Spezial-Zeichen werden uebersetzt
                '
                If (akt_zeichen = """") Then
                
                    If (m_toggle_mr_stringer_fkt) Then
                
                        akt_zeichen = "&quot;"
                        
                    End If
            
                ElseIf (akt_zeichen = "&") Then
                
                    If (m_toggle_mr_stringer_fkt) Then
            
                        akt_zeichen = "&amp;"
                    
                    End If
            
                ElseIf (akt_zeichen = " ") Then
                
                    If (m_toggle_mr_stringer_fkt) Then
            
                        akt_zeichen = "&nbsp;"
                    
                    End If
            
                ElseIf (akt_zeichen = "'") Then
                            
                    If (m_toggle_mr_stringer_fkt) Then
            
                        akt_zeichen = "&apos;"
                        
                    End If
            
                ElseIf (akt_zeichen = "<") Then
                
                    If (m_toggle_mr_stringer_fkt) Then
            
                        akt_zeichen = "&lt;"
                        
                    End If
            
                ElseIf (akt_zeichen = ">") Then
                
                    If (m_toggle_mr_stringer_fkt) Then
            
                        akt_zeichen = "&gt;"
                        
                    End If
            
                ElseIf (akt_zeichen = "°") Then
            
                    akt_zeichen = "&iexcl;"
            
                ElseIf (akt_zeichen = "¢") Then
            
                    akt_zeichen = "&cent;"
            
                ElseIf (akt_zeichen = "£") Then
            
                    akt_zeichen = "&pound;"
            
                ElseIf (akt_zeichen = "§") Then
            
                    akt_zeichen = "&curren;"
            
                ElseIf (akt_zeichen = "•") Then
            
                    akt_zeichen = "&yen;"
            
                ElseIf (akt_zeichen = "¶") Then
            
                    akt_zeichen = "&brvbar;"
            
                ElseIf (akt_zeichen = "ß") Then
            
                    akt_zeichen = "&sect;"
            
                ElseIf (akt_zeichen = "®") Then
            
                    akt_zeichen = "&uml;"
            
                ElseIf (akt_zeichen = "©") Then
            
                    akt_zeichen = "&copy;"
            
                ElseIf (akt_zeichen = "™") Then
            
                    akt_zeichen = "&ordf;"
            
                ElseIf (akt_zeichen = "´") Then
            
                    akt_zeichen = "&laquo;"
            
                ElseIf (akt_zeichen = "¨") Then
            
                    akt_zeichen = "&not;"
            
                ElseIf (akt_zeichen = "≠") Then
            
                    akt_zeichen = "&shy;"
            
                ElseIf (akt_zeichen = "Æ") Then
            
                    akt_zeichen = "&reg;"
            
                ElseIf (akt_zeichen = "Ø") Then
            
                    akt_zeichen = "&macr;"
            
                ElseIf (akt_zeichen = "∞") Then
            
                    akt_zeichen = "&deg;"
            
                ElseIf (akt_zeichen = "±") Then
            
                    akt_zeichen = "&plusmn;"
            
                ElseIf (akt_zeichen = "≤") Then
            
                    akt_zeichen = "&sup2;"
            
                ElseIf (akt_zeichen = "≥") Then
            
                    akt_zeichen = "&sup3;"
            
                ElseIf (akt_zeichen = "¥") Then
            
                    akt_zeichen = "&acute;"
            
                ElseIf (akt_zeichen = "µ") Then
            
                    akt_zeichen = "&micro;"
            
                ElseIf (akt_zeichen = "∂") Then
            
                    akt_zeichen = "&para;"
            
                ElseIf (akt_zeichen = "∑") Then
            
                    akt_zeichen = "&middot;"
            
                ElseIf (akt_zeichen = "∏") Then
            
                    akt_zeichen = "&cedil;"
            
                ElseIf (akt_zeichen = "π") Then
            
                    akt_zeichen = "&sup1;"
            
                ElseIf (akt_zeichen = "∫") Then
            
                    akt_zeichen = "&ordm;"
            
                ElseIf (akt_zeichen = "ª") Then
            
                    akt_zeichen = "&raquo;"
            
                ElseIf (akt_zeichen = "º") Then
            
                    akt_zeichen = "&frac14;"
            
                ElseIf (akt_zeichen = "Ω") Then
            
                    akt_zeichen = "&frac12;"
            
                ElseIf (akt_zeichen = "æ") Then
            
                    akt_zeichen = "&frac34;"
            
                ElseIf (akt_zeichen = "ø") Then
            
                    akt_zeichen = "&iquest;"
            
                ElseIf (akt_zeichen = "¿") Then
            
                    akt_zeichen = "&Agrave;"
            
                ElseIf (akt_zeichen = "¡") Then
            
                    akt_zeichen = "&Aacute;"
            
                ElseIf (akt_zeichen = "¬") Then
            
                    akt_zeichen = "&Acirc;"
            
                ElseIf (akt_zeichen = "√") Then
            
                    akt_zeichen = "&Atilde;"
            
                ElseIf (akt_zeichen = "ƒ") Then
            
                    akt_zeichen = "&Auml;"
            
                ElseIf (akt_zeichen = "≈") Then
            
                    akt_zeichen = "&Aring;"
            
                ElseIf (akt_zeichen = "∆") Then
            
                    akt_zeichen = "&AElig;"
            
                ElseIf (akt_zeichen = "«") Then
            
                    akt_zeichen = "&Ccedil;"
            
                ElseIf (akt_zeichen = "»") Then
            
                    akt_zeichen = "&Egrave;"
            
                ElseIf (akt_zeichen = "…") Then
            
                    akt_zeichen = "&Eacute;"
            
                ElseIf (akt_zeichen = " ") Then
            
                    akt_zeichen = "&Ecirc;"
            
                ElseIf (akt_zeichen = "À") Then
            
                    akt_zeichen = "&Euml;"
            
                ElseIf (akt_zeichen = "Ã") Then
            
                    akt_zeichen = "&Igrave;"
            
                ElseIf (akt_zeichen = "Õ") Then
            
                    akt_zeichen = "&Iacute;"
            
                ElseIf (akt_zeichen = "Œ") Then
            
                    akt_zeichen = "&Icirc;"
            
                ElseIf (akt_zeichen = "œ") Then
            
                    akt_zeichen = "&Iuml;"
            
                ElseIf (akt_zeichen = "–") Then
            
                    akt_zeichen = "&ETH;"
            
                ElseIf (akt_zeichen = "—") Then
            
                    akt_zeichen = "&Ntilde;"
            
                ElseIf (akt_zeichen = "“") Then
            
                    akt_zeichen = "&Ograve;"
            
                ElseIf (akt_zeichen = "”") Then
            
                    akt_zeichen = "&Oacute;"
            
                ElseIf (akt_zeichen = "‘") Then
            
                    akt_zeichen = "&Ocirc;"
            
                ElseIf (akt_zeichen = "’") Then
            
                    akt_zeichen = "&Otilde;"
            
                ElseIf (akt_zeichen = "÷") Then
            
                    akt_zeichen = "&Ouml;"
            
                ElseIf (akt_zeichen = "◊") Then
            
                    akt_zeichen = "&times;"
            
                ElseIf (akt_zeichen = "ÿ") Then
            
                    akt_zeichen = "&Oslash;"
            
                ElseIf (akt_zeichen = "Ÿ") Then
            
                    akt_zeichen = "&Ugrave;"
            
                ElseIf (akt_zeichen = "⁄") Then
            
                    akt_zeichen = "&Uacute;"
            
                ElseIf (akt_zeichen = "€") Then
            
                    akt_zeichen = "&Ucirc;"
            
                ElseIf (akt_zeichen = "‹") Then
            
                    akt_zeichen = "&Uuml;"
            
                ElseIf (akt_zeichen = "›") Then
            
                    akt_zeichen = "&Yacute;"
            
                ElseIf (akt_zeichen = "ﬁ") Then
            
                    akt_zeichen = "&THORN;"
            
                ElseIf (akt_zeichen = "ﬂ") Then
            
                    akt_zeichen = "&szlig;"
            
                ElseIf (akt_zeichen = "‡") Then
            
                    akt_zeichen = "&agrave;"
            
                ElseIf (akt_zeichen = "·") Then
            
                    akt_zeichen = "&aacute;"
            
                ElseIf (akt_zeichen = "‚") Then
            
                    akt_zeichen = "&acirc;"
            
                ElseIf (akt_zeichen = "„") Then
            
                    akt_zeichen = "&atilde;"
            
                ElseIf (akt_zeichen = "‰") Then
            
                    akt_zeichen = "&auml;"
            
                ElseIf (akt_zeichen = "Â") Then
            
                    akt_zeichen = "&aring;"
            
                ElseIf (akt_zeichen = "Ê") Then
            
                    akt_zeichen = "&aelig;"
            
                ElseIf (akt_zeichen = "Á") Then
            
                    akt_zeichen = "&ccedil;"
            
                ElseIf (akt_zeichen = "Ë") Then
            
                    akt_zeichen = "&egrave;"
            
                ElseIf (akt_zeichen = "È") Then
            
                    akt_zeichen = "&eacute;"
            
                ElseIf (akt_zeichen = "Í") Then
            
                    akt_zeichen = "&ecirc;"
            
                ElseIf (akt_zeichen = "Î") Then
            
                    akt_zeichen = "&euml;"
            
                ElseIf (akt_zeichen = "Ï") Then
            
                    akt_zeichen = "&igrave;"
            
                ElseIf (akt_zeichen = "Ì") Then
            
                    akt_zeichen = "&iacute;"
            
                ElseIf (akt_zeichen = "Ó") Then
            
                    akt_zeichen = "&icirc;"
            
                ElseIf (akt_zeichen = "Ô") Then
            
                    akt_zeichen = "&iuml;"
            
                ElseIf (akt_zeichen = "") Then
            
                    akt_zeichen = "&eth;"
            
                ElseIf (akt_zeichen = "Ò") Then
            
                    akt_zeichen = "&ntilde;"
            
                ElseIf (akt_zeichen = "Ú") Then
            
                    akt_zeichen = "&ograve;"
            
                ElseIf (akt_zeichen = "Û") Then
            
                    akt_zeichen = "&oacute;"
            
                ElseIf (akt_zeichen = "Ù") Then
            
                    akt_zeichen = "&ocirc;"
            
                ElseIf (akt_zeichen = "ı") Then
            
                    akt_zeichen = "&otilde;"
            
                ElseIf (akt_zeichen = "ˆ") Then
            
                    akt_zeichen = "&ouml;"
            
                ElseIf (akt_zeichen = "˜") Then
            
                    akt_zeichen = "&divide;"
            
                ElseIf (akt_zeichen = "¯") Then
            
                    akt_zeichen = "&oslash;"
            
                ElseIf (akt_zeichen = "˘") Then
            
                    akt_zeichen = "&ugrave;"
            
                ElseIf (akt_zeichen = "˙") Then
            
                    akt_zeichen = "&uacute;"
            
                ElseIf (akt_zeichen = "˚") Then
            
                    akt_zeichen = "&ucirc;"
            
                ElseIf (akt_zeichen = "¸") Then
            
                    akt_zeichen = "&uuml;"
            
                ElseIf (akt_zeichen = "˝") Then
            
                    akt_zeichen = "&yacute;"
            
                ElseIf (akt_zeichen = "˛") Then
            
                    akt_zeichen = "&thorn;"
            
                ElseIf (akt_zeichen = "ˇ") Then
            
                    akt_zeichen = "&yuml;"
            
                ElseIf (akt_zeichen = "ë") Then
            
                    akt_zeichen = "&lsquo;"
            
                ElseIf (akt_zeichen = "í") Then
            
                    akt_zeichen = "&rsquo;"
    
                End If
            
            End If

            html_ergebnis_string = html_ergebnis_string & akt_zeichen
            
            '
            ' Leseprozess eine Stelle weiterstellen
            '
            akt_index = akt_index + 1
            
        Wend

    End If

EndFunktion:

    On Error Resume Next

    DoEvents

    quoteHtmlCharacter = html_ergebnis_string

    Exit Function

errQuoteHtmlCharacter:

    quoteHtmlCharacter = "Fehler: errQuoteHtmlCharacter: " & Error & " " & Erl

    Resume EndFunktion

End Function

'################################################################################
'
Private Function replaceUmlaute(ByVal pString As String) As String

On Error GoTo errReplaceUmlaute

Dim akt_position     As Integer
Dim akt_zeichen      As String
Dim str_fkt_ergebnis As String
    
    akt_position = 1
    
    '
    ' Die While-Schleife laeuft ueber die Laenge des Eingabestrings.
    '
    While (akt_position <= Len(pString))
    
        akt_zeichen = Mid(pString, akt_position, 1)
        
        If (akt_zeichen = "‰") Then

            akt_zeichen = "ae"

        ElseIf (akt_zeichen = "ƒ") Then

            akt_zeichen = "Ae"

        ElseIf (akt_zeichen = "ˆ") Then

            akt_zeichen = "oe"

        ElseIf (akt_zeichen = "÷") Then

            akt_zeichen = "Oe"

        ElseIf (akt_zeichen = "¸") Then

            akt_zeichen = "ue"

        ElseIf (akt_zeichen = "‹") Then

            akt_zeichen = "Ue"

        ElseIf (akt_zeichen = "ﬂ") Then

            akt_zeichen = "ss"

        ElseIf (akt_zeichen = "È") Then

            akt_zeichen = "e"

        ElseIf (akt_zeichen = "Ë") Then

            akt_zeichen = "e"

        ElseIf (akt_zeichen = "Ä") Then

            akt_zeichen = "eur"

        ElseIf (akt_zeichen = "$") Then

            akt_zeichen = "dollar"

        End If
        
        
        'If (akt_zeichen<> LEER_STRING) Then
                
            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
        
        'End If
        
        akt_position = akt_position + 1
    
    Wend
    
EndFunktion:

    On Error Resume Next

    '
    ' DoEvents aufrufen
    '
    DoEvents

    '
    ' Funktionsergebnis setzen
    '
    replaceUmlaute = str_fkt_ergebnis

    '
    ' Funktion verlassen
    '
    Exit Function

errReplaceUmlaute:

    'Call debug.Print("Fehler: errReplaceUmlaute: " & Err & " " & Error & " " & Erl)

    Resume EndFunktion

End Function

'##############################################################################
'
Public Function getPosWortende(ByVal pString As String, ByVal pPositionStart As Integer) As Integer

On Local Error Resume Next

Dim akt_position                    As Integer ' aktuelle Leseposition
Dim anzahl_zeichen                  As Integer ' Laenge des Eingabestrings
Dim zeichen_wortbestandteil         As String
Dim knz_weiterer_schleifendurchlauf As Boolean

    zeichen_wortbestandteil = "enirstaudhgolcmfbkVvwz1paeSDA0E2RBGueMIPKF9UNW3L78oeH4T5CZJy6ssxjOueYXqQae_"

    knz_weiterer_schleifendurchlauf = True
    
    anzahl_zeichen = Len(pString)
    
    akt_position = pPositionStart

    While ((akt_position <= anzahl_zeichen) And (knz_weiterer_schleifendurchlauf))

        If (InStr(zeichen_wortbestandteil, Mid(pString, akt_position, 1)) > 0) Then
            
            akt_position = akt_position + 1
            
        Else
        
            knz_weiterer_schleifendurchlauf = False
        
        End If
    
    Wend
    
    If (akt_position = pPositionStart) Then
        
        getPosWortende = akt_position
    
    Else
        
        getPosWortende = akt_position - 1
    
    End If

End Function

'##############################################################################
'
Public Function getPosWortAnfang(ByVal pString As String, ByVal pPositionStart As Integer) As Integer

On Local Error Resume Next

Dim akt_position                     As Integer ' aktuelle Leseposition
Dim anzahl_zeichen                   As Integer ' Laenge des Eingabestrings
Dim zeichen_wortbestandteil          As String
Dim knz_weiterer_schleifendurchlauf  As Boolean

    zeichen_wortbestandteil = "enirstaudhgolcmfbkVvwz1paeSDA0E2RBGueMIPKF9UNW3L78oeH4T5CZJy6ssxjOueYXqQae_"

    knz_weiterer_schleifendurchlauf = True
    
    anzahl_zeichen = Len(pString)
    
    akt_position = pPositionStart

    While ((akt_position > 0) And (knz_weiterer_schleifendurchlauf))

        If (InStr(zeichen_wortbestandteil, Mid(pString, akt_position, 1)) > 0) Then
            
            akt_position = akt_position - 1
            
        Else
        
            knz_weiterer_schleifendurchlauf = False
        
        End If
    
    Wend
    
    If (akt_position = pPositionStart) Then
        
        getPosWortAnfang = akt_position
        
    Else
    
        getPosWortAnfang = akt_position + 1
    
    End If

End Function

'################################################################################
'
' Gibt dem Aufrufer eine x-malige Stringverkettung von "pEingabe" zurueck.
' Werden die Anzahl der Wiederholungen negativ angegeben, gibt es einen Leerstring.
'
' ? FkString.getStringXmal( "[a-z]", 4 ) = [a-z][a-z][a-z][a-z]
'
' ? FkString.getStringXmal( "A-",  1 ) = "A-"
' ? FkString.getStringXmal( "A-",  3 ) = "A-A-A-"
' ? FkString.getStringXmal( "A-", -3 ) = ""
' ? FkString.getStringXmal(   "", 10 ) = ""
'
' PARAMETER: pEingabe       = der zu wiederholende String
' PARAMETER: pAnzahlWiederholungen = die Anzahl der Wiederholungen
'
' RETURN : Eine x-malige Stringverkettung der Eingabezeichenfolge
'
Private Function getStringXmal(pEingabe As String, pAnzahlWiederholungen As Integer) As String

Dim str_fkt_ergebnis As String
Dim zaehler_schleife As Integer

    '
    ' Pruefung: "pEingabe" ungleich Leerstring ?
    '
    ' Ist die Eingabe ein Leerstring, bekommt der Aufrufer gleich
    ' einen Leerstring zurueck. Es muss keine Schleife ausgefuehrt
    ' werden.
    '
    If (pEingabe <> LEER_STRING) Then

        '
        ' Der Zaehler startet bei 0
        '
        zaehler_schleife = 0
        
        '
        ' In einer While-Schleife, wird der Ergebnisstring aufgebaut.
        ' Die While-Schleife laeuft solange, bis der Zaehler gleich
        ' der geforderten Wiederholungsanzahl ist.
        '
        While (zaehler_schleife < pAnzahlWiederholungen)
    
            str_fkt_ergebnis = str_fkt_ergebnis & pEingabe
            
            zaehler_schleife = zaehler_schleife + 1
    
        Wend
        
    End If
    
    '
    ' Am Funktionsende wird der aufgebaute Ergebnisstring zurueckgegeben.
    '
    getStringXmal = str_fkt_ergebnis

End Function

'################################################################################
'
' Entfernt von der Eingabe die nachlaufenden Leerzeichen
'
' trimTail( "ABC   " ) = "ABC"
' trimTail( "      " ) = ""
' trimTail( ""       ) = ""
'
' PARAMETER: pString der zu trimmende String
'
' RETURN den getrimmten String
'
Private Function trimTail(pString As String) As String

Dim akt_position As Long

    '
    ' Vorgaberueckgabe ist ein Leerstring
    '
    trimTail = ""
    
    '
    ' Pruefung: "pString" ungleich Leerstring ?
    '
    ' Ist "pString" gleich einem Leerstring bekommt der
    ' Aufrufer einen Leerstring zurueck.
    '
    ' Ist "pString" ungleich einem Leerstring, wird die
    ' Suchschleife gestartet.
    '
    If (pString <> LEER_STRING) Then
    
        '
        ' Die Leseposition fuer die Suchschleife beginnt am Stringende.
        ' Die Variable "akt_position" ist gleich der Eingabelaenge.
        '
        akt_position = Len(pString)
        
        '
        ' Die Suchschleife wird ausgefuehrt solange die Leseposition
        ' groesser als 0 ist (Es noch zu pruefende Zeichen gibt.)
        '
        While (akt_position > 0)
        
            '
            ' Es wird das Zeichen an der Leseposition darauf geprueft,
            ' ob es anders als ein Leerzeichen ist. Wenn dem so ist,
            ' ist, wird das Funktionsergebnis auf den Teilstring gesetzt,
            ' welcher sich von 0 bis zur aktuellen Leseposition ergibt.
            ' Die Funktion wird anschliessend verlassen.
            '
            If (Mid(pString, akt_position, 1) <> " ") Then
            
                trimTail = Left(pString, akt_position)
                
                Exit Function
                
            End If
        
            '
            ' War das Zeichen an der Leseposition noch ein Leerzeichen,
            ' wird die Leseposition verringert und der naechste
            ' Schleifendurchlauf gestartet.
            '
            akt_position = akt_position - 1
            
        Wend
                
    End If
    
End Function

'################################################################################
'
' Ersetzt alle Vorkommen des Suchstrings in pQuellstring mit der Zeichenfolge pStringNeu.
' Dabei wird die Gross/Kleinschreibung bei der Suche des Suchtextes ignoriert.
'
' ? replaceIgnoreCase( "ABC..XYZ.def..xyz.GHI..xYz.jkl.MNO", ".XyZ.", "--" ) = ABC.--def.--GHI.--jkl.MNO
' ? replaceIgnoreCase( "ABC..XYZ.def..xyz.GHI..xYz.jkl.MNO", ".XyZ.", ""   ) = ABC.def.GHI.jkl.MNO
'
' PARAMETER: pQuellString  = der zu durchsuchende String
' PARAMETER: pSuchString   = der Suchstring
' PARAMETER: pStringNeu    = der Ersatzstring fuer den Suchstring ( Leerzeichen = Eliminierung Suchstring )
'
' RETURN  : ein String, in welchem die Suchzeichenfolge durch die Ersatzzeichenfolge ersetzt wurde
'
Private Function replaceIgnoreCase(pQuellString As String, pSuchString As String, pStringNeu As String) As String

Dim such_string_ucase     As String  ' Suchtext in Grossbuchstaben
Dim quell_string_ucase    As String  ' durchsuchter Text in Grossbuchstaben
Dim ergebnis_string       As String  ' text fuer die Rueckgabe
Dim position_such_string  As Long    ' die aktuell gefundene Startposition des Suchstrings
Dim position_such_prozess As Long    ' die aktuelle AB-Position fuer die Suche in quell_string_ucase
Dim zaehler               As Long    ' ein Zaehler zur Vermeidung von Endlossschleifen
    
    '
    ' Variableninitialisierung
    ' Der zu durchsuchende String und der Suchstring werden auf Grossbuchstaben
    ' konvertiert. Die beiden Positionsvariablen bekommen den Startwert 1.
    '
    such_string_ucase = UCase(pSuchString)
    
    quell_string_ucase = UCase(pQuellString)
    
    position_such_prozess = 1
    
    '
    ' Initiale Position suchen
    '
    position_such_string = InStr(position_such_prozess, quell_string_ucase, such_string_ucase)
    
    If (position_such_string <= 0) Then
    
        replaceIgnoreCase = pQuellString
        
        Exit Function
        
    End If
    
    '
    ' Die Suchschleife wird solange durchlaufen wie
    ' ... die Position des Suchstrings noch groesser als 0 ist
    ' ... der Zaehler noch kleiner 500 ist (Vermeidung Endlossschleife)
    '
    While ((position_such_string > 0) And (zaehler < 500))
        
        '
        ' Pruefung: Suchstring gefunden ?
        ' Das ist der Fall, wenn die Positon einen Wert groesser 0 hat.
        '
        If (position_such_string > 0) Then
        
            '
            ' Ergebnisstring aufbauen
            ' Aus dem Parameter-Quellstring wird von der letzten Position bis zur aktuellen
            ' Position des Suchstrings die Zeichen kopiert. Anschliessend wird die neue
            ' Zeichenfolge aus "pStringNeu" dem Ergebnis hinzugefuegt.
            '
            ' Ist "pStringNeu" ein Leerstring, wird eben der Suchstring aus dem
            ' Quellstring entfernt (es gibt keinen Ersatzstring).
            '
            ergebnis_string = ergebnis_string & Mid(pQuellString, position_such_prozess, position_such_string - position_such_prozess)
            
            ergebnis_string = ergebnis_string & pStringNeu
            
            '
            ' Position Leseprozess setzen
            ' Die neue Startposition fuer den naechsten Suchvorgang beginnt ab der
            ' eben gefundenen Position des Suchstrings zuzueglich dessen Laenge.
            '
            position_such_prozess = position_such_string + Len(such_string_ucase)

        End If
        
        '
        ' Position Suchstring ermitteln
        ' Im Upper-Case-Quellstring wird der Upper-Case-Suchstring gesucht. Somit
        ' wird die Gross/Klein-Schreibung eliminiert. Die Position wird in der
        ' Variablen "position_such_string" gespeichert.
        '
        position_such_string = InStr(position_such_prozess, quell_string_ucase, such_string_ucase)
        
        '
        ' Zaehler erhoehen
        '
        zaehler = zaehler + 1

    Wend
    '
    ' Pruefung: wurden alle Zeichen der Eingabe behandelt ?
    '
    If (position_such_prozess < Len(pQuellString)) Then

        ergebnis_string = ergebnis_string & Mid(pQuellString, position_such_prozess)

    End If

    replaceIgnoreCase = ergebnis_string

End Function

'################################################################################
'
' ? getRemoveAbBis( "1234567890", 4, 8    ) = 12390
' ? getRemoveAbBis( "1234567890", 4, 18   ) = 123
' ? getRemoveAbBis( "1234567890", 4, -8   ) =
' ? getRemoveAbBis( "1234567890", 0, 8    ) = 90
'
Private Function getRemoveAbBis(pEingabe As String, pAbPosition As Long, pBisPosition As Long) As String

Dim str_fkt_ergebnis As String

    str_fkt_ergebnis = LEER_STRING
    
    If ((pAbPosition > 1) And (pAbPosition <= pBisPosition)) Then
        
        str_fkt_ergebnis = Left(pEingabe, pAbPosition - 1)
            
    End If

    If ((pBisPosition > 0) And (pAbPosition <= pBisPosition) And (pBisPosition < Len(pEingabe))) Then
        
        str_fkt_ergebnis = str_fkt_ergebnis & Right(pEingabe, Len(pEingabe) - pBisPosition)
         
    End If
    
    getRemoveAbBis = str_fkt_ergebnis

End Function

'################################################################################
'
' ? replaceSubstringAbBis( "1234567890", 4,  8 , "ABC" ) 123ABC90
' ? replaceSubstringAbBis( "1234567890", 4, 18 , "ABC" ) 123ABC
'
Private Function replaceSubstringAbBis(pEingabe As String, pAbPosition As Long, pBisPosition As Long, pReplaceWith As String) As String

Dim str_fkt_ergebnis As String

    str_fkt_ergebnis = LEER_STRING
    
    If ((pAbPosition > 1) And (pAbPosition <= pBisPosition)) Then
        
        str_fkt_ergebnis = Left(pEingabe, pAbPosition - 1)
            
    End If
    
    str_fkt_ergebnis = str_fkt_ergebnis & pReplaceWith

    If ((pBisPosition > 0) And (pAbPosition <= pBisPosition) And (pBisPosition < Len(pEingabe))) Then
        
        str_fkt_ergebnis = str_fkt_ergebnis & Right(pEingabe, Len(pEingabe) - pBisPosition)
         
    End If
     
    replaceSubstringAbBis = str_fkt_ergebnis

End Function

'################################################################################
'
' ? getStringAbBis( "ABC.DEF.GHI.JKL.MNO", -6, 8 )
'
Public Function getStringAbBis(pEingabe As String, ByVal pAbPosition As Long, ByVal pBisPosition As Long) As String

    '
    ' 1. Die Ab-Position muss grosser 0 sein
    ' 2. Die Ab-Position muss kleiner gleich der Laenge der Eingabe sein
    ' 3. Die Ab-Position muss kleiner gleich der Bis-Position sein
    '

Dim len_eingabe As Long
    
    getStringAbBis = LEER_STRING
    
    len_eingabe = Len(pEingabe)
    
    If (len_eingabe > 0) Then
    
        '
        ' Ist die AB-Position groesser als die Laenge der Eingabe,
        ' bleibt vom String nichts nach.
        '
        If (pAbPosition <= len_eingabe) Then
        
            If (pAbPosition <= 0) Then
        
                pAbPosition = 1
            
            End If
        
            If (pAbPosition <= pBisPosition) Then
        
                getStringAbBis = Mid(pEingabe, pAbPosition, (pBisPosition - pAbPosition) + 1)
            
            End If
        
        End If

    End If

End Function

'################################################################################
'
Private Function getBenutztesChr13(pString As String) As String
    
    If (InStr(1, pString, MY_CHR_13_10, vbBinaryCompare) > 0) Then
    
        getBenutztesChr13 = MY_CHR_13_10
    
    ElseIf (InStr(1, pString, Chr(13), vbBinaryCompare) > 0) Then
        
        getBenutztesChr13 = Chr(13)
        
    Else
    
        getBenutztesChr13 = MY_CHR_13_10
    
    End If

End Function

'################################################################################
'
' Gibt die Position des letzten Vorkommens der Suchzeichenfolge zurueck.
' Wenn die Suchzeichenfolge in der Eingabe nicht vorhanden ist, wird -1 zurueckgeliefert.
'
'                             1234567890123456789012
' ? getLetztePositionVorPos( "ABC.DEF.GHI.JKL.MNO", ".", 9 )  8
'
'
' PARAMETER: pEingabe     = der zu untersuchende String
' PARAMETER: pSuchString  = der zu suchende Trennstring
' PARAMETER: pEndPosition = die Position bis zu welcher gesucht werden soll
'
' RETURN  : die Position des letzten Vorkommens vom Suchstring, bzw. -1 wenn dieser nicht gefunden wurde
'
Private Function getLetztePositionVorPos(ByVal pEingabe As String, ByVal pSuchString As String, ByVal pEndPosition As Long) As Long

Dim akt_position As Long ' Speichert die aktuell gefundene Position

    '
    ' Das Funktionsergebnis wird auf -1 gesetzt
    '
    getLetztePositionVorPos = -1
    
    '
    ' Pruefung: Suchstring vorhanden ?
    '
    If ((pSuchString <> LEER_STRING) And (pEndPosition > 0)) Then
    
        '
        ' Initiale Position des Suchstrings bestimmen
        '
        akt_position = InStr(pEingabe, pSuchString)

        While (akt_position > 0)
        
            '
            ' Pruefung: End-Position vorhanden?
            '
            If (pEndPosition > 0) Then
            
                '
                ' Pruefung: Endposition ueberschritten ?
                '
                ' Liegt die aktuelle Fundstelle der Suchzeichenfolge vor der Endposition,
                ' wird das Funktionsergebnis auf diese Fundstelle gestellt.
                '
                ' Liegt die aktuelle Fundstelle nach der Endposition, wird die Funktion beendet
                '
                If (akt_position <= pEndPosition) Then
                    
                    getLetztePositionVorPos = akt_position
                
                Else
                    
                    Exit Function
                
                End If
            
            Else
                '
                ' Wurde keine Endposition uebergeben, wird das Funktionsergebnis auf
                ' die aktuelle Fundstelle gestellt.
                '
                getLetztePositionVorPos = akt_position
                 
            End If

            '
            ' Naechste Fundstelle suchen.
            '
            akt_position = InStr(akt_position + Len(pSuchString), pEingabe, pSuchString)

        Wend
        
    End If

End Function

'################################################################################
'
' Aus der Eingabe werden nur die in der Funktion hinterlegten gueltigen Zeichen uebernommen.
'
' ? getStringGueltigeZeichen( " ABC#DEF <GHI@JKL.MN> " ) = ABCDEFGHIJKLMN
'
' PARAMETER: pString        = der zu behandelnde Eingabestring
'
' RETURN : Einen String mit nur den Zeichen der gueltigen Zeichen
'
Public Function getStringGueltigeZeichen(pString As String) As String

Dim akt_position     As Integer ' aktuelle Leseposition der While-Schleife
Dim akt_zeichen      As String  ' aktuelle Zeichen in der While-Schleife
Dim str_my_cr        As String  ' Das in dieser Funktion verwendete Zeilenumbruchszeichen
Dim str_fkt_ergebnis As String  ' Ergebnisstring fuer die Rueckgabe
Dim gueltige_zeichen As String

    '
    ' Initialisierung des Strings mit den gueltigen Zeichen
    '
    gueltige_zeichen = " enirstaudhgolcmfbkVvwz1paeSDA0E2RBGueMIPKF9UNW3L78oeH4T5CZJy6xjOUeYXqQ‰ˆ¸ƒ÷‹_?!""¬ß$%&/()<>[]{}=*'/*-+:;,.#\/1234567890"

    m_toggle_mr_stringer_fkt = Not m_toggle_mr_stringer_fkt
  
    If (m_toggle_mr_stringer_fkt) Then
    
        str_my_cr = Chr(13)
    
    Else
    
        str_my_cr = MY_CHR_13_10
    
    End If
  
    akt_position = 1

    '
    ' While-Schleife ueber alle Zeichen der Eingabe
    '
    While (akt_position <= Len(pString))

        '
        ' Zeichen aus der Eingabe an der aktuellen Lesepositon lesen
        '
        akt_zeichen = Mid(pString, akt_position, 1)

        '
        ' Pruefung, ob das aktuelle Zeichen in der Menge der
        ' gueltigen Zeichen enthalten ist.
        '
        If (InStr(gueltige_zeichen, akt_zeichen) > 0) Then

            '
            ' Ist das Zeichen fuer die Aufnahme in das Ergebnis OK,
            ' wird das Zeichen dem Ergebnisstring hinzugefuegt.
            '
            ' Ist das Zeichen nicht in OK, wird das Zeichen ueberlesen
            '
            str_fkt_ergebnis = str_fkt_ergebnis & akt_zeichen
            
        ElseIf (akt_zeichen = vbCr) Then
        
            '
            ' Zeilenumbrueche werden uebernommen.
            '
            str_fkt_ergebnis = str_fkt_ergebnis & str_my_cr
        
        End If

        '
        ' Es wird die Leseposition der While-Schleife erhoeht und
        ' mit dem naechsten Schleifedurchlauf weitergemacht.
        '
        akt_position = akt_position + 1

    Wend

    '
    ' Der Aufrufer bekommt den Ergebnisstring zurueck.
    '
    getStringGueltigeZeichen = str_fkt_ergebnis

End Function

'################################################################################
'
' Ermittelt wie oft die Zeichenfolge pSuchString in der Zeichenfolge pEingabeString vorkommt.
'
'   ? fkString.getAnzahlVorkommen( " A   A   A  B  A ", "A" ) = 4
'   ? fkString.getAnzahlVorkommen( " A   A   A  B  A ", "B" ) = 1
'   ? fkString.getAnzahlVorkommen( " A   A   A  B  A ", " " ) = 12
'   ? fkString.getAnzahlVorkommen( " A   A   A  B  A ", "C" ) = 0   Nicht vorhandener Suchstring
'   ? fkString.getAnzahlVorkommen( " A   A   A  B  A ", ""  ) = 0   Suche nach Leerstring
'   ? fkString.getAnzahlVorkommen( "", "B"                  ) = 0   Suche in einem Leerstring
'   ? fkString.getAnzahlVorkommen( "", ""                   ) = 0   Suche nach Leerstring im Leerstring
'
'   ? fkString.getAnzahlVorkommen( "A A A A A A A A", "A A" ) = 4   Keine "geschachtelten" Vorkommen mitzaehlen
'
' PARAMETER: pEingabeString = der zu untersuchende String
' PARAMETER: pSuchString    = der zu suchende String
'
' RETURN  : Die Anzahl der Vorkommen von der Zeichenkette aus dem Parameter "pSuchString"
'
Private Function getAnzahlVorkommen(ByVal pEingabeString As String, ByVal pSuchString As String) As Long

Dim zaehler_vorkommen  As Long
Dim aktuelle_position  As Long
Dim laenge_such_string As Integer

'
' Einzig denkbarer Fehler ist, wenn der Zaehler fuer die aktuelle Position groesser
' als der Bereich Long wird. In einem solchen Fall wird zum Funktionsende
' verzweigt, welcher dann den aktuellen Zaehler zurueck gibt.
'
On Error GoTo endGetAnzahlVorkommen

    zaehler_vorkommen = 0
    
    '
    ' Die Laenge der Such-Zeichenfolge wird der aktuellen Startposition hinzugezaehlt.
    '
    laenge_such_string = Len(pSuchString)
    
    '
    ' Sollen die geschachtelten Vorkommen gezaehlt werden, muss
    ' der Suchprozess nur um 1 Zeichen weitergesetzt werden.
    '
    'laenge_such_string = 1
    '
    ' Bei einer Suche nach einem Leerstring wuerde es zu einer Endlosschleife kommen.
    ' Um das zu verhindern, darf die Schleife nur bei einem Suchstring mit mehr
    ' als 0 Zeichen gestartet werden.
    '
    If (laenge_such_string > 0) Then
        '
        ' Damit die Startposition fuer den ersten Aufruf der Funktion "InStr"
        ' wieder 1 ergibt, wird die Laenge des Suchstrings von 1 abgezogen.
        '
        ' (Das hinzuaddieren der Suchstringlaenge koennte auch seperat ausserhalb
        '  der Funktion "InStr" gemacht werden.)
        '
        aktuelle_position = 1 - laenge_such_string

        Do
            aktuelle_position = InStr(aktuelle_position + laenge_such_string, pEingabeString, pSuchString)

            '
            ' Solange der Suchstring noch gefunden werden kann, ist die
            ' Variable "aktuelle_position" groesser als 0.
            '
            ' Kann der Suchstring nicht mehr gefunden werden, wird die Schleife verlassen
            '
            If (aktuelle_position > 0) Then

                zaehler_vorkommen = zaehler_vorkommen + 1

            Else

                Exit Do

            End If
            
        Loop

    End If

endGetAnzahlVorkommen:

    getAnzahlVorkommen = zaehler_vorkommen

End Function



'##############################################################################
'
Public Function renameYoutube(pDateiName As String, pKnzReplaceLeerzeichen As Boolean) As String

Dim akt_position      As Integer
Dim akt_zeichen       As String
Dim datei_name_neu    As String
Dim datei_erweiterung As String
Dim zaehler_1         As Integer
Dim zaehler_2         As Integer
Dim temp_string       As String

    datei_name_neu = ""
    
    zaehler_1 = 0
    zaehler_2 = 0
    
    akt_position = 1
    
    '
    ' Die While-Schleife laeuft ueber die Laenge des Eingabestrings.
    '
    While (akt_position <= Len(pDateiName))
    
        akt_zeichen = Mid(pDateiName, akt_position, 1)
        
        '
        ' Auswertung des aktuellen Zeichens
        '
        ' Es werden alle Uni-Code-Zeichen ignoriert
        '
        ' Zeichen ohne Spezialbehandlung sind in der
        ' Konstanten "gueltige_zeichen" hinterlegt.
        '
        ' Deutsche Umlaute werden entsprechend konvertiert.
        ' Alle Klammern werden zu runden Klammern konvertiert.
        ' Das Apostrophzeichen wird auf ' hin konvertiert.
        ' Alle anderen Zeichen werden zu einem Unterstrich konvertiert.
        '
        'If ((AscW(akt_zeichen) And &HFFFF&) > &H7F&) Then

            'https://stackoverflow.com/questions/30712461/need-code-for-removing-all-unicode-characters-in-vb6
            'https://www.experts-exchange.com/questions/26026296/handling-Unicode-in-VB6-How-to-replace-convert-to-ansi-or-to-find-in-a-text-string.html
            
        '    akt_zeichen = ""

        If (InStr(GUELTIGE_ZEICHEN_DATEI_NAME, akt_zeichen) <= 0) Then
        
            If (akt_zeichen = "‰") Then

                akt_zeichen = "ae"

            ElseIf (akt_zeichen = "ƒ") Then

                akt_zeichen = "Ae"

            ElseIf (akt_zeichen = "ˆ") Then

                akt_zeichen = "oe"

            ElseIf (akt_zeichen = "÷") Then

                akt_zeichen = "Oe"

            ElseIf (akt_zeichen = "¸") Then

                akt_zeichen = "ue"

            ElseIf (akt_zeichen = "‹") Then

                akt_zeichen = "Ue"

            ElseIf (akt_zeichen = "ﬂ") Then

                akt_zeichen = "ss"
            
            ElseIf (akt_zeichen = "È") Then

                akt_zeichen = "e"
            
            ElseIf (akt_zeichen = "Ë") Then

                akt_zeichen = "e"

            ElseIf (akt_zeichen = "¥" Or akt_zeichen = "`" Or akt_zeichen = "ë" Or akt_zeichen = "í") Then

                akt_zeichen = "'"

            ElseIf (akt_zeichen = "[") Then

                akt_zeichen = "("
            
            ElseIf (akt_zeichen = "]") Then

                akt_zeichen = ")"
            
            ElseIf (akt_zeichen = "<") Then

                akt_zeichen = "("

            ElseIf (akt_zeichen = ">") Then

                akt_zeichen = ")"
            
            ElseIf (akt_zeichen = "'") Then

                akt_zeichen = "'"
            
            ElseIf (akt_zeichen = """") Then

                akt_zeichen = "'"
            
            ElseIf ((akt_zeichen = "©") Or (akt_zeichen = "ˇ")) Then

                akt_zeichen = ""
                
            ElseIf (akt_zeichen = "" Or akt_zeichen = "~") Then
            
                akt_zeichen = "-"
       
            'ElseIf (InStr("~", akt_zeichen) <= 0) Then
            '    akt_zeichen = "-"
            
            End If
            
        End If

        If (akt_zeichen <> LEER_STRING) Then
        
            If (akt_zeichen = " ") Then
            
                If (zaehler_1 = 0) Then
        
                    datei_name_neu = datei_name_neu & akt_zeichen
                    
                    zaehler_1 = 1
                
                End If
            
            ElseIf (akt_zeichen = "!") Then
            
                If (zaehler_2 = 0) Then
        
                    datei_name_neu = datei_name_neu & akt_zeichen
                    
                    zaehler_2 = 1
                
                End If
            
            Else
        
                datei_name_neu = datei_name_neu & akt_zeichen
                
                zaehler_1 = 0
                zaehler_2 = 0
                
            End If

        End If
        
        akt_position = akt_position + 1
    
    Wend
    
    '
    ' Dateierweiterung ermitteln
    '
    datei_erweiterung = getErweiterung(datei_name_neu)
    
    '
    ' Pruefung: Dateierweiterung vorhanden ?
    '
    If (datei_erweiterung <> LEER_STRING) Then
    
        '
        ' Dateierweiterungen in Kleinbuchstaben wandeln
        '
        datei_name_neu = Replace(datei_name_neu, datei_erweiterung, LCase(datei_erweiterung))
        
        datei_erweiterung = LCase(datei_erweiterung)
        
        '
        ' Ungueltige Erweiterungen werden konvertiert. (mpeg zu mpg)
        '
        If (datei_erweiterung = ".xvid") Then
        
            datei_erweiterung = ".avi"
            
            datei_name_neu = Replace(datei_name_neu, ".xvid", datei_erweiterung)
        
        ElseIf (datei_erweiterung = ".divx") Then
            
            datei_erweiterung = ".avi"
        
            datei_name_neu = Replace(datei_name_neu, ".divx", datei_erweiterung)
        
        ElseIf (datei_erweiterung = ".mpeg") Then
            
            datei_erweiterung = ".mpg"
            
            datei_name_neu = Replace(datei_name_neu, ".mpeg", datei_erweiterung)
        
        End If
        
    End If
    
    If (strEnthaelt(datei_name_neu, "NDR", False) = 1) Then
        
        datei_name_neu = Replace(datei_name_neu, "_ NDR Doku", "")
        datei_name_neu = Replace(datei_name_neu, "_ NDR (", "(")
    
    End If
    
    If (strEnthaelt(datei_name_neu, "_ die nordstory", False) = 1) Then
    
        datei_name_neu = "Die Nordstory - " & datei_name_neu
        
        datei_name_neu = Replace(datei_name_neu, "_ die nordstory", "")
        
    Else
    
        datei_name_neu = Replace(datei_name_neu, "die nordstory", "Die Nordstory")
    
    End If
    
    If (strEnthaelt(datei_name_neu, "die nordreportage", False) = 1) Then
    
        datei_name_neu = "Die Nordreportage - " & datei_name_neu
        
        datei_name_neu = Replace(datei_name_neu, "die nordreportage", "")
    
    End If
    
    If (strEnthaelt(datei_name_neu, "NZZ Format ", False) = 1) Then
    
        datei_name_neu = "NZZ Format - " & datei_name_neu
        
        datei_name_neu = Replace(datei_name_neu, "(NZZ Format ", "(")
    
    End If
    
    If (strEnthaelt(datei_name_neu, "Sportclub", False) = 1) Then
    
        datei_name_neu = Replace(datei_name_neu, "Sportclub", "")
        
        datei_name_neu = "Sportclub - " & datei_name_neu
    
    End If
    
    If (strEnthaelt(datei_name_neu, "Super Easy Russian ", False) = 1) Then
        
        datei_name_neu = Replace(datei_name_neu, "Super Easy Russian ", "")
    
        datei_name_neu = "Super Easy Russian  " & datei_name_neu
    
    ElseIf (strEnthaelt(datei_name_neu, "Easy Russian ", False) = 1) Then
        
        datei_name_neu = Replace(datei_name_neu, "Super Easy Russian ", "")
    
        datei_name_neu = "Easy Russian  " & datei_name_neu
    
    End If
    
    
    datei_name_neu = Replace(datei_name_neu, "Kuzgesagt", "Kurzgesagt")
    datei_name_neu = Replace(datei_name_neu, "techmoan", "Techmoan")
    
    '
    ' Youtube Frameraten und Codec-Informationen entfernen
    '
    datei_name_neu = Replace(datei_name_neu, "-128kbit_", "")
    datei_name_neu = Replace(datei_name_neu, "-192kbit_", "")
    datei_name_neu = Replace(datei_name_neu, "_192kbit_", "")
    
    datei_name_neu = Replace(datei_name_neu, "_60fps_H264AAC", "")
    datei_name_neu = Replace(datei_name_neu, "_50fps_H264AAC", "")
    datei_name_neu = Replace(datei_name_neu, "_30fps_H264AAC", "")
    datei_name_neu = Replace(datei_name_neu, "_29fps_H264AAC", "")
    datei_name_neu = Replace(datei_name_neu, "_27fps_H264AAC", "")
    datei_name_neu = Replace(datei_name_neu, "_25fps_H264AAC", "")
    datei_name_neu = Replace(datei_name_neu, "_24fps_H264AAC", "")

    datei_name_neu = Replace(datei_name_neu, "_60fps_", "")
    datei_name_neu = Replace(datei_name_neu, "_50fps_", "")
    datei_name_neu = Replace(datei_name_neu, "_30fps_", "")
    datei_name_neu = Replace(datei_name_neu, "_29fps_", "")
    datei_name_neu = Replace(datei_name_neu, "_27fps_", "")
    datei_name_neu = Replace(datei_name_neu, "_25fps_", "")
    datei_name_neu = Replace(datei_name_neu, "_24fps_", "")

'(720p_30fps_192kbit_AAC)
    '
    ' Aufloesungsinformationen entfernen
    '
    datei_name_neu = Replace(datei_name_neu, "(356p)", "##MARKIERUNG_1##")
    datei_name_neu = Replace(datei_name_neu, "(360p)", "##MARKIERUNG_1##")
    datei_name_neu = Replace(datei_name_neu, "(356p)", "##MARKIERUNG_1##")
    datei_name_neu = Replace(datei_name_neu, "(480p)", "##MARKIERUNG_1##")
    datei_name_neu = Replace(datei_name_neu, "(470p)", "##MARKIERUNG_1##")
    datei_name_neu = Replace(datei_name_neu, "(720p)", "##MARKIERUNG_1##")
    datei_name_neu = Replace(datei_name_neu, "(1080p)", "##MARKIERUNG_1##")
    datei_name_neu = Replace(datei_name_neu, "(1044p)", "##MARKIERUNG_1##")
    datei_name_neu = Replace(datei_name_neu, "(2160p)", "##MARKIERUNG_1##")
    
    datei_name_neu = Replace(datei_name_neu, "_)_", ") ")
    
    datei_name_neu = Replace(datei_name_neu, "--", " - ")
    datei_name_neu = Replace(datei_name_neu, " -", " - ")
    datei_name_neu = Replace(datei_name_neu, " -  ", " - ")
    datei_name_neu = Replace(datei_name_neu, " -  ", " - ")
    
    datei_name_neu = Replace(datei_name_neu, "_ _", " - ")
    datei_name_neu = Replace(datei_name_neu, " _ ", " - ")
    datei_name_neu = Replace(datei_name_neu, "_ñ_", " - ")
    datei_name_neu = Replace(datei_name_neu, ".-.", " - ")
    
    datei_name_neu = Replace(datei_name_neu, "_ ", " - ")
    datei_name_neu = Replace(datei_name_neu, " - _", " - ")
    datei_name_neu = Replace(datei_name_neu, " -- ", " - ")
    datei_name_neu = Replace(datei_name_neu, " ... ", " - ")
    datei_name_neu = Replace(datei_name_neu, " - - ", " - ")
    datei_name_neu = Replace(datei_name_neu, "_) - ", ") - ")
            
    datei_name_neu = Replace(datei_name_neu, " - ##MARKIERUNG_1##", "")
    datei_name_neu = Replace(datei_name_neu, "##MARKIERUNG_1##", "")
    
    datei_name_neu = Replace(datei_name_neu, "_! ", "! ")
    datei_name_neu = Replace(datei_name_neu, "! - ", "! ")
        
    datei_name_neu = Replace(datei_name_neu, " & ", " and ")
    
    datei_name_neu = Replace(datei_name_neu, ".webxl.h264.", ".")
    datei_name_neu = Replace(datei_name_neu, "(Doku,", "(")
    
    datei_name_neu = Replace(datei_name_neu, "HD Documentary", "")
    datei_name_neu = Replace(datei_name_neu, "(NEU)", "")
    
    datei_name_neu = Replace(datei_name_neu, "How-to", "How-To")
    datei_name_neu = Replace(datei_name_neu, "techmoan - ", "Techmoan - ")
    datei_name_neu = Replace(datei_name_neu, "nessiejudge - ", "Nessiejudge - ")
    datei_name_neu = Replace(datei_name_neu, "Vsauce - VSouce - ", "Vsauce - ")
    
    datei_name_neu = Replace(datei_name_neu, ".webxl.h264.", ".")
    
    '
    ' Entfernen bzw. Konvertieren von Partteil-Angaben
    '
    zaehler_1 = 1
    
    While (zaehler_1 < 10)
    
        zaehler_2 = 1
    
        While (zaehler_2 < zaehler_1)
    
            datei_name_neu = Replace(datei_name_neu, zaehler_2 & "of" & zaehler_1, "0" & zaehler_2)
            
            datei_name_neu = Replace(datei_name_neu, zaehler_2 & " of " & zaehler_1, "0" & zaehler_2)
            
            datei_name_neu = Replace(datei_name_neu, zaehler_2 & "Of" & zaehler_1, "0" & zaehler_2)
            
            datei_name_neu = Replace(datei_name_neu, zaehler_2 & " Of " & zaehler_1, "0" & zaehler_2)
            
            datei_name_neu = Replace(datei_name_neu, zaehler_2 & " - " & zaehler_1, "0" & zaehler_2)
            
            datei_name_neu = Replace(datei_name_neu, "(" & zaehler_2 & " von " & zaehler_1 & ")", "0" & zaehler_2)
            
            datei_name_neu = Replace(datei_name_neu, "(" & zaehler_2 & "_von_" & zaehler_1 & ")", "0" & zaehler_2)
            
            datei_name_neu = Replace(datei_name_neu, zaehler_2 & " von " & zaehler_1, "0" & zaehler_2)
            
            zaehler_2 = zaehler_2 + 1
        
        Wend
        
        datei_name_neu = Replace(datei_name_neu, "(Pt. " & zaehler_1 & ")", " - Part 0" & zaehler_1 & " ")
        datei_name_neu = Replace(datei_name_neu, "(Pt." & zaehler_1 & ")", " - Part 0" & zaehler_1 & " ")
        datei_name_neu = Replace(datei_name_neu, "(Pt " & zaehler_1 & ")", " - Part 0" & zaehler_1 & " ")
        datei_name_neu = Replace(datei_name_neu, "(Part " & zaehler_1 & ")", " - Part 0" & zaehler_1 & " ")
        datei_name_neu = Replace(datei_name_neu, "(Ep " & zaehler_1 & ")", " - Part 0" & zaehler_1 & " ")
        datei_name_neu = Replace(datei_name_neu, "(Episode " & zaehler_1 & ")", " - Episode 0" & zaehler_1 & " ")
        datei_name_neu = Replace(datei_name_neu, "(Series Part " & zaehler_1 & ")", " - Part 0" & zaehler_1 & " ")
    
        datei_name_neu = Replace(datei_name_neu, " Pt." & zaehler_1 & datei_erweiterung, " - Part 0" & zaehler_1 & datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, " Pt" & zaehler_1 & datei_erweiterung, " - Part 0" & zaehler_1 & datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, " Pt " & zaehler_1 & datei_erweiterung, " - Part 0" & zaehler_1 & datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, " Part " & zaehler_1 & datei_erweiterung, " - Part 0" & zaehler_1 & datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, " Ep." & zaehler_1 & datei_erweiterung, " - Part 0" & zaehler_1 & datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, " Ep" & zaehler_1 & datei_erweiterung, " - Part 0" & zaehler_1 & datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, " Ep " & zaehler_1 & datei_erweiterung, " - Part 0" & zaehler_1 & datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, " Episode " & zaehler_1 & datei_erweiterung, " - Episode 0" & zaehler_1 & datei_erweiterung)
    
        temp_string = " _#!,-" ' Punkt nicht aufnehmen, wegen Dateierweiterung (Fehler ist, dass der Punkt kann sonst entfernt koennte)
        
        akt_position = 1
        
        While (akt_position <= Len(temp_string))
    
            akt_zeichen = Mid(temp_string, akt_position, 1)
            
            datei_name_neu = Replace(datei_name_neu, " Pt." & zaehler_1 & akt_zeichen, " - Part 0" & zaehler_1 & " ")
            datei_name_neu = Replace(datei_name_neu, " Pt" & zaehler_1 & akt_zeichen, " - Part 0" & zaehler_1 & " ")
            datei_name_neu = Replace(datei_name_neu, " Pt " & zaehler_1 & akt_zeichen, " - Part 0" & zaehler_1 & " ")
            datei_name_neu = Replace(datei_name_neu, " Part " & zaehler_1 & akt_zeichen, " - Part 0" & zaehler_1 & " ")
            datei_name_neu = Replace(datei_name_neu, " Ep." & zaehler_1 & akt_zeichen, " - Part 0" & zaehler_1 & " ")
            datei_name_neu = Replace(datei_name_neu, " Ep" & zaehler_1 & akt_zeichen, " - Part 0" & zaehler_1 & " ")
            datei_name_neu = Replace(datei_name_neu, " Ep " & zaehler_1 & akt_zeichen, " - Part 0" & zaehler_1 & " ")
            datei_name_neu = Replace(datei_name_neu, " Episode " & zaehler_1 & akt_zeichen, " - Episode 0" & zaehler_1 & " ")
    
            akt_position = akt_position + 1
    
        Wend
    
        zaehler_1 = zaehler_1 + 1
        
    Wend
    
    
    zaehler_1 = 1990
    
    While (zaehler_1 < 2030)
        
        datei_name_neu = Replace(datei_name_neu, "(Doku, " & zaehler_1 & ")", "(" & zaehler_1 & ")")
        
        datei_name_neu = Replace(datei_name_neu, "(Doku,_" & zaehler_1 & ")", "(" & zaehler_1 & ")")
        
        datei_name_neu = Replace(datei_name_neu, "(Doku " & zaehler_1 & ")", "(" & zaehler_1 & ")")
        datei_name_neu = Replace(datei_name_neu, " Doku " & zaehler_1 & " ", "(" & zaehler_1 & ")")
        
        datei_name_neu = Replace(datei_name_neu, " Doku " & zaehler_1 & "]", " " & zaehler_1 & "]")
        datei_name_neu = Replace(datei_name_neu, " Doku " & zaehler_1 & ")", " " & zaehler_1 & ")")
        
        datei_name_neu = Replace(datei_name_neu, "(From " & zaehler_1 & ")", "(" & zaehler_1 & ")")
        
        datei_name_neu = Replace(datei_name_neu, "( " & zaehler_1 & " )", "(" & zaehler_1 & ")")
        
        zaehler_1 = zaehler_1 + 1
        
    Wend
    
    datei_name_neu = Replace(datei_name_neu, " Doku, HD ", "")
    
    datei_name_neu = Replace(datei_name_neu, "(Neu)", "")
    datei_name_neu = Replace(datei_name_neu, "(NEU)", "")
    datei_name_neu = Replace(datei_name_neu, "(Neue)", "")
    datei_name_neu = Replace(datei_name_neu, "(NEUE)", "")
    
    datei_name_neu = Replace(datei_name_neu, "()", "")

    If (datei_erweiterung <> LEER_STRING) Then
    
        temp_string = Right(datei_name_neu, Len(datei_erweiterung))
        
        If (temp_string <> datei_erweiterung) Then
            
            datei_name_neu = Left(datei_name_neu, Len(datei_name_neu) - Len(datei_erweiterung))
            
            datei_name_neu = datei_name_neu + datei_erweiterung
            
        End If
    
        zaehler_1 = 0
        
        While (zaehler_1 < 3)
            
            datei_name_neu = Replace(datei_name_neu, "  " & datei_erweiterung, datei_erweiterung)
            datei_name_neu = Replace(datei_name_neu, " " & datei_erweiterung, datei_erweiterung)
        
            zaehler_1 = zaehler_1 + 1
            
        Wend
        
        datei_name_neu = Replace(datei_name_neu, "Computerphile" & datei_erweiterung, datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, "!)" & datei_erweiterung, ")" & datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, "_)" & datei_erweiterung, ")" & datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, " )" & datei_erweiterung, ")" & datei_erweiterung)

        datei_name_neu = Replace(datei_name_neu, "(HQ)" & datei_erweiterung, datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, "(HD)" & datei_erweiterung, datei_erweiterung)
        
        datei_name_neu = Replace(datei_name_neu, " HD" & datei_erweiterung, datei_erweiterung)
        datei_name_neu = Replace(datei_name_neu, "_HD" & datei_erweiterung, datei_erweiterung)
        
        datei_name_neu = Replace(datei_name_neu, "(100% DIY)", "")
        datei_name_neu = Replace(datei_name_neu, "(amazing DIY)", "")
        
        datei_name_neu = Replace(datei_name_neu, " - " & datei_erweiterung, datei_erweiterung)
        
        temp_string = "#!,.-_ "
        
        akt_position = 1
        
        While (akt_position <= Len(temp_string))
    
            akt_zeichen = Mid(temp_string, akt_position, 1)
            
            datei_name_neu = Replace(datei_name_neu, " " & akt_zeichen & "  " & datei_erweiterung, datei_erweiterung)
            datei_name_neu = Replace(datei_name_neu, " " & akt_zeichen & " " & datei_erweiterung, datei_erweiterung)
            
            datei_name_neu = Replace(datei_name_neu, akt_zeichen & datei_erweiterung, datei_erweiterung)
            datei_name_neu = Replace(datei_name_neu, akt_zeichen & datei_erweiterung, datei_erweiterung)
            datei_name_neu = Replace(datei_name_neu, akt_zeichen & datei_erweiterung, datei_erweiterung)
            datei_name_neu = Replace(datei_name_neu, akt_zeichen & datei_erweiterung, datei_erweiterung)
    
            akt_position = akt_position + 1
    
        Wend

        datei_name_neu = Replace(datei_name_neu, " - " & datei_erweiterung, datei_erweiterung)

        zaehler_1 = 0
        
        While (zaehler_1 < 3)
            
            datei_name_neu = Replace(datei_name_neu, "  " & datei_erweiterung, datei_erweiterung)
            datei_name_neu = Replace(datei_name_neu, " " & datei_erweiterung, datei_erweiterung)
        
            zaehler_1 = zaehler_1 + 1
            
        Wend
        
    End If
    
    datei_name_neu = Replace(datei_name_neu, " - - ", " - ")
    
    datei_name_neu = Replace(datei_name_neu, " -  ", " - ")
    datei_name_neu = Replace(datei_name_neu, " -  ", " - ")
    
    datei_name_neu = Replace(datei_name_neu, "  - ", " - ")
    datei_name_neu = Replace(datei_name_neu, "  - ", " - ")
    
    datei_name_neu = Replace(datei_name_neu, " - - ", " - ")
    
    datei_name_neu = Replace(datei_name_neu, "   ", "  ")
    
    datei_name_neu = Replace(datei_name_neu, " Ft ", " ft ")
    datei_name_neu = Replace(datei_name_neu, " FT ", " ft ")
    datei_name_neu = Replace(datei_name_neu, " ft. ", " ft ")
    datei_name_neu = Replace(datei_name_neu, " Ft. ", " ft ")
    datei_name_neu = Replace(datei_name_neu, " FT. ", " ft ")
    
    
    If (pKnzReplaceLeerzeichen) Then
    
    
        datei_name_neu = Replace(datei_name_neu, " - ", "__")
        
        datei_name_neu = Replace(datei_name_neu, " ", "_")
        
        datei_name_neu = Replace(datei_name_neu, "___", "__")
        
        'datei_name_neu = renameWorteEnglish(datei_name_neu)

    End If
    
    renameYoutube = datei_name_neu

End Function

'################################################################################
'
' Gibt dem Aufrufer 1 zurueck, wenn der uebergebene Suchstring in der zu
' durchsuchenden Zeichenfolge enthalten ist.
'
' ? strEnthaelt( "ABCDE", "AB" ) = 1 oder KNZ_TRUE
' ? strEnthaelt( "ABCDE", "XZ" ) = 0 oder KNZ_FALSE
'
' PARAMETER: pZuDurchsuchenderString = der zu untersuchende String
' PARAMETER: pSuchString    = der zu suchende String
' PARAMETER: pKnzCaseSensitive = TRUE = Gross/Kleinschreibung wird beachtet, FALSE = ignoriere Gross/Kleinschreibung
'
' RETURN : 1 bei Vorkommen des Suchstrings, 0 wenn der Suchstring nicht enthalten ist
'
Private Function strEnthaelt(pZuDurchsuchenderString As String, pSuchString As String, Optional ByVal pKnzCaseSensitive As Boolean = True) As Integer

    strEnthaelt = 0

    If (pKnzCaseSensitive) Then

        If (InStr(pZuDurchsuchenderString, pSuchString) > 0) Then

            strEnthaelt = 1

        End If

    Else

        If (InStr(LCase(pZuDurchsuchenderString), LCase(pSuchString)) > 0) Then

            strEnthaelt = 1

        End If

    End If

End Function

'##############################################################################
'
Public Function renameDatei(pString As String) As String

Dim zaehler_unterstrich   As Integer
Dim datei_name_neu        As String ' Ergebnisstring fuer die Rueckgabe
Dim datei_erweiterung     As String
Dim akt_zeichen           As String
Dim str_unterstrich       As String  ' Vermeidung von fuehrenden Unterstrichen
Dim akt_position          As Integer ' aktuelle Leseposition der While-Schleife
Dim pos_letzter_punkt     As Integer
Dim akt_chr_wert          As Integer
Dim eingabe_string        As String

    str_unterstrich = "!"

    eingabe_string = Trim(pString)

    eingabe_string = Replace(eingabe_string, " - ", "__")

    akt_position = 1

    pos_letzter_punkt = -1

    zaehler_unterstrich = 0
    
    '
    ' Die While-Schleife laeuft ueber die Laenge des Eingabestrings.
    '
    While (akt_position <= Len(eingabe_string))
    
        akt_zeichen = Mid(eingabe_string, akt_position, 1)
        
        '
        ' Auswertung des aktuellen Zeichens
        '
        ' Zeichen ohne Spezialbehandlung sind in der
        ' Konstanten "gueltige_zeichen" hinterlegt.
        '
        ' Deutsche Umlaute werden entsprechend konvertiert.
        ' Alle Klammern werden zu runden Klammern konvertiert.
        ' Das Apostrophzeichen wird auf ' hin konvertiert.
        ' Alle anderen Zeichen werden zu einem Unterstrich konvertiert.
        '
        If (InStr(GUELTIGE_ZEICHEN_DATEI_NAME, akt_zeichen) <= 0) Then
        
            If (akt_zeichen = ".") Then
                
                akt_zeichen = "_"

                pos_letzter_punkt = akt_position

            ElseIf (akt_zeichen = "‰") Then

                akt_zeichen = "ae"

            ElseIf (akt_zeichen = "ƒ") Then

                akt_zeichen = "Ae"

            ElseIf (akt_zeichen = "ˆ") Then

                akt_zeichen = "oe"

            ElseIf (akt_zeichen = "÷") Then

                akt_zeichen = "Oe"

            ElseIf (akt_zeichen = "¸") Then

                akt_zeichen = "ue"

            ElseIf (akt_zeichen = "‹") Then

                akt_zeichen = "Ue"

            ElseIf (akt_zeichen = "ﬂ") Then

                akt_zeichen = "ss"

            ElseIf (akt_zeichen = "È") Then

                akt_zeichen = "e"
            
            ElseIf (akt_zeichen = "Ë") Then

                akt_zeichen = "e"

            ElseIf (akt_zeichen = "¥" Or akt_zeichen = "`" Or akt_zeichen = "ë") Then

                akt_zeichen = "'"

            ElseIf (akt_zeichen = "'") Then

                akt_zeichen = "'"
            
            ElseIf (akt_zeichen = """") Then

                akt_zeichen = "'"

            ElseIf (akt_zeichen = "[") Then

                akt_zeichen = "("

            ElseIf (akt_zeichen = "]") Then

                akt_zeichen = ")"

            ElseIf (akt_zeichen = "{") Then

                akt_zeichen = "("

            ElseIf (akt_zeichen = "}") Then

                akt_zeichen = ")"

            ElseIf (akt_zeichen = "<") Then

                akt_zeichen = "("

            ElseIf (akt_zeichen = ">") Then

                akt_zeichen = ")"

            Else

                akt_zeichen = "_"
                
            End If
        
        End If
        
        '
        ' Pruefung: aktuelles Zeichen anders als Unterstrich?
        ' Alle Zeichen anders als ein Unterstrich bewirken eine uebernahme des
        ' Zeichens in den Ergebnisstring.
        '
        ' Der Zaehler fuer die Unterstriche wird auf 0 gesetzt.
        '
        ' Da jetzt ein gueltiges Zeichen in das Ergebnis aufgenommen wurde, wird
        ' in der Variablen "str_unterstrich" das Zeichen fuer eben den Unterstrich
        ' hinterlegt. Dieses bewirkt eine "Freischaltung" der Unterstrichaufnahme.
        '
        If (akt_zeichen <> "_") Then
        
            datei_name_neu = datei_name_neu & akt_zeichen
            
            str_unterstrich = "_"
            
            zaehler_unterstrich = 0
        '
        ' Pruefung: aktuelles Zeichen ist Unterstrich?
        ' Um fuehrende Unterstriche zu vermeiden wird durch die Variable "str_unterstrich"
        ' in den ersten Laeufen ein anderes Zeichen hinterlegt. Somit kann die
        ' Abfrage nicht greifen. Wird ein gueltiges Startzeichen gefunden, wird
        ' gleichzeitig in der Variablen "str_unterstrich" ein Unterstrich
        ' hinterlegt und schaltet somit diese Abfrage erst frei.
        '
        ' Pruefung: Zaehler fuer Unterstrich kleiner 2 ?
        ' Es duerfen maximal 2 Unterstriche hintereinander stehen. Dieses wird
        ' durch eine Zaehlvariable fuer hinzugefuegte Unterstriche gemacht. Fuer
        ' jeden hinzugefuegten Unterstrich wird der Zaehler erhoeht.
        '
        ElseIf (akt_zeichen = str_unterstrich) Then
        
            If (zaehler_unterstrich < 2) Then
            
                datei_name_neu = datei_name_neu & akt_zeichen
                
                zaehler_unterstrich = zaehler_unterstrich + 1
                
            End If

        End If
        
        akt_position = akt_position + 1
    
    Wend
    '
    ' Es wird eine Dateierweiterung angenommen, wenn die Positon des letzten
    ' Punktes nicht weiter als 6 Stellen vor dem Ende des Eingabestrings lag.
    ' Die Position des letzten Punktes darf aber erst ab dem 4ten Zeichen liegen.
    '
    ' Die Erweiterung selber wird aus dem Eingabestring genommen, da die Position
    ' des letzten Punktes aus der Eingabe gespeichert wurde und zweitens nicht
    ' jedes Zeichen der Eingabe auch ein Zeichen in der Ausgabe nach sich zieht.
    '
    If ((pos_letzter_punkt > 4) And (pos_letzter_punkt > Len(eingabe_string) - 6)) Then
        '
        ' Die Dateierweiterung wird aus der Eingabezeichenfolge gelesen und in
        ' Kleinbuchstaben konvertiert. Der Punkt wird nicht mit aufgenommen.
        '
        datei_erweiterung = LCase(Mid(eingabe_string, pos_letzter_punkt + 1))
        '
        ' Der Ausgabestring wird bis zur Laenge der Erweiterung (bis zur Punktposition)
        ' gelesen.
        '
        datei_name_neu = Left(datei_name_neu, (Len(datei_name_neu) - Len(datei_erweiterung)) - 1) & "_##MARKIERUNG_2##" & "."
        '
        ' Falsche Erweiterungen werden nach AVI konvertiert, aus MPEG wird MPG.
        ' Die Dateierweiterung soll aus 3 Zeichen und aus Kleinbuchstaben bestehen.
        '
        If (datei_erweiterung = "xvid") Then
        
            datei_erweiterung = "avi"
        
        ElseIf (datei_erweiterung = "divx") Then
            
            datei_erweiterung = "avi"
        
        ElseIf (datei_erweiterung = "mpeg") Then
            
            datei_erweiterung = "mpg"
        
        ElseIf (datei_erweiterung = "jpeg") Then
            
            datei_erweiterung = "jpg"
        
        End If
        
        '
        ' Die Erweiterung wird dem Ausgabestring hinzugefuegt.
        '
        datei_name_neu = datei_name_neu & datei_erweiterung
        '
        ' Unterstriche vor der Erweiterung werden geloescht.
        '
        datei_erweiterung = "." & datei_erweiterung
        
    End If
    
    datei_name_neu = renameWorteEnglish(datei_name_neu)
    
    datei_name_neu = Replace(datei_name_neu, "_REPACK", "")
    datei_name_neu = Replace(datei_name_neu, "_XviD", "")
    datei_name_neu = Replace(datei_name_neu, "_DivX", "")
    datei_name_neu = Replace(datei_name_neu, "_GERMAN", "")
    datei_name_neu = Replace(datei_name_neu, "_HDTVRiP", "")
    datei_name_neu = Replace(datei_name_neu, "_HDRiP", "")
    datei_name_neu = Replace(datei_name_neu, "_DVDRiP", "")
    datei_name_neu = Replace(datei_name_neu, "_VHSRiP", "")
    datei_name_neu = Replace(datei_name_neu, "_SatRip", "")
    datei_name_neu = Replace(datei_name_neu, "_x264", "")
    datei_name_neu = Replace(datei_name_neu, "_H264", "")
    datei_name_neu = Replace(datei_name_neu, "_HDTV", "")
    datei_name_neu = Replace(datei_name_neu, "_H264", "")
    datei_name_neu = Replace(datei_name_neu, "_BluRay", "")
    datei_name_neu = Replace(datei_name_neu, "_DVDR", "")
    datei_name_neu = Replace(datei_name_neu, "_BDRiP", "")
    datei_name_neu = Replace(datei_name_neu, "_TVP", "")
    datei_name_neu = Replace(datei_name_neu, "_DVD9", "")
    datei_name_neu = Replace(datei_name_neu, "_DVD5", "")
    datei_name_neu = Replace(datei_name_neu, "_DVD1", "")
    datei_name_neu = Replace(datei_name_neu, "_DVD2", "")
    datei_name_neu = Replace(datei_name_neu, "_DVD4", "")
    datei_name_neu = Replace(datei_name_neu, "_DVD5", "")
    datei_name_neu = Replace(datei_name_neu, "_AC3", "")
    datei_name_neu = Replace(datei_name_neu, "_2160p", "")
    datei_name_neu = Replace(datei_name_neu, "_1080p", "")
    datei_name_neu = Replace(datei_name_neu, "_1080i", "")
    datei_name_neu = Replace(datei_name_neu, "_720p", "")
    datei_name_neu = Replace(datei_name_neu, "_480p", "")
    datei_name_neu = Replace(datei_name_neu, "_720i", "")
    
    datei_name_neu = Replace(datei_name_neu, "_AC3LD", "")
    datei_name_neu = Replace(datei_name_neu, "_Avi", ".avi")

    datei_name_neu = Replace(datei_name_neu, "Wissenschaft__Technik_Und_Innovation_Xenius", "Xenius")
    datei_name_neu = Replace(datei_name_neu, "Wissenschaft__Wissen_kompakt_Xenius", "Xenius")
    datei_name_neu = Replace(datei_name_neu, "Wissenschaft__Gesundheit_Und_Medizin_Xenius", "Xenius")
    datei_name_neu = Replace(datei_name_neu, "Wissenschaft__Umwelt_Und_Natur_Xenius", "Xenius")
    
    If (datei_erweiterung <> LEER_STRING) Then
    
        akt_position = 0
        
        While (akt_position < 10)
            
            datei_name_neu = Replace(datei_name_neu, "_" & datei_erweiterung, datei_erweiterung)
        
            akt_position = akt_position + 1
            
        Wend
        
        If (Len(datei_name_neu) > 15) Then
        
            Dim start_pos_ext As Integer
            
            start_pos_ext = Len(datei_name_neu) - 4
            
            akt_position = start_pos_ext
            
            Dim knz_muster_ok As Boolean
            Dim akt_zeichen_asc As Integer
            Dim act_pos_zaehler As Integer
            
            knz_muster_ok = True
            
            While ((act_pos_zaehler < 10) And (knz_muster_ok))
                
                '
                ' Aktuelles Zeichen
                ' Aus der Eingabe wird fuer die aktuelle Leseposition der
                ' ASCI-Wert des dort stehenden Zeichens ermittelt.
                '
                akt_zeichen = Asc(Mid(datei_name_neu, akt_position, 1))
    
                '
                ' Pruefung des aktuellen Zeichens
                '
                If ((akt_zeichen >= 48) And (akt_zeichen <= 57)) Then
    
                    ' Aktuelles Zeichen ist eine Ziffer = OK
                    
                    akt_position = akt_position - 1
                
                    act_pos_zaehler = act_pos_zaehler + 1
                    
                Else
                
                    knz_muster_ok = False
                    
                End If

            Wend
            
            If (knz_muster_ok) Then
            
                If (Mid(datei_name_neu, akt_position, 1) = "_") Then
                
                    'Debug.Print Left(datei_name_neu, start_pos_ext)
                    'Debug.Print Mid(datei_name_neu, akt_position, 11)
                
                    datei_name_neu = Replace(datei_name_neu, Mid(datei_name_neu, akt_position, 11), "")
                
                End If
            
            End If
            
        End If
        
    End If
        
    renameDatei = datei_name_neu

End Function

Private Function renameWorteEnglish(pString As String) As String

Dim datei_name_neu   As String
Dim csv_werte        As String  ' Die zu durchlaufenen Werte
Dim csv_trennzeichen As String  ' Das Trennzeichen, mit wlechem die Werte getrennt sind
Dim csv_feld_zaehler As Integer ' Zaehler fuer die Elemente
Dim position_start   As Integer ' Die aktuelle Startposition des CSV-Wertes
Dim position_ende    As Integer ' Die aktuelle Endpositon des CSV-Wertes
Dim akt_teil_string  As String  ' Der aktuelle Teilstring aus den CSV-Werten

    datei_name_neu = "##MARKIERUNG_1##_" & pString & "_##MARKIERUNG_2##"
    
    datei_name_neu = Replace(datei_name_neu, "_a_", "_A_")

    csv_trennzeichen = csv_trennzeichen & ""
    csv_trennzeichen = csv_trennzeichen & "an, as, at, be, by, do, go, if, im, in,"
    csv_trennzeichen = csv_trennzeichen & "is, it, me, mr, my, no, of, ol, on, or,"
    csv_trennzeichen = csv_trennzeichen & "so, to, up, us, vs, we, add, afd, aim, air,"
    csv_trennzeichen = csv_trennzeichen & "aka, all, amp, and, app, aps, are, art, bad, ban,"
    csv_trennzeichen = csv_trennzeichen & "bar, bbs, bgm, bic, big, bin, bit, blu, box, boy,"
    csv_trennzeichen = csv_trennzeichen & "but, buy, cam, can, car, cd3, cds, ced, cgp, cld,"
    csv_trennzeichen = csv_trennzeichen & "cpu, cup, dat, day, dbx, dcc, did, die, diy, duo,"
    csv_trennzeichen = csv_trennzeichen & "dvd, egg, ein, elk, end, eye, fan, far, few,"
    csv_trennzeichen = csv_trennzeichen & "fix, for, fun, gas, get, gp3, guy, had, hap, has,"
    csv_trennzeichen = csv_trennzeichen & "her, how, hrs, ibm, ice, ich, ide, ink, ion,"
    csv_trennzeichen = csv_trennzeichen & "its, jbl, jet, jfk, jot, jun, key, kit, lcd, led,"
    csv_trennzeichen = csv_trennzeichen & "let, lit, mac, mad, mai, map, may, men, met, min,"
    csv_trennzeichen = csv_trennzeichen & "mix, mkv, mp3, mqa, mqs, new, not, now, nuc, obd,"
    csv_trennzeichen = csv_trennzeichen & "odd, off, old, one, osx, our, out, own, pac, pad,"
    csv_trennzeichen = csv_trennzeichen & "pbs, pci, pen, per, pet, pre, pro, pwp, ram, rca,"
    csv_trennzeichen = csv_trennzeichen & "rcd, rcs, red, rgb, rig, rom, rpm, sea,"
    csv_trennzeichen = csv_trennzeichen & "see, set, sky, spy, stc, the, tnt, too,"
    csv_trennzeichen = csv_trennzeichen & "top, trs, tvs, two, uhd, und, usa, usb, use, "
    csv_trennzeichen = csv_trennzeichen & "vcd, vcr, vhd, vhs, vid, war, was, wax, way, web,"
    csv_trennzeichen = csv_trennzeichen & "wes, who, why, win, won, yet, you, ages, also, azur,"
    csv_trennzeichen = csv_trennzeichen & "back, baru, bear, been, beer, bell, best, bits, bloc, blue,"
    csv_trennzeichen = csv_trennzeichen & "bond, bone, boom, burn, call, came, cant, card, cars, cart,"
    csv_trennzeichen = csv_trennzeichen & "case, cash, cats, cave, cctv, cd32, chip, city, coca, cola,"
    csv_trennzeichen = csv_trennzeichen & "cold, cool, cope, copy, cost, cows, cube, cult, dark, data,"
    csv_trennzeichen = csv_trennzeichen & "dead, deck, deep, demo, disc, disk, dock, does, dont, down,"
    csv_trennzeichen = csv_trennzeichen & "east, easy, echo, edge, emmy, envy, even, ever, fall, feat,"
    csv_trennzeichen = csv_trennzeichen & "feet, film, find, fire, fish, flaw, flip, food, ford, form,"
    csv_trennzeichen = csv_trennzeichen & "four, free, from, full, game, gave, gear, gfci, glow, gold,"
    csv_trennzeichen = csv_trennzeichen & "good, grey, guys, hack, hair, half, hard, have, hdmi, heat,"
    csv_trennzeichen = csv_trennzeichen & "held, hell, help, here, hifi, high, holy, home, hour, hype,"
    csv_trennzeichen = csv_trennzeichen & "inch, info, into, ipad, jarl, jets, jobs, june, just, kill,"
    csv_trennzeichen = csv_trennzeichen & "know, korg, last, late, lead, leds, left, lego, lens, less,"
    csv_trennzeichen = csv_trennzeichen & "life, like, link, list, lock, logo, long, look, lost, love,"
    csv_trennzeichen = csv_trennzeichen & "lynx, made, mail, main, make, male, many, mars, meat, meet,"
    csv_trennzeichen = csv_trennzeichen & "mega, mein, memo, menu, meta, mini, mkii, moon, more, mova,"
    csv_trennzeichen = csv_trennzeichen & "move, mpeg, much, must, name, nasa, navy, neat, need, nest,"
    csv_trennzeichen = csv_trennzeichen & "next, nova, omni, once, only, onto, open, oslo, over, pace,"
    csv_trennzeichen = csv_trennzeichen & "pack, page, paid, part, pays, pete, plan, play, plug, plus,"
    csv_trennzeichen = csv_trennzeichen & "pong, poor, prix, pure, race, rare, real, reel, rica, rich,"
    csv_trennzeichen = csv_trennzeichen & "ring, rise, road, role, room, ruvi, safe, said, sail, sale,"
    csv_trennzeichen = csv_trennzeichen & "same, scan, sega, send, shot, show, side, sink, some, song,"
    csv_trennzeichen = csv_trennzeichen & "sony, spot, star, stop, such, take, talk, tank, tape, teac,"
    csv_trennzeichen = csv_trennzeichen & "tear, tech, tell, test, than, that, them, they, this, time,"
    csv_trennzeichen = csv_trennzeichen & "tiny, tips, took, tour, tube, turn, twin, used, user, uses,"
    csv_trennzeichen = csv_trennzeichen & "ussr, wall, want, wars, wave, ways, week, well, wena, went,"
    csv_trennzeichen = csv_trennzeichen & "were, west, what, when, wifi, will, wird, wire, with, woes,"
    csv_trennzeichen = csv_trennzeichen & "word, work, wrap, xbmc, yang, year, yoga, your, zero, zink,"
    csv_trennzeichen = csv_trennzeichen & "zone, about, acorn, added, after, again, agree, album, alpex, amiga,"
    csv_trennzeichen = csv_trennzeichen & "apple, atari, audio, aukey, award, bacon, badge, bafta, balls, based,"
    csv_trennzeichen = csv_trennzeichen & "beach, beans, beige, black, blues, board, bonus, books, brits, broad,"
    csv_trennzeichen = csv_trennzeichen & "build, built, cable, candy, carts, cases, casio, catch, cause, caves,"
    csv_trennzeichen = csv_trennzeichen & "chaos, cheap, chips, cited, clash, class, claus, clean, click, clock,"
    csv_trennzeichen = csv_trennzeichen & "cloud, cohen, color, compo, costa, could, craft, crazy, cream, cross,"
    csv_trennzeichen = csv_trennzeichen & "cubes, curta, danny, death, decay, delay, denon, depth, didnt, discs,"
    csv_trennzeichen = csv_trennzeichen & "disks, drama, dream, drive, droid, dunia, early, earth, email, emile,"
    csv_trennzeichen = csv_trennzeichen & "enden, enemy, enjoy, epson, exist, extra, fails, false, fatal, fault,"
    csv_trennzeichen = csv_trennzeichen & "ferry, fiber, field, fifty, films, final, first, fixes, flash, focus,"
    csv_trennzeichen = csv_trennzeichen & "foods, found, fries, fully, gamer, games, giant, girls, given, goats,"
    csv_trennzeichen = csv_trennzeichen & "going, grail, grand, great, grill, gtech, guide, hafen, hands, happy,"
    csv_trennzeichen = csv_trennzeichen & "heads, heard, hello, heute, hipac, horse, house, human, idiot, india,"
    csv_trennzeichen = csv_trennzeichen & "intel, jahre, japan, kampf, knell, lapse, laser, lemon, light, lines,"
    csv_trennzeichen = csv_trennzeichen & "links, linux, lives, lixie, looks, loops, magic, makan, maker, maths,"
    csv_trennzeichen = csv_trennzeichen & "maybe, means, mechs, media, meets, merry, metre, micro, model, mount,"
    csv_trennzeichen = csv_trennzeichen & "movie, multi, music, names, naval, never, nigel, night, nixie, nomad,"
    csv_trennzeichen = csv_trennzeichen & "notes, ocean, often, other, owned, parts, party, phone, phono, photo,"
    csv_trennzeichen = csv_trennzeichen & "pitch, place, plane, plays, point, power, press, price, proto, prynt,"
    csv_trennzeichen = csv_trennzeichen & "pulse, quick, radio, range, raven, ready, recap, relic, retro, ricoh,"
    csv_trennzeichen = csv_trennzeichen & "rings, rival, robot, route, royal, ruark, rules, santa, sanyo, saves,"
    csv_trennzeichen = csv_trennzeichen & "scale, scary, scope, sense, setup, seven, shack, shape, sheep, sheer,"
    csv_trennzeichen = csv_trennzeichen & "shiny, shoot, short, shots, shown, signs, since, sized, slang, smart,"
    csv_trennzeichen = csv_trennzeichen & "smith, solar, solve, sound, space, spain, speed, spray, squad, stand,"
    csv_trennzeichen = csv_trennzeichen & "state, steve, stick, store, story, strap, strip, style, sucks, super,"
    csv_trennzeichen = csv_trennzeichen & "swear, sweet, swiss, tahun, tails, tandy, tapes, taste, taxes, tease,"
    csv_trennzeichen = csv_trennzeichen & "texas, their, there, these, theta, thing, think, third, those, three,"
    csv_trennzeichen = csv_trennzeichen & "throw, timer, titan, touch, track, trade, trash, trekz, trs80, trump,"
    csv_trennzeichen = csv_trennzeichen & "trust, truth, tuner, twist, types, ultra, under, until, using, valid,"
    csv_trennzeichen = csv_trennzeichen & "value, vault, video, views, vinyl, voice, wales, wants, watch, water,"
    csv_trennzeichen = csv_trennzeichen & "waves, webxl, weird, whale, wheel, where, which, while, white, whose,"
    csv_trennzeichen = csv_trennzeichen & "wings, wired, words, works, world, worse, worst, worth, would, wrist,"
    csv_trennzeichen = csv_trennzeichen & "wrong, yaqin, years, yours, accent, active, advent, advice, alloys, almost,"
    csv_trennzeichen = csv_trennzeichen & "always, amazon, analog, andrew, anthem, arcade, arctic, around, autism, awards,"
    csv_trennzeichen = csv_trennzeichen & "backup, banana, banned, became, become, behind, berlin, better, boardi, bonnin,"
    csv_trennzeichen = csv_trennzeichen & "boogie, bottle, brause, brexit, bright, brings, broken, budget, burger, burner,"
    csv_trennzeichen = csv_trennzeichen & "button, buyers, buying, called, camera, carbon, centre, chance, change, checks,"
    csv_trennzeichen = csv_trennzeichen & "chimps, cinema, circle, clever, clicky, clocks, clones, closed, coding, coffee,"
    csv_trennzeichen = csv_trennzeichen & "colour, comics, common, cooler, copper, corner, corona, cosmos, coulda, covers,"
    csv_trennzeichen = csv_trennzeichen & "create, crimea, custom, danger, deeper, degree, dengon, design, digits, disney,"
    csv_trennzeichen = csv_trennzeichen & "diving, dragon, dramas, driver, dubbed, edible, edison, edited, editor, effect,"
    csv_trennzeichen = csv_trennzeichen & "eminem, empire, energy, engine, enough, ensure, errors, failed, faking, faults,"
    csv_trennzeichen = csv_trennzeichen & "fidget, filmic, finale, fixing, flawed, flight, floppy, flying, follow, fooled,"
    csv_trennzeichen = csv_trennzeichen & "forced, forgot, format, former, fortwo, french, fuller, future, gadget, galaxy,"
    csv_trennzeichen = csv_trennzeichen & "gaming, garish, gently, german, ghetto, ghosts, giants, globes, golden, guides,"
    csv_trennzeichen = csv_trennzeichen & "havent, heaven, helium, hidden, hybrid, images, impact, import, indoor, infuse,"
    csv_trennzeichen = csv_trennzeichen & "inkjet, innovv, insane, inside, insult, iphone, issues, italia, jammer, killed,"
    csv_trennzeichen = csv_trennzeichen & "killer, latter, launch, laying, length, lenovo, lesson, lights, longer, looked,"
    csv_trennzeichen = csv_trennzeichen & "lovely, lyrics, making, market, master, megacd, memory, meters, micros, missed,"
    csv_trennzeichen = csv_trennzeichen & "mobile, morgen, movies, moving, needed, newbie, newest, normal, norway, number,"
    csv_trennzeichen = csv_trennzeichen & "ocasse, octane, opened, optane, orange, parlor, people, pepper, period, philip,"
    csv_trennzeichen = csv_trennzeichen & "photos, pixels, planes, player, pocket, potato, preamp, public, purple, ramsay,"
    csv_trennzeichen = csv_trennzeichen & "really, recent, recipe, record, region, relate, remain, remote, repair, report,"
    csv_trennzeichen = csv_trennzeichen & "reverb, review, ripped, riscpc, robert, roland, rollie, runner, runway, russia,"
    csv_trennzeichen = csv_trennzeichen & "sahara, sample, saving, saying, scenes, screen, screws, search, searle, secret,"
    csv_trennzeichen = csv_trennzeichen & "secure, selfie, senseo, series, served, server, silent, silver, simple, single,"
    csv_trennzeichen = csv_trennzeichen & "sodium, solder, soviet, spacex, sparks, speech, spider, spirit, spooky, starts,"
    csv_trennzeichen = csv_trennzeichen & "stereo, sticks, stones, strand, studio, supply, survey, system, tablet, talisa,"
    csv_trennzeichen = csv_trennzeichen & "things, though, titles, topics, toyota, trains, tricks, triple, trying, tunnel,"
    csv_trennzeichen = csv_trennzeichen & "unique, unused, update, urchin, vacuum, valves, varied, vector, victor, video8,"
    csv_trennzeichen = csv_trennzeichen & "videos, vision, wanted, weekly, whales, wheels, window, woeful, worlds, writer,"
    csv_trennzeichen = csv_trennzeichen & "zombie, adapted, adapter, address, affects, agenten, aimtrak, airport, alcohol, alegria,"
    csv_trennzeichen = csv_trennzeichen & "amazing, america, amstrad, animals, another, antique, archive, article, assault, attempt,"
    csv_trennzeichen = csv_trennzeichen & "auction, awkward, bargain, battery, beatles, behaved, believe, beneath, berguna, biggest,"
    csv_trennzeichen = csv_trennzeichen & "biology, bizarre, blender, bonkers, boombox, borders, britain, british, brought, bubbles,"
    csv_trennzeichen = csv_trennzeichen & "bublcam, cabinet, calorie, cantata, capsule, captain, capture, cathode, changed, channel,"
    csv_trennzeichen = csv_trennzeichen & "classic, cleaner, climate, commons, compact, complex, compost, concept, concert, connect,"
    csv_trennzeichen = csv_trennzeichen & "console, contact, control, costume, couldve, covered, created, curious, cutting, dashcam,"
    csv_trennzeichen = csv_trennzeichen & "decline, desktop, digital, discman, display, dropped, dumbing, eastern, economy, ecotank,"
    csv_trennzeichen = csv_trennzeichen & "editing, edition, eevblog, ejected, elcaset, elecrow, encoder, endless, enemies, english,"
    csv_trennzeichen = csv_trennzeichen & "enjoyed, episode, erotica, experts, explore, express, factual, failure, fastest, fighter,"
    csv_trennzeichen = csv_trennzeichen & "figures, finding, fitting, flatten, follies, footage, foreign, forging, formats, formula,"
    csv_trennzeichen = csv_trennzeichen & "freight, friends, gadgets, gateway, geeekpi, genders, germans, germany, getting, glasses,"
    csv_trennzeichen = csv_trennzeichen & "graphic, gravity, grundig, hammond, haunted, healthy, hearted, hewlett, highest, history,"
    csv_trennzeichen = csv_trennzeichen & "holiday, horizon, however, insight, install, invited, italian, journey, kaisers, keepers,"
    csv_trennzeichen = csv_trennzeichen & "killing, kitchen, kopriso, landing, lapoint, lasagna, learned, legally, letters, machine,"
    csv_trennzeichen = csv_trennzeichen & "mansion, massive, maximal, message, michael, minutes, missing, mission, mistake, monitor,"
    csv_trennzeichen = csv_trennzeichen & "monster, montage, mystery, nescafe, netgear, nibbles, norways, nuclear, numbers, nyquist,"
    csv_trennzeichen = csv_trennzeichen & "obscure, october, oddness, odyssey, okashda, olympic, onboard, opening, openvms, optical,"
    csv_trennzeichen = csv_trennzeichen & "ordered, origins, outdoor, packard, panapic, pandora, patreon, perfect, philips, physics,"
    csv_trennzeichen = csv_trennzeichen & "pioneer, planets, plastic, player_, playing, popular, powered, printer, private, problem,"
    csv_trennzeichen = csv_trennzeichen & "product, program, project, propels, provide, putting, quantel, radiant, rainbow, reality,"
    csv_trennzeichen = csv_trennzeichen & "reasons, records, redneck, repairs, rescued, reverse, reviews, richard, rokblok, running,"
    csv_trennzeichen = csv_trennzeichen & "russian, samples, samsung, sarista, scanner, scarlet, science, scratch, seeburg, selamat,"
    csv_trennzeichen = csv_trennzeichen & "serious, setaram, setting, shannon, shelley, showing, silicon, simcity, similar, sinking,"
    csv_trennzeichen = csv_trennzeichen & "sleeves, smaller, smashed, society, spanish, speaker, special, species, spinbox, spinner,"
    csv_trennzeichen = csv_trennzeichen & "squeeze, station, stopped, stories, strange, streets, studios, summary, sunbeam, support,"
    csv_trennzeichen = csv_trennzeichen & "surfing, swedish, systems, talking, tefifon, testing, theorem, thermal, thomson, through,"
    csv_trennzeichen = csv_trennzeichen & "tickets, toaster, toilets, traffic, typical, ukulele, unclear, unusual, updated, upgrade,"
    csv_trennzeichen = csv_trennzeichen & "useless, usually, vectrex, version, viewers, viewing, vintage, virtual, visited, volcano,"
    csv_trennzeichen = csv_trennzeichen & "voyager, walkman, wallace, warming, watches, welcome, winders, windows, winning, writing,"
    csv_trennzeichen = csv_trennzeichen & "yonanas, youtube, absolute, adapters, affected, airforce, airports, american, analogue, analyser,"
    csv_trennzeichen = csv_trennzeichen & "analysis, analyzer, andersen, anderson, aperture, approach, asksunny, assembly, atlantic, audities,"
    csv_trennzeichen = csv_trennzeichen & "backbone, baseball, beginner, berliner, beverage, birthday, branding, browsing, building, business,"
    csv_trennzeichen = csv_trennzeichen & "calcupen, capsules, cassette, catching, changers, changing, citation, classics, cleaning, climbing,"
    csv_trennzeichen = csv_trennzeichen & "cocktail, columbia, combined, computer, conquest, controls, cordless, coverage, creative, critical,"
    csv_trennzeichen = csv_trennzeichen & "cronixie, cultural, cylinder, dataplay, december, defeated, dialects, digitise, directed, director,"
    csv_trennzeichen = csv_trennzeichen & "disaster, division, dolphins, doughnut, dragon32, dumpster, duratool, dynamite, eclectic, electron,"
    csv_trennzeichen = csv_trennzeichen & "enclaves, entirely, episodes, esposita, exclaves, executel, explored, extended, features, february,"
    csv_trennzeichen = csv_trennzeichen & "fighting, findings, finished, flamingo, flexplay, flipping, floating, followed, fountain, frontier,"
    csv_trennzeichen = csv_trennzeichen & "gauntlet, geertsen, genitive, gigantic, giveaway, graphics, greatest, greeting, grierson, hamilton,"
    csv_trennzeichen = csv_trennzeichen & "handheld, happened, headache, illusion, included, includes, industry, inspired, intercom, intercut,"
    csv_trennzeichen = csv_trennzeichen & "internet, invaders, inventor, japanese, joystick, junkyard, keyboard, koestler, language, launched,"
    csv_trennzeichen = csv_trennzeichen & "laziness, lectures, levitate, lightgun, linesman, machines, mailtime, material, mayflash, metaphor,"
    csv_trennzeichen = csv_trennzeichen & "military, minibeam, minidisc, national, nextstep, nintendo, offering, offshore, orbitrac, ordnance,"
    csv_trennzeichen = csv_trennzeichen & "pachinko, paintbox, particle, personal, physical, pioneers, platform, playable, playlist, playtape,"
    csv_trennzeichen = csv_trennzeichen & "playtest, polaroid, portable, possible, practice, pressure, printers, probably, problems, produced,"
    csv_trennzeichen = csv_trennzeichen & "producer, projects, pyramids, reacting, recorded, recorder, regional, remember, resolusi, retropie,"
    csv_trennzeichen = csv_trennzeichen & "revealed, reverber, rockfire, sampling, sandwich, security, seinfeld, shootout, sinclair, skittles,"
    csv_trennzeichen = csv_trennzeichen & "smallest, smartest, software, solmecke, solution, soundbox, speakers, speaking, spectrum, spinners,"
    csv_trennzeichen = csv_trennzeichen & "spitfire, squaring, sriracha, starlink, stirling, subjects, superman, surprise, surround, swapping,"
    csv_trennzeichen = csv_trennzeichen & "switches, tanashin, teachers, teardown, techmoan, technics, termahal, terpedas, terrible, theories,"
    csv_trennzeichen = csv_trennzeichen & "together, trappist, treasure, tutorial, twilight, ultimarc, ultimate, unboxing, universe, unstable,"
    csv_trennzeichen = csv_trennzeichen & "upgrades, validate, vertical, whatever, wireless, workflow, yosemite, according, addressed, addresses,"
    csv_trennzeichen = csv_trennzeichen & "ambitions, ambrosino, announced, anxieties, assembled, attending, automatic, available, awareness, backwards,"
    csv_trennzeichen = csv_trennzeichen & "batteries, betamovie, bluetooth, breakfast, broadcast, calendars, cardboard, cartridge, cassettes, challenge,"
    csv_trennzeichen = csv_trennzeichen & "chocolate, christmas, commodore, companies, component, computers, considers, consumers, continued, continues,"
    csv_trennzeichen = csv_trennzeichen & "corporate, creatures, criticism, crunching, curiosity, designing, detective, developed, direction, directory,"
    csv_trennzeichen = csv_trennzeichen & "discovery, dragonfly, effective, eggmaster, elections, elevators, employing, engenders, equalizer, evolution,"
    csv_trennzeichen = csv_trennzeichen & "executive, expensive, explained, exploring, fertility, following, foolproof, forgotten, halloween, handhelds,"
    csv_trennzeichen = csv_trennzeichen & "headphone, including, ingenious, interface, internals, introduce, invention, invisible, involving, knowledge,"
    csv_trennzeichen = csv_trennzeichen & "languages, laserdisc, macintosh, measuring, molecular, narrative, newspaper, obliquely, obviously, packaging,"
    csv_trennzeichen = csv_trennzeichen & "panasonic, paperless, polarized, preparing, presented, presenter, processor, programme, projector, protector,"
    csv_trennzeichen = csv_trennzeichen & "providing, questions, raspberry, recorders, repairing, restoring, resurrect, revisited, scientist, screening,"
    csv_trennzeichen = csv_trennzeichen & "slideshow, sometimes, starlight, statement, strangest, structure, subtitled, telephone, terraform, transport,"
    csv_trennzeichen = csv_trennzeichen & "treatment, turntable, typically, ultrawide, upgrading, videodisc, artificial, assembling, australian, background,"
    csv_trennzeichen = csv_trennzeichen & "beginnings, calculator, california, captioning, chromecast, compatible, conducting, creativity, decreasing, disposable,"
    csv_trennzeichen = csv_trennzeichen & "dissecting, dockintosh, documented, dramatised, experiment, fahrenheit, fellowship, frequently, futuristic, guaranteed,"
    csv_trennzeichen = csv_trennzeichen & "gyroscopic, headphones, hurricanes, impossible, incredibly, individual, installing, mechanical, multimedia, mysterious,"
    csv_trennzeichen = csv_trennzeichen & "phonograph, playbutton, popularity, powerhouse, precession, production, programmes, propulsion, protection, psychology,"
    csv_trennzeichen = csv_trennzeichen & "recordings, recreating, reinvented, repopulate, resolution, retrobrite, revisiting, revolution, scientific, scientists,"
    csv_trennzeichen = csv_trennzeichen & "scratching, soundtrack, specialist, structured, structures, summarised, superpower, superscope, supersonic, surprising,"
    csv_trennzeichen = csv_trennzeichen & "surrounder, technology, telephones, television, typewriter, ubiquitous, unbeatable, underlying, unexpected, validation,"
    csv_trennzeichen = csv_trennzeichen & "videodiscs, vinylvideo, visualiser, accusations, agriculture, alternative, anniversary, buckminster, calculating, communicate,"
    csv_trennzeichen = csv_trennzeichen & "differently, discoveries, distinctive, documentary, electronics, elucidation, exclusively, experiments, explanatory, gingerbread,"
    csv_trennzeichen = csv_trennzeichen & "grammatical, impressions, independent, innovations, interactive, interesting, netherlands, oktoberfest, performance, personality,"
    csv_trennzeichen = csv_trennzeichen & "playstation, reflections, reinventing, restoration, sarcophagus, specialbuys, stereotypes, thermostats, transistors, translating,"
    csv_trennzeichen = csv_trennzeichen & "workstation, appreciation, broadcasters, broadcasting, concentrates, considerable, construction, encyclopedia, implications, intellectual,"
    csv_trennzeichen = csv_trennzeichen & "interviewees, introduction, investigated, observations, occasionally, philosophers, presentation, rechargeable, registration, standardised,"

    csv_trennzeichen = ","

    Dim str_eins As String
    Dim str_zwei As String

    '
    ' Parameterpruefung
    ' csv_werte        ==> muss gesetzt sein ( laenge > 0 / ohne trim )
    ' csv_trennzeichen ==> muss gesetzt sein ( laenge > 0 / ohne trim )
    '
    If (Len(csv_werte) > 0) And (Len(csv_trennzeichen) > 0) Then
        
        position_ende = 1
        
        position_start = 1

        csv_feld_zaehler = 1

        '
        ' Die While-Schleife wird solange durchlaufen, wie
        ' ... das Trennzeichen noch gefunden wird (position_ende > 0)
        ' ... der Feldzaehler noch kleiner 32123 ist
        '
        While ((position_ende > 0) And (csv_feld_zaehler < 32123))

            '
            ' End Trennzeichen suchen
            ' Von der Start-Position wird das Trennzeichen gesucht.
            ' Die Position wird in der Variablen "position_ende" gespeichert.
            '
            position_ende = InStr(position_start, csv_werte, csv_trennzeichen)
            
            '
            ' Wert in "position_ende"
            ' Wurde noch ein weiteres vorkommen des Trennzeichens gefunden, ist
            ' die Variable "position_ende" groesser 0. Das Ergebnis wird vom
            ' Startindex bis zu dem naechsten Auftreten des Trennzeichens gelesen.
            '
            ' Ist kein weiteres Trennzeichen vorhanden, wird vom Startindex bis
            ' zum Stringende der Teilstring gelesen.
            '
            If (position_ende > 0) Then

                akt_teil_string = Mid(csv_werte, position_start, position_ende - position_start)

            Else

                akt_teil_string = Mid(csv_werte, position_start, Len(csv_werte))

            End If
            
            '
            ' Pruefung: String fuer Verarbeitung vorhanden?
            ' Der aktuelle Teilstring wird nur verarbeitet, wenn dieser kein Leerstring ist.
            '
            If (Trim(akt_teil_string) <> LEER_STRING) Then
            
                akt_teil_string = Trim(akt_teil_string)
                
                str_eins = "_" & akt_teil_string & "_"
                
                str_zwei = "_" & UCase(Left(akt_teil_string, 1)) & Right(akt_teil_string, Len(akt_teil_string) - 1) & "_"
                
                datei_name_neu = Replace(datei_name_neu, str_eins, str_zwei)
                
            End If
            
            '
            ' Berechnung der neuen Startposition
            ' Die Variable "position_start" zeigt auf das erste Zeichen der Rueckgabe.
            ' Diese Position ist die aktuell gefundene Position des Trennzeichens zuzuglich dessen Laenge.
            '
            position_start = position_ende + Len(csv_trennzeichen)

            ' Feldzaehler
            ' Feldzaehler um 1 erhoehen und naechsten Schleifendurchlauf machen.
            '
            csv_feld_zaehler = csv_feld_zaehler + 1

        Wend

    End If

    datei_name_neu = Replace(datei_name_neu, "##MARKIERUNG_1##_", "")
    datei_name_neu = Replace(datei_name_neu, "_##MARKIERUNG_2##", "")
    
    renameWorteEnglish = datei_name_neu

End Function

'################################################################################
'
' Liefert von dem uebergebenen Dateinamen, die Erweiterung zurueck.
'
' Eine Erweiterung besteht aus dem Punkt und 1 bis 4 Zeichen.
'
' Ist keine Erweiterung vorhanden, wird ein Leerstring zurueck gegebgen.
'
' ? getErweiterung( "c:\Daten\test\test_datei.txt"   ) --> .txt
' ? getErweiterung( "c:\Daten\test\test_datei.1234"  ) --> .1234
' ? getErweiterung( "c:\Daten\test\test_datei.12345" ) --> ""
'
' ? getErweiterung( "c:\Daten\test\test_datei"       ) -->
' ? getErweiterung( ""                               ) -->
' ? getErweiterung( "............."                  ) -->
' ? getErweiterung( "c:\Daten\tes.t\test_datei"      ) -->
'
' PARAMETER: pDateiName     = String aus welchem die Dateierweiterung ermittelt werden soll
'
' RETURN : die gefundene Dateierweiterung, oder einen Leerstring, wenn nichts ermittelt werden konnte.
'
Public Function getErweiterung(ByVal pDateiName As String) As String

On Error GoTo errGetErweiterung

Dim position_trenner As Integer
Dim akt_zeichen      As String
Dim knz_while_aktiv  As Boolean

    getErweiterung = ""

    knz_while_aktiv = True

    '
    ' Es wird von hinten nach vorne gesucht. Die Leseposition
    ' wird auf das letzte Zeichen der Eingabe gestellt.
    '
    position_trenner = Len(pDateiName)

    '
    ' Die While-Schleife laeuft solange wie, die Position des Trennzeichens
    ' groesser 0 ist und das Kennzeichen "knz_while_aktiv" auf "true" steht.
    '
    While ((position_trenner > 0) And (knz_while_aktiv))

        akt_zeichen = Mid(pDateiName, position_trenner, 1)

        '
        ' Bei einem Punkt ist das Trennzeichen fuer eine Erweiterung gefunden.
        '
        ' Da das Trennzeichen selber mitaufgenommen werden soll, wird die
        ' Position um ein Zeichen weiter runtergezaehlt.
        '
        If (akt_zeichen = ".") Then

            position_trenner = position_trenner - 1

            knz_while_aktiv = False

        '
        ' Ein Trennzeichen fuer den Pfad darf im Suchprozess nicht auftreten.
        '
        ' Die Dateierweiterung muss vor einem solchen Zeichen liegen.
        '
        ' Damit die Schleife beendet werden kann wird die Position des
        ' Trennzeichens auf -1 gestellt.
        '
        ElseIf ((akt_zeichen = "/") Or (akt_zeichen = "\")) Then

            position_trenner = -1

            knz_while_aktiv = False

        Else

        '
        ' Im Normalfall wird der Leseprozess auf das naechste Zeichen gesetzt.
        '
            position_trenner = position_trenner - 1

        End If

    Wend

    '
    ' Pruefung: Erweiterung gefunden ?
    '
    ' Es muss eine Trennzeichenposition geben, welche groesser 0 ist.
    ' Die Trennzeichenposition darf nicht am Ende liegen.
    '
    ' Die Erweiterung muss mindestens 1 Zeichen haben.
    ' Die Erweiterung darf maximal 5 Zeichen haben.
    '
    If ((position_trenner > 0) And (Len(pDateiName) - position_trenner > 1) And (Len(pDateiName) - position_trenner <= 5)) Then

        '
        ' Wurde ein trennzeichen gefunden, wird das Funktionsergebnis gesetzt.
        '
        ' Es werden von Rechts die Zeichen bis zur Trennzeichenposition zurueckgegeben
        '
        getErweiterung = Right(pDateiName, Len(pDateiName) - position_trenner)

    End If

    '
    ' DoEvents aufrufen
    '
    DoEvents

    '
    ' Funktion verlassen
    '
    Exit Function

errGetErweiterung:

    getErweiterung = LEER_STRING

    Exit Function

End Function

