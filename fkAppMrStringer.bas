Attribute VB_Name = "fkAppMrStringer"
Option Explicit

Public m_knz_aktiv                       As Boolean ' public wegen initialisierung
Private m_csv_feld_nummer                As Integer
Private m_knz_join_anfang                As Integer
Private m_toggle_mr_stringer_fkt         As Boolean
Private m_zaehler_debug_print            As Integer
Private m_zaehler_string_it              As Integer

Public Const FKT_AUSRICHTER_POSITION = 1
Public Const FKT_AUSRICHTER_STRING = 2
Public Const FKT_CAMEL_CASE = 3
Public Const FKT_CHECK_LEERSTRING = 4
Public Const FKT_CLIP_POSITION = 5
Public Const FKT_CLIP_GET_TEXT = 6
Public Const FKT_CMD_RENAME = 7
Public Const FKT_CSV_2_ZEILE = 8
Public Const FKT_CSV_CASE = 9
Public Const FKT_CSV_ERSTELLE_CSV = 10
Public Const FKT_CSV_REPLACE_MARKIERUNG_MIT_CSV = 11
Public Const FKT_CSV_SWAP = 12
Public Const FKT_DEBUG_AUSGABE = 13
Public Const FKT_DEKLARATION = 14
Public Const FKT_DUPLIZIERUNG = 15
Public Const FKT_ENTFERNE_LEERZEILEN = 16
Public Const FKT_ERSTELLE_BLOCK = 17
Public Const FKT_ERSTELLE_NAMEN = 18
Public Const FKT_ERSTELLE_XML = 19
Public Const FKT_ERSTELLE_XML_2 = 20
Public Const FKT_EXTRAHIERE_WORTE = 21
Public Const FKT_FORMAT_TXT = 22
Public Const FKT_GENERATOR_IF = 23
Public Const FKT_GENERATOR_IF_2 = 24
Public Const FKT_GETTER_SETTER_JAVA = 25
Public Const FKT_GETTER_SETTER_JAVA_SCRIPT = 26
Public Const FKT_GETTER_SETTER_VB = 27
Public Const FKT_GET_DIR = 28
Public Const FKT_GET_DOPPELTE_VORKOMMEN = 29
Public Const FKT_GET_EINMALIGE_VORKOMMEN = 30
Public Const FKT_GET_UNIQUE = 31
Public Const FKT_GREP_PLUS_MINUS = 32
Public Const FKT_GREP_DUPLIZIERE_MARKZEILEN = 33
Public Const FKT_GREP_MARK = 34
Public Const FKT_GREP_WORT = 35
Public Const FKT_GREP_ZAHLEN = 36
Public Const FKT_JAVA_GENERATOR = 37
Public Const FKT_JAVA_XML_WRITER_NUMMER = 38
Public Const FKT_JAVA_XML_WRITER_STRING = 39
Public Const FKT_JSON_LESEN_SCHREIBEN = 40
Public Const FKT_LEERZEILEN_EINFUEGEN = 41
Public Const FKT_MAKE_LONG_DATUM = 42
Public Const FKT_MARKIERE_DOPPELTE_PLUS = 43
Public Const FKT_MARKIERE_DOPPELTE_PLUS_MINUS = 44
Public Const FKT_MARKIERE_VORNE_ODER_HINTEN = 45
Public Const FKT_MARKIERE_VORNE_UND_HINTEN = 46
Public Const FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE = 47
Public Const FKT_MARKIERE_WORT = 48
Public Const FKT_NOTES_DEBUG_FELD_WERTE = 49
Public Const FKT_NOTES_LESEN_SCHREIBEN = 50
Public Const FKT_STRING_REMOVE = 51
Public Const FKT_SET_NULL = 52
Public Const FKT_SET_TRENNZEICHEN = 53
Public Const FKT_SET_TRENNZEICHEN_STR = 54
Public Const FKT_SET_TRENNZEICHEN_VOR = 55
Public Const FKT_SET_TRENNZEICHEN_ZURUECK = 56
Public Const FKT_SINGLETON_JAVA = 57
Public Const FKT_SORTIEREN_ABC = 58
Public Const FKT_SORTIEREN_DATUM = 59
Public Const FKT_SORTIEREN_LAENGE = 60
Public Const FKT_SORTIEREN_ZUFALL = 61
Public Const FKT_SPLIT = 62
Public Const FKT_STRING_IT = 63
Public Const FKT_STRING_LIT = 64
Public Const FKT_CALC_SUMME = 65
Public Const FKT_TO_ZEILE = 66
Public Const FKT_TRIM = 67
Public Const FKT_TRIM_X = 68
Public Const FKT_UCASE_LCASE = 69
Public Const FKT_UMDREHEN = 70
Public Const FKT_VERSCHIEBEN = 71
Public Const FKT_ZAEHLER = 72
Public Const FKT_ZEILEN_ADD = 73
Public Const FKT_ZEILEN_BOOLEAN = 74
Public Const FKT_HEX_DUMP = 75
Public Const FKT_GROUP_NACH_STRING = 76
Public Const FKT_BLOCK_ZUFALL = 78
Public Const FKT_CSV_JAVA_PROP = 79
Public Const FKT_GREP_MARK_HINTEN = 90
Public Const FKT_GREP_MARK_VORNE = 91
Public Const FKT_MARKIERE_STR_VORNE_UND_HINTEN = 92
Public Const FKT_MASKIERE_ANFZEICHEN = 93
Public Const FKT_MARKIERE_CSV_VORNE_ODER_HINTEN = 94

Public Const STR_VAR_NAME_PROPERTIES_LOKAL = "inst_properties" ' " & STR_VAR_NAME_PROPERTIES_LOKAL & "

Public Const LEER_STRING = ""
Public Const TRENN_STRING_6 = "#6"
Public Const TRENN_STRING_7 = "#7"
Public Const TRENN_STRING_8 = "#8"
Private Const TRENN_STRING_9 = "#9"
Private Const MARKIERUNG_DOPPELTE_VORKOMMEN = "#D#"
Public Const AUSRICHT_STRING_TEMP_1 = "##AUSRICHT_STRING_TEMP_1##" ' " & AUSRICHT_STRING_TEMP_1 & "
Public Const AUSRICHT_STRING_TEMP_2 = "##AUSRICHT_STRING_TEMP_2##" ' " & AUSRICHT_STRING_TEMP_2 & "

Private Const GUELTIGE_ZEICHEN_DATEI_NAME = "enirstl_audhgocmfbkVvwz1pSDA0E2RBGMIPKF9UNW3L78H4T5CZJy6xjOYXqQ&,'()" ' sortiert Haeufigkeit Deutsch
Private Const NULL_ZIFFERN_100 = "00000000000000000000000000000000000000000000000000000000000000000000000000000000"

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

Dim cls_string_array        As clsStringArray
Dim zeichen_zeilenumbruch   As String
Dim ergebnis_fkt            As String
Dim aktuelle_zeile          As String
Dim inhalt_markierung       As String
Dim temp_string_1           As String
Dim akt_zeile_mark          As String
Dim temp_string_2           As String
Dim temp_string_3           As String
Dim temp_double_1           As Double
Dim temp_double_2           As Double
Dim ab_position             As Long
Dim bis_position            As Long
Dim zeilen_zaehler          As Long
Dim zeilen_anzahl           As Long
Dim temp_long_1             As Long
Dim temp_long_2             As Long
Dim temp_long_3             As Long
Dim knz_benutze_markierung  As Boolean
Dim knz_schleifen_durchlauf As Boolean
    
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
    ' Das vorlaeufige Ergebnis auf einen Leerstring stellen.
    '
    ergebnis_fkt = LEER_STRING
    
    If (pFunktion = FKT_CSV_2_ZEILE) Then
        '
        ' Funktion "CSV 2 Zeile"
        ' Nach jedem Trennzeichen aus dem Parameter "pOptString1" wird ein Zeilenumbruch
        ' eingefuegt. Dieses wird einmal mit Loeschung des Trennzeichens selber gemacht und
        ' einmal bleibt dass Trennzeichen selber erhalten.
        '
        If (pOptString1 = LEER_STRING) Then
        
        Else
            
            If (m_toggle_mr_stringer_fkt) Then
                
                ergebnis_fkt = Replace(pString, pOptString1, pOptString1 & Chr(13) & Chr(10))
            
            Else
                
                ergebnis_fkt = Replace(pString, pOptString1, Chr(13) & Chr(10))
            
            End If
            
        End If
    
    ElseIf (pFunktion = FKT_STRING_LIT) Then
        '
        ' Funktion "String Literale"
        ' Die Funktion fuer die Ermittlung der String-Literale ist eine eigenstaendige
        ' Funktion und wird nur aufgerufen.
        '
        ergebnis_fkt = getStringLitKonst(pString)
    
    ElseIf (pFunktion = FKT_STRING_REMOVE) Then
        '
        ' Funktion "Remove"
        ' Die selektierte Zeichenkette wird aus dem Eingabestring geloescht.
        '
        ergebnis_fkt = Replace(pString, Mid(pString, pSelStart + 1, pSelLength), LEER_STRING)
    
    ElseIf (pFunktion = FKT_EXTRAHIERE_WORTE) Then
    
        If (m_toggle_mr_stringer_fkt) Then
            
            ergebnis_fkt = extrahiereWoerter(pString, pOptString1, 200)
        
        Else
            
            ergebnis_fkt = extrahiereWoerter(pString, pOptString1, 3)
        
        End If
    
    ElseIf (pFunktion = FKT_FORMAT_TXT) Then
    
        If (m_toggle_mr_stringer_fkt) Then
        
            ergebnis_fkt = getStringMaxCols(pString, 55, LEER_STRING, Chr(13) & Chr(10))
            
        Else
        
            ergebnis_fkt = getStringMaxCols(pString, 80, LEER_STRING, Chr(13) & Chr(10))
        
        End If
    
    Else
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
        ' ueberfuehrt. Sollte die Funktion "nothing" zurueckkommen ist die
        ' Funktion "startMrStringer" beendet.
        '
        Set cls_string_array = startMultiline(pString)
        
        If (cls_string_array Is Nothing) Then
        '
        ' keine Aktionen machen, wenn das String-Array-Objekt nicht gesetzt ist.
        '
        Else
            '
            ' Voreinstellung dass ein Schleifendurchlauf notwendig ist
            '
            knz_schleifen_durchlauf = True
            '
            ' Verarbeitung Selektion
            ' Es wird geprueft, ob eine Position in "pSelStart" vorhanden ist.
            '
            If (pSelStart <= 0) Then
                '
                ' Liegt keine Startposition einer Selektion vor, werden die
                ' Variablen Ab- und Bis auf 0 und das Kennzeichen fuer die
                ' Verwendung der Markierung auf false gestellt.
                '
                ab_position = 0
                bis_position = 0
                knz_benutze_markierung = False
            
            Else
                '
                ' Liegt eine Startposition vor, muss dessen relative Startposition zu
                ' dem letzten Zeilenumbruch ermittelt werden. Eine Markierung gilt nur
                ' fuer eine Zeile, nicht fuer den gesamten Text aus "pString".
                '
                ' Es wird das letzte Zeilenumbruchszeichen vor der Startposition gesucht.
                '
                temp_long_1 = getLetztePositionVorPos(pString, getBenutztesChr13(pString), pSelStart)
                '
                ' Wird ein Zeilenumbruch gefunden, wird dessen absolute Position von der
                ' Selektionsstartposition abgezogen.
                '
                ' Wird kein Zeilenumbruch gefunden (Markierung befindet sich in Zeile 1),
                ' wird auf die Selektionsstartposition eine Position draufgerechnet.
                '
                ' 1234567890 1234567890 1234567890 1234567890 1234567890 1234567890
                ' 1234567890 1234567890 1234567890 1234567890 1234567890 1234567890
                '
                If (temp_long_1 > 0) Then
                
                    ab_position = pSelStart - temp_long_1
                
                Else
                    
                    ab_position = pSelStart + 1
                
                End If
                '
                ' Bestimmung Bis-Position
                ' Auf die Ab-Position wird die Selektionslaenge hinzugezaehlt und abschliessend
                ' wieder einer abgezogen (da die Startposition schon selber enthalten ist).
                '
                bis_position = (ab_position + pSelLength) - 1
                '
                ' Kennzeichen "knz_benutze_markierung"
                ' Das Kennzeichen ist TRUE, wenn die Ab-Position groesser gleich 0 ist und
                ' die Bis-Position gleich oder groesser ist.
                '
                knz_benutze_markierung = (ab_position >= 0) And (bis_position >= ab_position)
            
            End If
            
            If (pFunktion = FKT_CSV_REPLACE_MARKIERUNG_MIT_CSV) Then
                
                If (pSelLength > 0) Then
                
                    temp_string_1 = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                End If
                
                If (temp_string_1 <> LEER_STRING) Then
                
                    ergebnis_fkt = Replace(pString, temp_string_1, pOptString1)
                    
                Else
                
                    ergebnis_fkt = pString
                
                End If
                
                knz_schleifen_durchlauf = False
            '
            ' Vorbereitung Funktion Sortieren
            '
            ElseIf (pFunktion = FKT_SORTIEREN_ABC) Then
            
                Call cls_string_array.startSortierung(1, m_toggle_mr_stringer_fkt, False, knz_benutze_markierung, ab_position, bis_position)
                
                ergebnis_fkt = cls_string_array.toString(zeichen_zeilenumbruch, True)
                
                knz_schleifen_durchlauf = False
            '
            ' Vorbereitung Funktion Sortieren Laenge
            '
            ElseIf (pFunktion = FKT_SORTIEREN_ZUFALL) Then
                
                Call cls_string_array.startZufallsUmsortierung
                
                ergebnis_fkt = cls_string_array.toString(zeichen_zeilenumbruch, True)
                
                knz_schleifen_durchlauf = False
            '
            ' Funktion Markiere Wort
            '
            ElseIf (pFunktion = FKT_MARKIERE_WORT) Then
                
                '
                ' Suchwort
                ' Ermittlung der zu markierenden Zeichenkette.
                ' Das ist der String aus der Markierung, welches ein
                ' Wort, oder auch mehrere Zeichen sein koennen.
                '
                temp_string_1 = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                '
                ' Suchwort ersetzen
                ' Die Suche wird ueber die Funktion "startReplaceSuchWorte" gemacht.
                ' Die Such/Ersatzstring werden in der Form "suchwort=ersatzwort" uebergeben.
                ' Es gibt hier nur ein solches Such/Ersatzpaar, daher muss kein Zeilenumbruch
                ' in den Parameter eingebaut werden.
                '
                ' ? startReplaceSuchWorte("A=c" & chr(13) & "B=d", "AABB")  = ccdd
                '
                ergebnis_fkt = startReplaceSuchWorte(temp_string_1 & "=" & IIf(m_toggle_mr_stringer_fkt, TRENN_STRING_8, TRENN_STRING_9) & temp_string_1, pString)
                
                '
                ' Es muss bei der Funktion "FKT_MARKIERE_WORT" kein weiterer Schleifendurchlauf
                ' gemacht werden. Die Variable "knz_schleifen_durchlauf" wird auf FALSE
                ' gestellt.
                '
                knz_schleifen_durchlauf = False
                
            ElseIf (pFunktion = FKT_LEERZEILEN_EINFUEGEN) Then
            
                temp_long_1 = 0
                
                knz_schleifen_durchlauf = True
            '
            ' Vorbereitung Funktion Sortieren Laenge
            '
            ElseIf (pFunktion = FKT_SORTIEREN_LAENGE) Then
                
                Call cls_string_array.startSortierung(234, m_toggle_mr_stringer_fkt, False, knz_benutze_markierung, ab_position, bis_position)
                
                ergebnis_fkt = cls_string_array.toString(zeichen_zeilenumbruch, True)
                
                knz_schleifen_durchlauf = False
                
            ElseIf (pFunktion = FKT_GROUP_NACH_STRING) Then
            
                knz_schleifen_durchlauf = True
                
            ElseIf (pFunktion = FKT_SORTIEREN_DATUM) Then
            
                If (ab_position = 0) Then
                    
                    ab_position = 1
                
                End If
                
                Call cls_string_array.startSortierung(1, m_toggle_mr_stringer_fkt, True, True, ab_position, bis_position)
                
                ergebnis_fkt = cls_string_array.toString(zeichen_zeilenumbruch, True)

                knz_schleifen_durchlauf = False
                
            ElseIf (pFunktion = FKT_CHECK_LEERSTRING) Then
            
            '
            ' Vorbereitung Funktion StringIt
            '
            ElseIf (pFunktion = FKT_STRING_IT) Then
            
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
            '
            ' Vorbereitung Funktion XML-Erstellung
            ' Umdrehen der boolschen Variable (mit / ohne Vorlauf)
            ' Parameterkennzeichen bestimmt ob volle oder nur einzelne TAGS
            '
            ElseIf (pFunktion = FKT_ERSTELLE_XML) Or (pFunktion = FKT_ERSTELLE_XML_2) Then
        
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
                
                If (pSelLength >= 3) And (pSelLength < 1000) Then
                    
                    temp_long_1 = pSelLength
                
                Else
                    
                    temp_long_1 = 80
                
                End If
                
            ElseIf (pFunktion = FKT_ZEILEN_ADD) Then
            
                temp_long_1 = 10
                
                knz_schleifen_durchlauf = True
            '
            ' Vorbereitung Funktion Split
            '
            ElseIf (pFunktion = FKT_SPLIT) Then
            
                knz_schleifen_durchlauf = ab_position > 0
                
                temp_string_1 = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                temp_long_2 = Len(temp_string_1)
                
                temp_long_1 = ab_position
            '
            ' Vorbereitung Funktion Zaehler
            '
            ElseIf (pFunktion = FKT_ZAEHLER) Then
                
                If (pSelLength >= 3) And (pSelLength < 70) Then
                    
                    temp_long_1 = pSelLength
                
                Else
                    
                    temp_long_1 = 6
                
                End If
            
            ElseIf (pFunktion = FKT_GREP_DUPLIZIERE_MARKZEILEN) Then
            
                knz_schleifen_durchlauf = ab_position > 0
                
                If (pSelLength > 0) Then
                
                    temp_string_1 = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                Else
                    
                    temp_long_1 = getPosWortAnfang(pString, pSelStart)

                    temp_long_2 = getPosWortende(pString, pSelStart)
                                
                    If ((temp_long_1 > 0) And (temp_long_2 > temp_long_1)) Then
                        
                        temp_string_1 = Mid(pString, temp_long_1, (temp_long_2 - (temp_long_1)) + 1)
                        
                    End If
    
                End If
            
            '
            ' Vorbereitung Funktion Grep
            '
            ElseIf (pFunktion = FKT_GREP_PLUS_MINUS) Or (pFunktion = FKT_GREP_MARK) Or (pFunktion = FKT_GREP_MARK_VORNE) Or (pFunktion = FKT_GREP_MARK_HINTEN) Then

                knz_schleifen_durchlauf = ab_position > 0
                
                If (pSelLength > 0) Then
                
                    temp_string_1 = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                Else
                    
                    temp_long_1 = getPosWortAnfang(pString, pSelStart)

                    temp_long_2 = getPosWortende(pString, pSelStart)
                                
                    If ((temp_long_1 > 0) And (temp_long_2 > temp_long_1)) Then
                        
                        temp_string_1 = Mid(pString, temp_long_1, (temp_long_2 - (temp_long_1)) + 1)
                        
                    End If
    
                End If
                
                If (pFunktion = FKT_GREP_MARK_VORNE) Then
                
                    temp_string_2 = TRENN_STRING_7
                    temp_string_3 = LEER_STRING
                    
                    pFunktion = FKT_GREP_MARK
                    
                    temp_string_1 = pOptString1
                
                ElseIf (pFunktion = FKT_GREP_MARK_HINTEN) Then
                    
                    temp_string_2 = LEER_STRING
                    temp_string_3 = TRENN_STRING_7
                
                    pFunktion = FKT_GREP_MARK
                    
                    temp_string_1 = pOptString1
                    
                
                ElseIf (m_toggle_mr_stringer_fkt) Then
                    
                    temp_string_2 = TRENN_STRING_7
                    temp_string_3 = LEER_STRING
                
                Else
                    
                    temp_string_2 = LEER_STRING
                    temp_string_3 = TRENN_STRING_7
                
                End If
                
                knz_schleifen_durchlauf = True
            '
            ' Vorbereitung Funktion Grep Wort
            '
            ElseIf (pFunktion = FKT_GREP_WORT) Then
            
                temp_string_1 = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                temp_long_3 = Len(temp_string_1)
            
            '
            ' Vorbereitung Funktionen Generator If
            '
            ElseIf (pFunktion = FKT_GENERATOR_IF) Or (pFunktion = FKT_GENERATOR_IF_2) Then
            
                temp_string_3 = "if"
                
                inhalt_markierung = getStringAbBis(pString, ab_position, bis_position)
                
                temp_long_2 = Len(inhalt_markierung)
            '
            ' Vorbereitung Funktion ToZeile
            '
            ElseIf (pFunktion = FKT_TO_ZEILE) Then
                
                temp_long_1 = 1
                    
                If (pSelStart = 0) Then
                    
                    temp_long_2 = getAnzahlVorkommen(Mid(pString, 1, pSelLength), zeichen_zeilenumbruch) + 1
                
                Else
                    
                    temp_long_2 = getAnzahlVorkommen(Mid(pString, pSelStart, pSelLength), zeichen_zeilenumbruch) + 1
                
                End If
                
                If (m_toggle_mr_stringer_fkt) Then
                    
                    temp_string_1 = pOptString1
                
                Else
                    
                    temp_string_1 = LEER_STRING
                
                End If
            '
            ' Vorbereitung Funktion Debugausgabe
            '
            ElseIf (pFunktion = FKT_DEBUG_AUSGABE) Then
            
                m_zaehler_debug_print = m_zaehler_debug_print + 1
                
                If (m_zaehler_debug_print > 6) Then
                    
                    m_zaehler_debug_print = 1
                
                End If
            '
            ' Vorbereitung Funktion SetTrennzeichen
            '
            ElseIf (pFunktion = FKT_SET_TRENNZEICHEN_STR) Then
                
                knz_schleifen_durchlauf = True
                
                temp_string_3 = IIf(Len(pOptString1) = 0, IIf(m_toggle_mr_stringer_fkt, TRENN_STRING_6, TRENN_STRING_9), pOptString1)
                
                pFunktion = FKT_SET_TRENNZEICHEN

            ElseIf (pFunktion = FKT_SET_TRENNZEICHEN) Then
                
                knz_schleifen_durchlauf = ab_position >= 0
                
                temp_string_3 = IIf(m_toggle_mr_stringer_fkt, TRENN_STRING_6, TRENN_STRING_9)

            ElseIf (pFunktion = FKT_SET_TRENNZEICHEN_VOR) Then
                
                knz_schleifen_durchlauf = ab_position >= 0
                
                temp_string_3 = IIf(m_toggle_mr_stringer_fkt, TRENN_STRING_6, TRENN_STRING_9)
            
            ElseIf (pFunktion = FKT_SET_TRENNZEICHEN_ZURUECK) Then
                
                knz_schleifen_durchlauf = ab_position >= 0
                
                temp_string_3 = IIf(m_toggle_mr_stringer_fkt, TRENN_STRING_6, TRENN_STRING_9)

            ElseIf ((pFunktion = FKT_MARKIERE_VORNE_ODER_HINTEN) Or (pFunktion = FKT_MARKIERE_VORNE_UND_HINTEN) Or (pFunktion = FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE)) Then
                
                knz_schleifen_durchlauf = ab_position >= 0
                
            ElseIf (pFunktion = FKT_MARKIERE_CSV_VORNE_ODER_HINTEN) Then
                
                knz_schleifen_durchlauf = True
                
            ElseIf (pFunktion = FKT_MASKIERE_ANFZEICHEN) Then
                
                knz_schleifen_durchlauf = True

            ElseIf (pFunktion = FKT_MARKIERE_STR_VORNE_UND_HINTEN) Then
                
                knz_schleifen_durchlauf = True

            ElseIf ((pFunktion = FKT_MARKIERE_DOPPELTE_PLUS) Or (pFunktion = FKT_MARKIERE_DOPPELTE_PLUS_MINUS)) Then
                
                knz_schleifen_durchlauf = ab_position >= 0

            ElseIf (pFunktion = FKT_JAVA_GENERATOR) Then
    
                temp_string_1 = "pBuffer.append( """
                
                temp_string_2 = """ );"
            '
            ' Vorbereitung Funktion Clip
            '
            ElseIf ((pFunktion = FKT_CLIP_POSITION) Or (pFunktion = FKT_CLIP_GET_TEXT)) Then
                    
                If (pSelLength = Len(pString)) Or (knz_benutze_markierung = False) Then
                
                    ergebnis_fkt = pString
                    
                    knz_schleifen_durchlauf = False
    
                End If
            
            ElseIf (pFunktion = FKT_CALC_SUMME) Then
            
                temp_long_1 = 0
                temp_double_1 = 0
                temp_double_2 = 0
            '
            ' Vorbereitung Funktion Ausrichter
            '
            ' Fuer die Ausrichter-Funktion muss in einem ersten Durchlauf die maximale Ausdehnung
            ' des "Suchstrings" ermittelt werden. Der Suchstring ist dabei dass markierte Wort,
            ' oder der String aus dem Parameter pOptString. Die gefundene max. Position wird in
            ' der Variablen "temp_long_2" gespeichert.
            '
            ElseIf (pFunktion = FKT_AUSRICHTER_POSITION) Or (pFunktion = FKT_AUSRICHTER_STRING) Then
            
                If (pFunktion = FKT_AUSRICHTER_STRING) Then
                    
                    temp_string_1 = pOptString1
                    
                    pFunktion = FKT_AUSRICHTER_POSITION
                
                Else
                    
                    temp_string_1 = getStringAbBis(pString, pSelStart + 1, pSelStart + pSelLength)
                
                End If
            
                knz_schleifen_durchlauf = Len(temp_string_1) > 0
                
                If (knz_schleifen_durchlauf) Then
                    
                    zeilen_anzahl = cls_string_array.getAnzahlStrings
                    
                    temp_long_2 = 1
                    
                    zeilen_zaehler = 1
                    
                    While (zeilen_zaehler <= zeilen_anzahl)
                    
                        aktuelle_zeile = cls_string_array.getString(zeilen_zaehler)
                        
                        If (aktuelle_zeile <> LEER_STRING) Then
                            
                            temp_long_1 = InStr(aktuelle_zeile, temp_string_1)
                            
                            If (temp_long_1 > temp_long_2) Then
            
                                 temp_long_2 = temp_long_1
                                
                            End If
                        
                        End If
                    
                        zeilen_zaehler = zeilen_zaehler + 1
                    
                    Wend
                    '
                    ' Der max. Position wird noch 1 Zeichen hinzugezaehlt.
                    '
                    'temp_long_2 = temp_long_2 + 1
                    
                    '
                    ' In der Variablen "temp_string_3" wird ein String aus Leerzeichen mit der
                    ' Laenge der maximalen Ausdehnung gespeichert.
                    '
                    temp_string_3 = String(temp_long_2, " ") & "  "
    
                End If
            End If
            '
            ' Pruefung: Schleifendurchlauf machen?
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
                    ' sich ergebende String aus der Ab- bis Bis-Position gespeichert.
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
                    ' Ueber eine If-Else-Kaskade wird die auszufuerhende Aktion ermittelt.
                    ' Damit die Ausfuehrung ein bisschen beschleunigen wird, stehen die
                    ' oft genutzten Aktionen vorne und die selteneren weiter hinten in
                    ' der If-Kaskade.
                    '
                    If (pFunktion = FKT_GREP_WORT) Then
                        '
                        ' Funktion "Grep Wort"
                        '
                        ' Bedingung ist, dass die aktuelle Zeile kein Leerstring ist.
                        ' Aus einem Leerstring kann kein Wort rausgezogen werden!
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
                        
                            temp_string_2 = LEER_STRING
                            
                            If (pSelLength > 0) Then
                            
                                temp_string_2 = getGrepSuchwort(aktuelle_zeile, temp_string_1, zeichen_zeilenumbruch)
                            
                            Else
                            
                                temp_long_1 = getPosWortAnfang(aktuelle_zeile, ab_position)
                                
                                temp_long_2 = getPosWortende(aktuelle_zeile, ab_position)
                                
                                If (temp_long_2 > 0) And (temp_long_1 > 0) And (temp_long_2 > temp_long_1) Then
                                    
                                    temp_string_2 = Mid(aktuelle_zeile, temp_long_1, (temp_long_2 - (temp_long_1)) + 1)
                                
                                End If
                            
                            End If
                            
                            If (temp_string_2 <> LEER_STRING) Then
                            
                                If (zeilen_zaehler = 1) Then
                                    
                                    ergebnis_fkt = ergebnis_fkt & temp_string_2
                                
                                Else
                                    
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & temp_string_2
                                
                                End If
                                
                            End If
                            
                        End If
                 
                    ElseIf (pFunktion = FKT_AUSRICHTER_POSITION) Then
                        '
                        ' Funktion Ausrichter
                        ' In der aktuellen Zeile wird der Ausrichtsuchbegriff gesucht.
                        ' Anders: es wird die Positon des Suchbegriffs gesucht.
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
                        If (temp_long_1 > 0) Then
                        
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
                        ' Ist das Suchwort enthalten, wird die aktuelle Zeile nur aufgenommen, wenn es sich um
                        ' die Funktion Grep+ handelt.
                        '
                        ' Ist das Suchwort nicht enthalten, wird die aktuelle Zeile nur aufgenommen, wenn es sich
                        ' um die Funktion Grep- handelt.
                        '
                        If (InStr(aktuelle_zeile, temp_string_1) > 0) Then
                            
                            If (pKennzeichen1) Then
                                
                                ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & aktuelle_zeile
                            
                            End If
                            
                        Else
                        
                            If (pKennzeichen1 = False) Then
                                
                                ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & aktuelle_zeile
                            
                            End If
                        
                        End If
                    
                    ElseIf (pFunktion = FKT_GREP_DUPLIZIERE_MARKZEILEN) Then
                    
                        temp_long_3 = 0
                        
                        If (InStr(aktuelle_zeile, temp_string_1) > 0) Then
                        
                            'If (pKennzeichen1) Then
                            
                                ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & temp_string_2 & aktuelle_zeile & temp_string_3 & aktuelle_zeile
                                
                                temp_long_3 = 1
                                
                            'End If
                        
                        End If
                        
                        If (temp_long_3 = 0) Then
                            
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & aktuelle_zeile
                        
                        End If
                    
                    ElseIf (pFunktion = FKT_GREP_MARK) Then

                        temp_long_3 = 0
                        
                        If (InStr(aktuelle_zeile, temp_string_1) > 0) Then
                        
                            If (pKennzeichen1) Then
                            
                                ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & temp_string_2 & aktuelle_zeile & temp_string_3
                                
                                temp_long_3 = 1
                                
                            End If
                            
                        Else
                        
                            If (pKennzeichen1 = False) Then
                            
                                ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & temp_string_2 & aktuelle_zeile & temp_string_3
                                
                                temp_long_3 = 1
                            
                            End If
                        
                        End If
                        
                        If (temp_long_3 = 0) Then
                            
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & aktuelle_zeile
                        
                        End If
                        
                    
                    ElseIf (pFunktion = FKT_MARKIERE_CSV_VORNE_ODER_HINTEN) Then
                        '
                        ' Funktion "CSV Markiere vorne oder hinten"
                        '
                        ' Die aktuelle Zeile wird abwechselnd vorne oder hinten mit
                        ' dem uebergebenen String im Parameter "pOptString1" versehen.
                        ' Soll die Markierung benutzt werden, wird in der aktuellen
                        ' Zeile der temp_string 1 ersetzt.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            temp_string_1 = pOptString1 & akt_zeile_mark
                        
                        Else
                        
                            temp_string_1 = akt_zeile_mark & pOptString1
                            
                        End If
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_1 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_1)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_1)
                        
                    ElseIf (pFunktion = FKT_MARKIERE_VORNE_ODER_HINTEN) Then
                        '
                        ' Funktion "Markiere vorne oder hinten"
                        '
                        ' Die aktuelle Zeile wird abwechselnd vorne oder hinten mit
                        ' einem Suchstring versehen. Soll die Markierung benutzt werden,
                        ' wird in der aktuellen Zeile der temp_string 1 ersetzt.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            temp_string_1 = TRENN_STRING_7 & akt_zeile_mark
                        
                        Else
                        
                            temp_string_1 = akt_zeile_mark & TRENN_STRING_7
                            
                        End If
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_1 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_1)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_1)
                        
                    ElseIf (pFunktion = FKT_MARKIERE_VORNE_UND_HINTEN) Then
                        '
                        ' Funktion "Markiere vorne und hinten"
                        '
                        ' Es wird vorne und hinten ein Suchzeichen gesezt.
                        ' Das kann auf die gesamte Zeile oder aber nur auf den Markierungsbereich erfolgen.
                        '
                        temp_string_1 = TRENN_STRING_7 & akt_zeile_mark & TRENN_STRING_8
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_1 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_1)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_1)

                    ElseIf (pFunktion = FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE) Then
                        
                        temp_string_1 = TRENN_STRING_6 & akt_zeile_mark & TRENN_STRING_7 & akt_zeile_mark & TRENN_STRING_8
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_1 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_1)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_1)
                    
                    ElseIf (pFunktion = FKT_MARKIERE_STR_VORNE_UND_HINTEN) Then
                        '
                        ' Funktion "Markiere vorne und hinten mit String"
                        '
                        ' Die aktuelle Zeile wird vorne mit dem Optstring 1 und hinten mit
                        ' dem Optstring 2 versehen. Eine Verwendung der Markierung ist
                        ' nicht vorhanden.
                        '
                        Call cls_string_array.setString(zeilen_zaehler, pOptString1 & akt_zeile_mark & pOptString2)
        
                    ElseIf (pFunktion = FKT_TRIM) Then
                        '
                        ' Funktion "Trim"
                        '
                        ' Auf jede Zeile wird ein Trim ausgefuehrt.
                        '
                        temp_string_1 = Trim(akt_zeile_mark)
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_1 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_1)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_1)
                    
                    ElseIf (pFunktion = FKT_MASKIERE_ANFZEICHEN) Then
                        '
                        ' Funktion "Maskiere Anfuehrungszeichen"
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            temp_string_1 = """"
                            temp_string_2 = "\"""
                            
                        Else
                        
                            temp_string_1 = """"
                            temp_string_2 = """"""
                        
                        End If
                        
                        temp_string_3 = Replace(akt_zeile_mark, temp_string_1, temp_string_2)
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_3 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_3)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_3)

                    
                    ElseIf (pFunktion = FKT_UCASE_LCASE) Then
                        '
                        ' Funktion Upper- Lower-Case
                        ' Die Funktionen "UCase" bzw. "LCase" werden auf die aktuelle Zeile oder
                        ' den Inhalt der Makierung ausgefuehrt.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                            
                            temp_string_1 = UCase(akt_zeile_mark)
                        
                        Else
                            
                            temp_string_1 = LCase(akt_zeile_mark)
                        
                        End If
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_1 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_1)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_1)

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
                                
                                If (temp_long_2 > 0) Then
                                
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
                            If (temp_long_2 > 0) Then
                            
                                temp_string_1 = Left(aktuelle_zeile, temp_long_2 - 1)
                                
                                temp_string_2 = Mid(aktuelle_zeile, temp_long_2, Len(aktuelle_zeile))
                            
                            End If
                                
                            Call cls_string_array.setString(zeilen_zaehler, temp_string_1 & temp_string_3 & temp_string_2)
                        
                        End If
                    
                    ElseIf (pFunktion = FKT_CLIP_POSITION) Then
                        '
                        ' Funktion "Clip"
                        '
                        ' 1. Entferne den selektierten Bereich
                        ' 2. lass nur den selektierten Bereich stehen
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            Call cls_string_array.setString(zeilen_zaehler, getRemoveAbBis(aktuelle_zeile, ab_position, bis_position))
                        
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, getStringAbBis(aktuelle_zeile, ab_position, bis_position))
                        
                        End If
                    
                    ElseIf (pFunktion = FKT_CLIP_GET_TEXT) Then
                        
                        Call cls_string_array.setString(zeilen_zaehler, getStringAbBis(aktuelle_zeile, ab_position, bis_position))
                    
                    ElseIf (pFunktion = FKT_SPLIT) Then
                        '
                        ' Funktion "Split"
                        '
                        ' Zerteilt die Zeile anhand einer Posittion oder Markierung.
                        ' Eine Zeile kann nur dann gesplittet werden, wenn diese kein Leerstring ist.
                        '
                        ' Durch den Wert in "temp_long_1" wird die Position vorgegeben. Wurde noch
                        ' eine Markierung vorgegeben, wird die Zeichenkette der Markierung gesucht.
                        ' Wird die Markierung nicht gefunden, wird "temp_long_1" zu 0.
                        '
                        ' Durch eine weitere Pruefung, wird der Wert in "temp_long_1" auf groesser 0 geprueft.
                        ' Wenn dem so ist, wird die Zeile gesplittet.
                        '
                        If (aktuelle_zeile <> LEER_STRING) Then
                        
                            If (pSelLength > 0) Then
    
                                temp_long_1 = InStr(aktuelle_zeile, temp_string_1)
                                
                            End If
    
                            If (temp_long_1 > 0) Then

                                If (m_toggle_mr_stringer_fkt) Then
                                    
                                    Call cls_string_array.setString(zeilen_zaehler, Left(aktuelle_zeile, temp_long_1 - 1))
                                
                                Else
                                    
                                    Call cls_string_array.setString(zeilen_zaehler, Mid(aktuelle_zeile, temp_long_1 + temp_long_2, Len(aktuelle_zeile)))
                                
                                End If

                            End If

                        End If
                        
                    ElseIf (pFunktion = FKT_CHECK_LEERSTRING) Then

                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            aktuelle_zeile = Replace(Trim(akt_zeile_mark), """", temp_string_3)
                            
                        End If
                        
                        If (Trim(aktuelle_zeile) <> LEER_STRING) Then
        
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "if ( " & aktuelle_zeile & " = """" ) then "
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "    '##sss( ""Fehler: " & aktuelle_zeile & " nicht gesetzt"" )"
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "    'fkt_ergebnis = false"
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "End If"
                        
                        End If

                    ElseIf (pFunktion = FKT_STRING_IT) Then
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
                
                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            If (m_toggle_mr_stringer_fkt) Then
                            
                               temp_string_1 = temp_string_1 & aktuelle_zeile
                            
                            Else
                            
                               temp_string_1 = temp_string_1 & " " & aktuelle_zeile
                            
                            End If
                            
                            temp_long_2 = temp_long_2 + 1
                            
                        End If
                
                        If (temp_long_2 = temp_long_1) Then
                        
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & temp_string_1
                            
                            temp_long_2 = 0
                            
                            temp_string_1 = LEER_STRING
                            
                        End If

                    ElseIf (pFunktion = FKT_CAMEL_CASE) Then
                        '
                        ' Funktion Upper-CamelCase
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            temp_string_1 = getKlartext(akt_zeile_mark, LEER_STRING, ",.-()[]""")
                        
                        Else
                        
                            temp_string_1 = getKlartext(akt_zeile_mark, "_", ",.-()[]""")
                        
                        End If
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_long_1 = Len(akt_zeile_mark) - Len(temp_string_1) ' Ergebnis kann kuerzer werden
                            
                            If (temp_long_1 > 0) Then
                            
                                temp_string_2 = String(temp_long_1, " ")
                                
                            End If
                           
                            temp_string_1 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_1 & temp_string_2)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_1)
                    
                    ElseIf (pFunktion = FKT_DEKLARATION) Then
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
                            
                                ergebnis_fkt = ergebnis_fkt & "String " & akt_zeile_mark & " = null;" & zeichen_zeilenumbruch
                                
                            Else
                            
                                ergebnis_fkt = ergebnis_fkt & "Dim " & akt_zeile_mark & " As String" & zeichen_zeilenumbruch
                                
                            End If
                            
                        End If
    
                    ElseIf (pFunktion = FKT_SET_NULL) Then
                        '
                        ' Funktion "Variablen auf null stellen"
                        ' Aus der aktuellen Zeile oder der Markierung wird abwechselnd fuer
                        ' Java und Visual-Basic eine Anweisung fuer die Null-Setzung erstellt.
                        '
                        ' Java =  xxx = null;
                        ' VB   =  set xxx = nothing
                        '
                        If (Trim(akt_zeile_mark) <> LEER_STRING) Then
                        
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                ergebnis_fkt = ergebnis_fkt & akt_zeile_mark & " = null;" & zeichen_zeilenumbruch
                                
                            Else
                            
                                ergebnis_fkt = ergebnis_fkt & "set " & akt_zeile_mark & " = nothing" & zeichen_zeilenumbruch
                                
                            End If
                            
                        End If
                        
                    ElseIf (pFunktion = FKT_GROUP_NACH_STRING) Then
                        '
                        ' Funktion "Group nach String"
                        '
                        If (Trim(akt_zeile_mark) <> LEER_STRING) Then
                        
                            If (akt_zeile_mark <> temp_string_1) Then
                            
                                ergebnis_fkt = ergebnis_fkt & vbCrLf
                                
                                temp_string_1 = akt_zeile_mark
                            
                            End If
                            
                            ergebnis_fkt = ergebnis_fkt & vbCrLf & aktuelle_zeile
                            
                        End If
    
                    ElseIf (pFunktion = FKT_CMD_RENAME) Then
                        '
                        ' Funktion "Rename"
                        ' Erstellung eines Rename-Kommando-Aufrufes fuer BAT-Dateien, wobei der
                        ' neue Dateiname mittels der Funktion "renameDatei" schon vorumgewandelt
                        ' wird.
                        '
                        ' Bedingung ist, dass die aktuelle Zeile ungleich einem Leerstring ist.
                        ' Soll eine Markierung benutzt werden, wird der Dateiname aus der
                        ' Markierung genommen, ansonsten wird die aktuelle Zeile genommen.
                        '
                        If (akt_zeile_mark <> LEER_STRING) Then
                            
                            temp_string_2 = renameDatei(Trim(akt_zeile_mark))
                            
                            temp_string_2 = "rename """ & akt_zeile_mark & """" & TRENN_STRING_7 & temp_string_2 & """"
                            
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                temp_string_2 = temp_string_2 & TRENN_STRING_8
                                
                            End If
    
                            Call cls_string_array.setString(zeilen_zaehler, temp_string_2)
                            
                        End If
                    
                    ElseIf (pFunktion = FKT_DEBUG_AUSGABE) Then
                        '
                        ' Funktion "Debug-Ausgabe"
                        ' Erstellt fuer die aktuelle Zeile oder die Markierung eine Debug-Ausgabe fuer
                        ' VB, PHP und Java.
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
                            
                                Call cls_string_array.setString(zeilen_zaehler, "DrLogger.wl( """ & Replace(akt_zeile_mark, """", "\""") & " =>"" + " & Replace(akt_zeile_mark, """", LEER_STRING) & " + ""<"" );")
                            
                            End If
                            
                        End If
                        
                    ElseIf (pFunktion = FKT_ERSTELLE_XML) Or (pFunktion = FKT_ERSTELLE_XML_2) Then
                        '
                        ' Funktion "XML-Erstellung"
                        ' Die aktuelle Zeile oder Markierung wird als Tag-Namen betrachtet.
                        ' Dabei wird der TAG-Name in Grossbuchstaben gewandelt und ein XML-TAG
                        ' erstellt. Dieses einmal in einer Klammer oder mit Start- und End-Tag.
                        '
                        akt_zeile_mark = Trim(akt_zeile_mark)
                        
                        If (akt_zeile_mark <> "") Then
                            
                            If (pFunktion = FKT_ERSTELLE_XML_2) Then

                                ergebnis_fkt = ergebnis_fkt & "<" & UCase(getKlartext(akt_zeile_mark, "_")) & " x_attribut=""" & akt_zeile_mark & """ /> " & zeichen_zeilenumbruch
                            
                            Else
                            
                                akt_zeile_mark = UCase(getKlartext(akt_zeile_mark, "_"))
                                
                                ergebnis_fkt = ergebnis_fkt & temp_string_1 & akt_zeile_mark & temp_string_2
                                
                                If (temp_string_3 <> "") Then
                                
                                    ergebnis_fkt = ergebnis_fkt & akt_zeile_mark & temp_string_3
                                
                                End If
                                
                                ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch
                                
                            End If
                        
                        End If
                        
                    ElseIf (pFunktion = FKT_SINGLETON_JAVA) Or (pFunktion = FKT_GETTER_SETTER_JAVA) Or (pFunktion = FKT_GETTER_SETTER_VB) Or (pFunktion = FKT_GETTER_SETTER_JAVA_SCRIPT) Then
                        '
                        ' Funktion "Getter Setter Java" oder "Getter Setter VB"
                        ' Die aktuelle Zeile oder die Markierung wird als Variablen-Name gesehen.
                        ' Wird in der Zeichenkette ein "=" gefunden, wird der erste Teil als
                        ' Name gesehen und der Zweite Teil als Typ.
                        '
                        ' Ist keine Typinformation vorhanden, wird als Typ "String" genommen.
                        '
                        akt_zeile_mark = Trim(akt_zeile_mark)
                        
                        If (akt_zeile_mark <> "") Then
                           
                            Dim str_var_typ As String
                            
                            temp_long_1 = InStr(akt_zeile_mark, "=")
                            
                            If (temp_long_1 <= 0) Then
                            
                                temp_long_1 = InStr(akt_zeile_mark, " As ")
                                
                                temp_long_3 = 4
                                
                            Else
                            
                                temp_long_3 = 1
                            
                            End If
                        
                            If (temp_long_1 > 0) Then
                            
                                temp_long_2 = temp_long_1 + temp_long_3
                                
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    temp_string_1 = "m_" & LCase(getKlartext(Trim(Left(akt_zeile_mark, temp_long_1 - 1)), "_"))
                                
                                Else
                                
                                    temp_string_1 = "lv_" & LCase(getKlartext(Trim(Left(akt_zeile_mark, temp_long_1 - 1)), "_"))
                                
                                End If
                                
                                temp_string_2 = getKlartext(Trim(Left(akt_zeile_mark, temp_long_1 - 1)), LEER_STRING)
                                
                                str_var_typ = Trim(Mid(akt_zeile_mark, temp_long_2, Len(akt_zeile_mark)))
                            
                            Else
                            
                                str_var_typ = "String" ' " & str_var_typ  & "
                            
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    temp_string_1 = "m_" & LCase(getKlartext(akt_zeile_mark, "_"))  ' member-Variable " & temp_string_1 & "
                                
                                Else
                                
                                    temp_string_1 = "lokale_var_" & LCase(getKlartext(akt_zeile_mark, "_"))  ' member-Variable " & temp_string_1 & "
                                
                                End If
                                
                                temp_string_2 = getKlartext(akt_zeile_mark, LEER_STRING) ' CamelCase-Grundname " & temp_string_2 & "

                            End If
                            
                            If (pFunktion = FKT_SINGLETON_JAVA) Then

                                temp_string_3 = temp_string_3 & zeichen_zeilenumbruch & IIf(m_toggle_mr_stringer_fkt, "private ", LEER_STRING) & str_var_typ & " " & temp_string_1 & " = "
                                
                                temp_long_1 = 0
                                
                                If (LCase(str_var_typ) = "boolean") Then
                                
                                    temp_string_3 = temp_string_3 & "false; // true;"
                                    
                                ElseIf (LCase(str_var_typ) = "long") Then
                                
                                    temp_string_3 = temp_string_3 & "0;"
                                    
                                ElseIf (LCase(str_var_typ) = "double") Then
                                
                                    temp_string_3 = temp_string_3 & "0.0d;"
                                    
                                Else
                                
                                    temp_string_3 = temp_string_3 & "null;"
                                    
                                    temp_long_1 = 1
                                
                                End If
                                
                                ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                
                                
                                If (temp_long_1 = 0) Then
                                    
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "public " & str_var_typ & " get" & temp_string_2 & "() { return " & temp_string_1 & "; }"
                                
                                Else
                                    
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "public " & str_var_typ & " get" & temp_string_2 & "() { if ( " & temp_string_1 & " == null ) { " & temp_string_1 & " = new " & str_var_typ & "(); } return " & temp_string_1 & "; }"
                                
                                End If
                                
                            ElseIf (pFunktion = FKT_GETTER_SETTER_VB) Then
                            
                                temp_string_3 = temp_string_3 & zeichen_zeilenumbruch & IIf(m_toggle_mr_stringer_fkt, "Private ", "Dim ") & temp_string_1 & " As " & str_var_typ & " ' = "
                                
                                If (LCase(str_var_typ) = "boolean") Then
                                
                                    temp_string_3 = temp_string_3 & "false ' true"
                                    
                                ElseIf (LCase(str_var_typ) = "bigdecimal") Then
                                
                                    temp_string_3 = temp_string_3 & " 0.0"
                                
                                ElseIf (LCase(str_var_typ) = "long") Then
                                
                                    temp_string_3 = temp_string_3 & "0"
                                
                                ElseIf (LCase(str_var_typ) = "double") Then
                                
                                    temp_string_3 = temp_string_3 & "0.0"
                                    
                                Else
                                
                                    temp_string_3 = temp_string_3 & """"""
                            
                                End If
                                
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "'######## VB #############"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "'"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "Public Function get" & temp_string_2 & "() As " & str_var_typ
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & zeichen_zeilenumbruch & "    get" & temp_string_2 & " = " & temp_string_1
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & zeichen_zeilenumbruch & "End Function"
                                    
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "'######## VB #############"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "'"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "Public Sub set" & temp_string_2 & "( p" & temp_string_2 & " As " & str_var_typ & " )"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & zeichen_zeilenumbruch & "    " & temp_string_1 & " = p" & temp_string_2
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & zeichen_zeilenumbruch & "End Sub"
                                
                                Else
                                 
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & temp_string_1 & " = inst_objekt.get" & temp_string_2 & "()"
                                    
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "inst_objekt.set" & temp_string_2 & "( " & temp_string_1 & " )"
                               
                                End If
                                
                            ElseIf (pFunktion = FKT_GETTER_SETTER_JAVA_SCRIPT) Then
                                
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                    
                                    If (temp_long_1 > 0) Then ' explizite Typangabe --> dann Initialisierung mit undefined
                                    
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  " & temp_string_1 & " : undefined,"
                                        
                                    Else
                                    
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  " & temp_string_1 & " : """","
                                        
                                    End If

                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "get" & temp_string_2 & " : function()"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "{"
                                    
                                    If (temp_long_1 > 0) Then ' explizite Typangabe --> dann Singletonpattern hinzufuegen
                                    
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  if ( this." & temp_string_1 & " == undefined )"
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  {"
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "    this." & temp_string_1 & " = new " & str_var_typ & "();"
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  }"
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                        
                                    End If
                                    
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  return this." & temp_string_1 & ";"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "},"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "set" & temp_string_2 & " : function( p" & temp_string_2 & " )"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "{"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  this." & temp_string_1 & " = p" & temp_string_2 & ";"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "},"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                
                                Else
                                
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                    
                                    If (temp_long_1 > 0) Then ' explizite Typangabe --> dann Initialisierung mit undefined
                                    
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  this." & temp_string_1 & " = undefined;"
                                        
                                    Else
                                    
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  this." & temp_string_1 & " = """";"
                                        
                                    End If
                                    
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "Bean.prototype.get" & temp_string_2 & " = function()"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "{"
                                    
                                    If (temp_long_1 > 0) Then ' explizite Typangabe --> dann Singletonpattern hinzufuegen
                                    
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  if ( this." & temp_string_1 & " == undefined )"
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  {"
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "    this." & temp_string_1 & " = new " & str_var_typ & "();"
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  }"
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                        
                                    End If
                                    
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  return this." & temp_string_1 & ";"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "}"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "Bean.prototype.set" & temp_string_2 & " = function( p" & temp_string_2 & " )"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "{"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "  this." & temp_string_1 & " = p" & temp_string_2 & ";"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "}"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                    
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

                                ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                
                                If (m_toggle_mr_stringer_fkt) Then

                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "public " & str_var_typ & " get" & temp_string_2 & "() { return " & temp_string_1 & "; }"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "public void set" & temp_string_2 & "( " & str_var_typ & " p" & temp_string_2 & " ) { " & temp_string_1 & " = p" & temp_string_2 & "; }"
                                
                                Else
                                
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & temp_string_1 & " = inst_objekt.get" & temp_string_2 & "();"
                                    
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "inst_objekt.set" & temp_string_2 & "( " & temp_string_1 & " );"
                                
                                End If
                                
                                
                                ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & ""
                                
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
                        
                        If (akt_zeile_mark <> "") Then
                            
                            temp_long_1 = temp_long_1 + 1
                            
                            temp_double_1 = Val(getzahl(akt_zeile_mark, 4, False))
                            
                            temp_double_2 = temp_double_2 + temp_double_1
                        
                            ergebnis_fkt = ergebnis_fkt & "OK     >" & temp_long_1 & "< >" & akt_zeile_mark & "< >" & temp_double_1 & "< >" & temp_double_2 & "<" & zeichen_zeilenumbruch
                        
                        Else
                        
                            temp_double_1 = 0
                            
                            ergebnis_fkt = ergebnis_fkt & "FEHLER >" & temp_long_1 & "< >" & akt_zeile_mark & "< >" & temp_double_1 & "< >" & temp_double_2 & "<" & zeichen_zeilenumbruch

                        End If
                           
                    ElseIf (pFunktion = FKT_ERSTELLE_NAMEN) Then
                        '
                        ' Funktion "Erstelle Namen"
                        '
                        ' Erstellt aus der aktuellen Zeile oder der Makierung Variablennamen.
                        '
                        ' Einmal mit keinem Trennzeichen (=Camelcase) und einmal mit einem
                        ' Unterstrich als Trennzeichen.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                            
                            temp_string_1 = getKlartext(akt_zeile_mark, "", " ")
                        
                        Else
                            
                            temp_string_1 = LCase(getKlartext(akt_zeile_mark, "_", " "))
                        
                        End If
                        
                        If (knz_benutze_markierung) Then
                        
                            temp_string_1 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_1)
                            
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_1)
                        
                    ElseIf (pFunktion = FKT_MARKIERE_DOPPELTE_PLUS) Then
                        
                        If (akt_zeile_mark = temp_string_1) Then
                        
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                Call cls_string_array.setString(zeilen_zaehler, aktuelle_zeile & MARKIERUNG_DOPPELTE_VORKOMMEN)
                                
                            Else
                            
                                Call cls_string_array.setString(zeilen_zaehler, MARKIERUNG_DOPPELTE_VORKOMMEN & aktuelle_zeile)
                                
                            End If
                        
                        Else
                            
                            Call cls_string_array.setString(zeilen_zaehler, aktuelle_zeile)
                        
                        End If
                        
                        temp_string_1 = akt_zeile_mark
                    
                    ElseIf (pFunktion = FKT_MARKIERE_DOPPELTE_PLUS_MINUS) Then
                        
                        If (akt_zeile_mark = temp_string_1) Then
                        
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                Call cls_string_array.setString(zeilen_zaehler, aktuelle_zeile & MARKIERUNG_DOPPELTE_VORKOMMEN)
                                
                            Else
                            
                                Call cls_string_array.setString(zeilen_zaehler, MARKIERUNG_DOPPELTE_VORKOMMEN & aktuelle_zeile)
                            
                            End If
                            
                            temp_string_1 = cls_string_array.getString(zeilen_zaehler - 1)
                           
                            If (InStr(temp_string_1, MARKIERUNG_DOPPELTE_VORKOMMEN) = 0) Then
                           
                                If (m_toggle_mr_stringer_fkt) Then
                                    
                                    Call cls_string_array.setString(zeilen_zaehler - 1, temp_string_1 & MARKIERUNG_DOPPELTE_VORKOMMEN)
                                    
                                Else

                                    Call cls_string_array.setString(zeilen_zaehler - 1, MARKIERUNG_DOPPELTE_VORKOMMEN & temp_string_1)
                                    
                                End If
                           
                           End If
                           
                        Else
                            
                            Call cls_string_array.setString(zeilen_zaehler, aktuelle_zeile)
                        
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
                        
                            Call cls_string_array.setString(zeilen_zaehler, "" & pKennzeichen1)
                            
                        End If
                        
                    ElseIf (pFunktion = FKT_ZAEHLER) Then
                        '
                        ' Funktion "Zaehler"
                        ' Zaehlt die Zeilen der Eingabe, bzw. jede Zeile bekommt eine
                        ' Zeilennummer im Ergebnis.
                        '
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            Call cls_string_array.setString(zeilen_zaehler, "" & zeilen_zaehler)
                            
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, Right(NULL_ZIFFERN_100 & zeilen_zaehler, temp_long_1))
                            
                        End If
                        
                    ElseIf (pFunktion = FKT_MAKE_LONG_DATUM) Then
                    
                        temp_string_1 = Mid(aktuelle_zeile, ab_position + 6, 4) & Mid(aktuelle_zeile, ab_position + 3, 2) & Mid(aktuelle_zeile, ab_position, 2)
                        
                        Call cls_string_array.setString(zeilen_zaehler, replaceSubstringAbBis(aktuelle_zeile, ab_position, ab_position + 9, temp_string_1))
                    
                    ElseIf (pFunktion = FKT_ENTFERNE_LEERZEILEN) Then
                    
                        If (Len(Trim(aktuelle_zeile)) > 0) Then
                        
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & aktuelle_zeile

                        End If
    
                    ElseIf (pFunktion = FKT_LEERZEILEN_EINFUEGEN) Then
                    
                        temp_long_2 = Len(Trim(aktuelle_zeile))
                        '
                        ' Uebernehme nur gesetzte Zeilen
                        ' Vor jeder gesetzten Zeile kommt ein Zeilenumbruchszeichen hinzu
                        '
                        If (temp_long_2 > 0) Then
                        
                            If (temp_long_1 = 1) Then
                            
                                'If (temp_long_3 > 1) Then
                                
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch
                                    
                                'End If
                            
                            End If
                        
                            ergebnis_fkt = ergebnis_fkt & aktuelle_zeile & zeichen_zeilenumbruch
                            
                            temp_long_1 = 1
                            
                            temp_long_3 = temp_long_2
                        End If
                        
                    ElseIf (pFunktion = FKT_GET_UNIQUE) Or (pFunktion = FKT_GET_DOPPELTE_VORKOMMEN) Or (pFunktion = FKT_GET_EINMALIGE_VORKOMMEN) Then
                        
                        temp_string_1 = Trim(akt_zeile_mark)
                        
                        If (temp_string_1 <> "") Then
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
                                If (InStr(temp_string_2, temp_string_1) <= 0) Then
                                
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & aktuelle_zeile
                                    
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
                                    ' Ist der aktuelle String aus "temp_string_1" nicht in t3 vorhanden,
                                    ' wird "temp_string_1" in t3 aufgenommen.
                                    '
                                    ' Gleichzeitig wird die aktuelle Zeile in das Funktionsergebnis aufgenommen.
                                    '
                                    If (InStr(temp_string_3, temp_string_1) <= 0) Then
                                                           
                                       temp_string_3 = temp_string_3 & temp_string_1
                                       
                                       ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & aktuelle_zeile
            
                                    End If
                                
                                End If
                            
                            ElseIf (pFunktion = FKT_GET_EINMALIGE_VORKOMMEN) Then
                            
                                If (InStr(temp_string_2, temp_string_1) <= 0) Then
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
                        
                    ElseIf (pFunktion = FKT_TO_ZEILE) Then
                    
                        ergebnis_fkt = ergebnis_fkt & akt_zeile_mark & temp_string_1
                        
                        If (temp_long_1 = temp_long_2) Then ' Zaehler gleich
                        
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch
                            
                            temp_long_1 = 1 ' Zaehler zurueckstellen
                        
                        Else
                        
                            temp_long_1 = temp_long_1 + 1
                        
                        End If
                    
                    ElseIf (pFunktion = FKT_ERSTELLE_BLOCK) Then
                    
                        temp_string_1 = aktuelle_zeile
                        
                        temp_long_2 = 1
                        
                        temp_string_2 = Mid(aktuelle_zeile, temp_long_2, temp_long_1)
                        
                        While (temp_string_2 <> "")
                        
                            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & temp_string_2
                             
                            temp_long_2 = temp_long_2 + temp_long_1
                        
                            temp_string_2 = Mid(aktuelle_zeile, temp_long_2, temp_long_1)
                            
                        Wend
    
                    ElseIf ((pFunktion = FKT_CSV_SWAP) Or (pFunktion = FKT_CSV_CASE) Or (pFunktion = FKT_CSV_JAVA_PROP)) Then
                    
                        temp_long_1 = InStr(aktuelle_zeile, pOptString1)
                    
                        If (temp_long_1 > 0) Then
                        
                            temp_long_2 = temp_long_1 + Len(pOptString1)
                            
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                temp_string_1 = Left(aktuelle_zeile, temp_long_1 - 1)
                                
                                temp_string_2 = Mid(aktuelle_zeile, temp_long_2, Len(aktuelle_zeile))
                                
                            Else
                            
                                temp_string_2 = Left(aktuelle_zeile, temp_long_1 - 1)
                                
                                temp_string_1 = Mid(aktuelle_zeile, temp_long_2, Len(aktuelle_zeile))
                                
                            End If
                      
                            If (pFunktion = FKT_CSV_CASE) Then
                            
                                Call cls_string_array.setString(zeilen_zaehler, "case " & temp_string_1 & " : { " & temp_string_2 & " break; }")
                      
                            ElseIf (pFunktion = FKT_CSV_JAVA_PROP) Then
                            
                                temp_string_1 = Trim(temp_string_1)
                                temp_string_2 = Trim(temp_string_2)
                                
                                Call cls_string_array.setString(zeilen_zaehler, "" & STR_VAR_NAME_PROPERTIES_LOKAL & ".setProperty( " & temp_string_1 & ", " & AUSRICHT_STRING_TEMP_1 & temp_string_2 & AUSRICHT_STRING_TEMP_2 & " );")
                            
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
                
                            ergebnis_fkt = ergebnis_fkt & pOptString1
                    
                        End If
                        
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            ergebnis_fkt = ergebnis_fkt & """" & akt_zeile_mark & """"
                            
                        Else
                        
                            ergebnis_fkt = ergebnis_fkt & akt_zeile_mark
                        
                        End If
                     
                    ElseIf (pFunktion = FKT_TRIM_X) Then
                        '
                        ' Funktion "Trim X"
                        ' TrimX eliminiert doppelte Leerzeichen durch den gesamten String hindurch.
                        '
                        If (knz_benutze_markierung) Then
                        
                            temp_string_1 = trimX(akt_zeile_mark)
                            
                            Call cls_string_array.setString(zeilen_zaehler, replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_1))
                        
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, trimX(akt_zeile_mark))
                        
                        End If
                        
                    ElseIf (pFunktion = FKT_BLOCK_ZUFALL) Then
                    
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
                        
                            Call cls_string_array.setString(zeilen_zaehler, akt_zeile_mark & "#4" & aktuelle_zeile)
                            
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, aktuelle_zeile & "#4" & akt_zeile_mark)
                            
                        End If
    
                    ElseIf (pFunktion = FKT_VERSCHIEBEN) Then
                        '
                        ' Funktion "Verschieben"
                        '
                        ' Verschieben eines Teilbereiches aus der aktuellen Zeile, einmal nach vorne
                        ' und einmal ans Ende der aktuellen Zeile.
                        '
                        temp_string_2 = getRemoveAbBis(aktuelle_zeile, ab_position, bis_position)
                
                        If (m_toggle_mr_stringer_fkt) Then
                        
                            Call cls_string_array.setString(zeilen_zaehler, akt_zeile_mark & "#5" & temp_string_2)
                            
                        Else
                        
                            Call cls_string_array.setString(zeilen_zaehler, temp_string_2 & "#5" & akt_zeile_mark)
                            
                        End If
                    
                    ElseIf ((pFunktion = FKT_JSON_LESEN_SCHREIBEN) Or (pFunktion = FKT_NOTES_LESEN_SCHREIBEN) Or (pFunktion = FKT_JAVA_XML_WRITER_STRING) Or (pFunktion = FKT_JAVA_XML_WRITER_NUMMER)) Then
                    
                        If (Trim(aktuelle_zeile) <> "") Then
                        
                            temp_string_1 = Trim(akt_zeile_mark)
                            
                            If (temp_string_1 <> "") Then
    
                                If (pFunktion = FKT_JSON_LESEN_SCHREIBEN) Then
                                
                                    If (m_toggle_mr_stringer_fkt) Then
                                    
                                        temp_string_2 = "json_erg += '\n" & temp_string_1 & " >' + AjaxErg." & temp_string_1 & " + '<';"
                                        
                                    Else
                                    
                                        temp_string_2 = "json_string += FkJson.getStringJson( """ & temp_string_1 & """, " & temp_string_1 & " ) + "","";" ' + "\"","";"
                                        
                                    End If
                                    
                                ElseIf (pFunktion = FKT_NOTES_LESEN_SCHREIBEN) Then
                                
                                    If (m_toggle_mr_stringer_fkt) Then
                                    
                                        temp_string_2 = "Call notesDokumentStringSet( notes_dokument, """ & temp_string_1 & """, " & temp_string_1 & " )"
                                        
                                    Else
                                    '"m_" & LCase(getKlartext(Trim(temp_string_1), "_"))
                                    
                                        temp_string_2 = LCase(getKlartext(Trim(temp_string_1), "_")) & " = notesDokumentStringGet( notes_dokument, """ & temp_string_1 & """ )"
                                        
                                    End If
                                    
                                ElseIf (pFunktion = FKT_JAVA_XML_WRITER_STRING) Then
                
                                    temp_string_1 = UCase(getKlartext(temp_string_1, "_"))
                                    
                                    temp_string_3 = "#Xp" + getKlartext(temp_string_1, "")
                                
                                    If (m_toggle_mr_stringer_fkt) Then
                                    
                                        temp_string_2 = "xml_string.append( FkXml.getXmlTag( TAG_" & temp_string_1 & ", " & temp_string_3 & ", TAG_VORGABE_" & temp_string_1 & " );"
                                        
                                    Else
                                    
                                        temp_string_2 = "pBuffer.append( ""<" & temp_string_1 & ">"" + " & temp_string_3 & " + ""</" & temp_string_1 & ">"" );"
                                        
                                    End If
                                
                                ElseIf (pFunktion = FKT_JAVA_XML_WRITER_NUMMER) Then
                                
                                    temp_long_1 = temp_long_1 + 1
                                    
                                    temp_string_3 = temp_string_1 ' "#Xp" + getKlartext(temp_string_1, "")
                                
                                    If (m_toggle_mr_stringer_fkt) Then
                                    
                                        temp_string_2 = temp_string_3 & " " & AUSRICHT_STRING_TEMP_1 & "= FkString.getXmlString( xml_root_x, """ & temp_long_1 & """);"
                                        
                                    Else
                                    
                                        temp_string_2 = "xml_string += ""<" & temp_long_1 & ">"" + " & temp_string_3 & " " & AUSRICHT_STRING_TEMP_1 & "+ ""</" & temp_long_1 & ">"";"
                                        
                                    End If
                                
                                End If
                                
                            End If

                        Else
                            
                            temp_string_2 = LEER_STRING
                        
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_2)
                     
                    ElseIf (pFunktion = FKT_NOTES_DEBUG_FELD_WERTE) Then
                    
                        If (Trim(aktuelle_zeile) <> "") Then
                        
                            temp_string_1 = Trim(akt_zeile_mark)
                            
                            If (temp_string_1 <> "") Then
                                
                                temp_string_2 = "drgetString( notes_dokument, """ & temp_string_1 & """ )"
                                
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    Call cls_string_array.setString(zeilen_zaehler, "'##sss( """ & Replace(temp_string_1, """", """""") & " =>"" & " & temp_string_2 & " & ""<"" )")
                                
                                Else
                                
                                    Call cls_string_array.setString(zeilen_zaehler, "temp_str = temp_str & chr(13) & """ & Replace(temp_string_1, """", """""") & " =>"" & " & temp_string_2 & " & ""<"" ")
                                
                                End If
                            
                            End If

                        End If
                    
                    ElseIf (pFunktion = FKT_GENERATOR_IF) Or (pFunktion = FKT_GENERATOR_IF_2) Then
                    
                        If (aktuelle_zeile <> "") And (ab_position > 0) Then
                        
                            If (pSelLength > 0) Then
                            
                                temp_long_1 = InStr(aktuelle_zeile, inhalt_markierung)
                                
                                If (temp_long_1 > 0) Then
                                
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
                            
                            ergebnis_fkt = LEER_STRING
                        
                        Else
                        
                            If (pFunktion = FKT_GENERATOR_IF_2) Then
                                
                                ergebnis_fkt = zeichen_zeilenumbruch & temp_string_3 & " ( " & temp_string_1 & " == " & temp_string_2 & " ) " & zeichen_zeilenumbruch & "{" & zeichen_zeilenumbruch & "}"
                                
                                temp_string_3 = "else if"
                      
                            Else
                                
                                If (m_toggle_mr_stringer_fkt) Then
                                
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "JAVA_IF1  " & temp_string_3 & " ( str_parameter_name.equalsIgnoreCase( """ & temp_string_1 & """ ) ) "
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "JAVA_IF2  " & temp_string_3 & " ( str_parameter_name == " & temp_string_1 & " ) "
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "{"
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "JAVA_TH1  " & "  str_lokale_variable  = """ & temp_string_2 & """; "
                                    
                                    If (Right(temp_string_2, 1) = ";") Then
                                        
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "JAVA_TH2  " & "  str_lokale_variable  = " & temp_string_2 & ""
                                    
                                    Else
                                        
                                        ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "JAVA_TH2  " & "  str_lokale_variable  = " & temp_string_2 & ";"
                                    
                                    End If
                                    
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "}"
                                    
                                    temp_string_3 = "else if"
                                
                                Else
                                   
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "VB_IF1" & temp_string_3 & " ( str_parameter_name  = " & temp_string_1 & ") Then"  '#1
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "VB_IF2" & temp_string_3 & " ( str_parameter_name  = """ & temp_string_1 & """) Then"  '#1
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "VB_TH1  " & "      str_lokale_variable  = " & temp_string_2 & " "
                                    ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & "VB_TH2  " & "      str_lokale_variable  = """ & temp_string_2 & """"
                                    
                                    temp_string_3 = "ElseIf"
        
                                End If
    
                            End If
    
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, ergebnis_fkt)
                        
                        ergebnis_fkt = LEER_STRING
    
                    ElseIf (pFunktion = FKT_JAVA_GENERATOR) Then
    
                        '
                        ' Anforderungen an Generatorquelltext
                        ' Slasches und Anfuehrungszeichen qualifizieren, da diese nachher in einem String stehen.
                        '
                        aktuelle_zeile = Replace(Replace(trimT(aktuelle_zeile), "\", "\\"), """", "\""")
                        '
                        ' Leerzeichen vor und nach Klammern
                        '
                        aktuelle_zeile = Replace(Replace(Replace(Replace(aktuelle_zeile, "(", "( "), ")", " )"), "[", "[ "), "]", " ]")
                        '
                        ' Eliminierung von doppelten Leerzeichen nach Klammern
                        '
                        aktuelle_zeile = Replace(Replace(Replace(Replace(Replace(Replace(aktuelle_zeile, "(  )", "()"), "[  ]", "[]"), "(  ", "( "), "  )", " )"), "[  ", "[ "), "  ]", " ]")
                        
                        aktuelle_zeile = temp_string_1 & aktuelle_zeile & temp_string_2
                        '
                        ' Ersetzungen fuer VB6 rausgenommen
                        'If (m_toggle_mr_stringer_fkt) Then
                        '
                        '    aktuelle_zeile = Replace(aktuelle_zeile, "( ""    ", "( TAB_STR + """)
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

    
                   ElseIf (pFunktion = FKT_UMDREHEN) Then
                        '
                        ' Funktion "Umdrehen"
                        ' Dreht die Zeichen der aktuellen Zeile oder Markierung um.
                        '
                        temp_string_1 = StrReverse(akt_zeile_mark)

                        If (knz_benutze_markierung) Then
                    
                            temp_string_1 = replaceSubstringAbBis(aktuelle_zeile, ab_position, bis_position, temp_string_1)
                        
                        End If
                        
                        Call cls_string_array.setString(zeilen_zaehler, temp_string_1)
                        
                    ElseIf (pFunktion = FKT_GET_DIR) Then
                        '
                        ' Funktion "Verzeichnis einlesen"
                        ' Liest das Verzeichis ein, welches in der aktuellen Zeile steht.
                        ' Ist die aktuelle Zeile kein Verzeichnisname, wird von der benutzten
                        ' VB-Funktion keine Datei zurueckgegeben, also nichts gelesen.
                        '
                        aktuelle_zeile = Trim(aktuelle_zeile)
                        
                        If aktuelle_zeile <> "" Then
    
                            If (InStr("\/", Right(aktuelle_zeile, 1)) <= 0) Then
                            
                                aktuelle_zeile = aktuelle_zeile & "\"
                                
                            End If
                            
                            If (m_toggle_mr_stringer_fkt) Then
                            
                                temp_string_2 = aktuelle_zeile
                                
                            Else
                            
                                temp_string_2 = LEER_STRING
                                
                            End If
                            
                            temp_long_1 = 0
                            temp_string_1 = Dir(aktuelle_zeile & "*.*")
    
                            While (temp_string_1 <> "") And (temp_long_1 < 32123)
                            
                                ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & temp_string_2 & temp_string_1
                                
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
                '
                If (pFunktion = FKT_ERSTELLE_XML) Or (pFunktion = FKT_ERSTELLE_XML_2) Then
                
                    '
                    ' Funktion "XML"
                    ' Dem Ergebnis wird noch der Vorlauf und der Nachlauf (abschliessende Klammer) hinzugefuegt.
                    '
                    ergebnis_fkt = "<?xml version=""1.0"" encoding=""UTF-8""?>" & zeichen_zeilenumbruch & "<XML_KLAMMER_1>" & zeichen_zeilenumbruch & ergebnis_fkt & "</XML_KLAMMER_1>" & zeichen_zeilenumbruch
    
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
                    
                    ergebnis_fkt = "ERGEBNIS >" & temp_long_1 & "<  >" & temp_double_2 & "< Durchschnitt >" & temp_double_1 & "<" & zeichen_zeilenumbruch & ergebnis_fkt
                
                ElseIf (pFunktion = FKT_GET_EINMALIGE_VORKOMMEN) Then
                    '
                    ' Funktion "Einmalige Vorkommen"
                    ' Zeilenumbruch fuer das temporaere "#1#", sowie ein Leerstring fuer "#2#".
                    '
                    ergebnis_fkt = Replace(Replace(temp_string_2, "#1#", zeichen_zeilenumbruch), "#2#", "")
                 
                ElseIf ((pFunktion = FKT_GETTER_SETTER_JAVA) Or (pFunktion = FKT_GETTER_SETTER_VB) Or (pFunktion = FKT_SINGLETON_JAVA)) Then
                    '
                    ' Funktion ""
                            
                    ergebnis_fkt = temp_string_3 & zeichen_zeilenumbruch & ergebnis_fkt & zeichen_zeilenumbruch & ""
                 
                ElseIf (Len(ergebnis_fkt) = 0) Then
                    '
                    ' Bestimmung Ergebnis aus Zeilenarray
                    ' Ist die Variable "ergebnis_fkt" noch nicht mit einem Wert versehen worden,
                    ' bestimmt sich das Funktionsergebnis aus allen Zeilen des Zeilenarrays.
                    ' In diesem Fall veraendern die Funktionen die gespeicherten Zeilen im Objekt.
                    '
                    ' Andere Funktieonen benutzten gleich die Variable "ergebnis_fkt".
                    '
                    ergebnis_fkt = cls_string_array.toString(zeichen_zeilenumbruch)
                    
                End If
                
                If (pFunktion = FKT_STRING_IT) Then
            
                    If (m_zaehler_string_it = 2) Then
                    
                        ergebnis_fkt = "String j_str  = """";" & vbCrLf & vbCrLf & ergebnis_fkt
                    
                    End If
                    
                End If

            
            End If
        
        End If
    
    End If
        
    m_knz_aktiv = False
    
    startMrStringer = ergebnis_fkt
    
EndFunktion:
    
    Set cls_string_array = Nothing
    
    Exit Function
    
errStartMrStringer:
    
    startMrStringer = "Fehler: " & Error
    
    Resume EndFunktion

End Function

'################################################################################
'
Public Function startJoin(pString1 As String, pString2 As String, pTrennzeichen As String, Optional knz_restart_zaehler_2 As Boolean = False) As String

On Error GoTo errStartJoin

Dim cls_string_array_1        As clsStringArray
Dim cls_string_array_2        As clsStringArray
Dim ergebnis_fkt              As String
Dim zeilen_zaehler_1          As Long
Dim zeilen_zaehler_2          As Long
Dim zeilen_anzahl_1           As Long
Dim zeilen_anzahl_2           As Long
Dim zeichen_zeilenumbruch     As String

    If (m_knz_join_anfang = 1) Then
                
        Set cls_string_array_1 = startMultiline(pString1)
        Set cls_string_array_2 = startMultiline(pString2)
            
    Else
                    
        Set cls_string_array_1 = startMultiline(pString2)
        Set cls_string_array_2 = startMultiline(pString1)
        
    End If
    
    If (cls_string_array_1 Is Nothing) And (cls_string_array_2 Is Nothing) Then
    
        startJoin = LEER_STRING
        
    ElseIf (cls_string_array_2 Is Nothing) Then
    
        startJoin = pString1
    
    ElseIf (cls_string_array_1 Is Nothing) Then
    
        startJoin = pString2
        
    Else
        
        m_knz_join_anfang = m_knz_join_anfang + 1
        
        If (m_knz_join_anfang > 1) Then
            
            m_knz_join_anfang = 0
        
        End If

        zeilen_anzahl_1 = cls_string_array_1.getAnzahlStrings
        zeilen_anzahl_2 = cls_string_array_2.getAnzahlStrings
        
        zeichen_zeilenumbruch = LEER_STRING
        
        zeilen_zaehler_1 = 1
        zeilen_zaehler_2 = 1
        
        While (zeilen_zaehler_1 <= zeilen_anzahl_1) Or (zeilen_zaehler_2 <= zeilen_anzahl_2)
            
            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & cls_string_array_1.getString(zeilen_zaehler_1) & pTrennzeichen & cls_string_array_2.getString(zeilen_zaehler_2)
            
            zeichen_zeilenumbruch = Chr(13) & Chr(10)
            
            zeilen_zaehler_1 = zeilen_zaehler_1 + 1
            zeilen_zaehler_2 = zeilen_zaehler_2 + 1
        
            If (knz_restart_zaehler_2) Then
            
                If (zeilen_zaehler_2 > zeilen_anzahl_2) Then
                    
                    zeilen_zaehler_2 = 1
                
                End If
                
                If (zeilen_zaehler_1 > zeilen_anzahl_1) Then
                    
                    zeilen_zaehler_2 = zeilen_anzahl_2 + 1
                
                End If
                
            End If
            
        Wend
          
        startJoin = ergebnis_fkt

    End If
    
EndFunktion:
    
    Set cls_string_array_1 = Nothing
    Set cls_string_array_2 = Nothing
    
    Exit Function
    
errStartJoin:
    Resume EndFunktion
    
End Function

'####################################################################################################
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
Dim ergebnis_fkt                    As String  ' Der Stringbuffer fuer die Aufnahme des Ergebnisses
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
    eingabe_str = Replace(Replace(pString, Chr(13) & Chr(10), ""), Chr(13), "")
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
    ' staendige Teil der ersten Ergebniszeile geht. Wenn nach dem ersten gefundenen
    ' XML-Startzeichen ein "/" kommt, so muss ab dem Startzeichen noch die
    ' Position des Endzeichens ermittelt werden. Dieses ist z.B. hier der Fall
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
    
        ergebnis_fkt = ergebnis_fkt & Mid(eingabe_str, lese_position_start, (lese_position_ende - lese_position_start) + 1) & Chr(13) & Chr(10)
        
        lese_position_start = lese_position_ende + 1
        
        If (pos_akt_xml_start_zeichen < len_eingabe_str) Then
            
            pos_akt_xml_start_zeichen = InStr(pos_akt_xml_start_zeichen + 1, eingabe_str, start_zeichen)
            
            If (pos_akt_xml_start_zeichen = 0) Then
            
                ergebnis_fkt = ergebnis_fkt & Trim(Mid(eingabe_str, lese_position_start, len_eingabe_str)) & Chr(13) & Chr(10)
            
            End If
        
        End If
        
    End If
    '
    ' Die Schleife wird ausgefuehrt solange noch das Startzeichen gefunden wird
    ' und der Endllosschleifenverhinderungszaehler noch kleiner als 32000 ist.
    '
    While (pos_akt_xml_start_zeichen > 0) And (zaehler_while_schleife < 32000)
    
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
            ' Pruefung: Endzeichen gefunden?
            ' Bedingung: Endzeichen ">" liegt hinter Startzeichen "<"
            '
            If (pos_akt_xml_end_zeichen > pos_akt_xml_start_zeichen) Then
                '
                ' Pruefung:   End-Tag Variante 1 - Normal "</end_tag>"
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
                ' Pruefung:   End-Tag Variante 2 "<end_tag/>"
                ' Bedingung: Liegt 1 Zeichen vor dem Ende ein "/"  ?
                '
                ' Das Kennzeichen fuer die Tag-Art wird auf "tag_art_start_ende__kommentar" gestellt.
                '
                If (Mid(eingabe_str, pos_akt_xml_end_zeichen - 1, 1) = "/") Then
                
                    knz_tag_art = tag_art_start_ende__kommentar
                    
                End If
                '
                ' Pruefung:   Kommentar  "<!-- Kommentar -->"
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
        ' Sonderfall ist, dass ein Start-Ende-Tag nach einer oeffnenden Klammer kommt
        ' Das Start-Tag muss noch in die Ergebnismenge uebertragen werden.
        '
        If ((knz_letzte_tag_art = tag_art_start) And (knz_tag_art = tag_art_start_ende__kommentar)) Then
        
            'str_debug = "A - "
            lese_position_ende = letztes_end_tag_ende
        
            ergebnis_fkt = ergebnis_fkt & str_debug & String(einrueck_zaehler, " ") & Trim(Mid(eingabe_str, lese_position_start, (lese_position_ende - lese_position_start) + 1)) & Chr(13) & Chr(10)
            '
            ' Leseposition-Start aktualisieren
            ' Die naechste Startposition fuer den Leseprozess liegt hinter dem aktuellen Endezeichen.
            ' Das ist genau die Startposition fuer das aktuelle XML-Tag
            '
            lese_position_start = lese_position_ende + 1
            '
            ' Einrueckung
            ' Nach einer oeffnenden XML-klammer (=letztes TAG) muss die Einrueckung fuer die
            ' aktuelle XML-Klammer vor dem evtl. naechsten Trennen erhoeht werden.
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
            If (akt_teil_string <> "") Then
                
                ergebnis_fkt = ergebnis_fkt & str_debug & String(einrueck_zaehler, " ") & akt_teil_string & Chr(13) & Chr(10)
            
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
            
            If (akt_teil_string <> "") Then
            
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
            
               ergebnis_fkt = ergebnis_fkt & String(einrueck_zaehler, " ") & Trim(Mid(eingabe_str, lese_position_start, len_eingabe_str)) & Chr(13) & Chr(10)
            
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
    ' Vor der Rueckgabe werden noch 2 hintereinander kommende Zeilenumbrueche in
    ' einen Zeilenumbruch gewandelt.
    '
    formatXML = Replace(ergebnis_fkt, Chr(13) & Chr(10) & Chr(13) & Chr(10), Chr(13) & Chr(10)) & Chr(13) & Chr(10)
    
End Function


'################################################################################
'
' Aus der Eingabe werden nur die in der Funktion hinterlegten gueltigen Zeichen uebernommen.
'
' PARAMETER: pString        = der zu behandelnde Eingabestring
'
' RETURN : Einen String mit nur den Zeichen der gueltigen Zeichen
'
Public Function getStringGueltigeZeichen(pString As String) As String

Dim ergebnis_str     As String
Dim akt_position     As Integer
Dim akt_zeichen      As String
Dim gueltige_zeichen As String

Dim my_cr As String

    gueltige_zeichen = " enirstaudhgolcmfbkVvwz1paeSDA0E2RBGueMIPKF9UNW3L78oeH4T5CZJy6xjOUeYXqQäöüÄÖÜ_?!""Â§$%&/()<>[]{}=*'/*-+:;,.#\/1234567890"

    m_toggle_mr_stringer_fkt = Not m_toggle_mr_stringer_fkt
  
    If (m_toggle_mr_stringer_fkt) Then
    
        my_cr = Chr(13)
    
    Else
    
        my_cr = Chr(13) & Chr(10)
    
    End If
  
    akt_position = 1

    While (akt_position <= Len(pString))

        akt_zeichen = Mid(pString, akt_position, 1)

        If (InStr(gueltige_zeichen, akt_zeichen) > 0) Then

            ergebnis_str = ergebnis_str & akt_zeichen
            
        ElseIf (akt_zeichen = vbCr) Then
        
            ergebnis_str = ergebnis_str & my_cr
        
        End If

        akt_position = akt_position + 1

    Wend

    getStringGueltigeZeichen = ergebnis_str

End Function

'################################################################################
'
Private Function renameDatei(pString As String) As String

Dim zaehler_unterstrich   As Integer
Dim eingabe_string        As String
Dim ausgabe_string        As String
Dim akt_zeichen           As String
Dim erweiterung_str       As String
Dim str_unterstrich       As String ' Vermeidung von fuehrenden Unterstrichen
Dim akt_position          As Integer
Dim pos_letzter_punkt     As Integer
Dim akt_chr_wert As Integer
    
    str_unterstrich = "!"
    eingabe_string = Trim(pString)
    akt_position = 1
    pos_letzter_punkt = -1
    zaehler_unterstrich = 0
    '
    ' Die While-Schleife laeuft ueber die Laenge des Eingabestrings.
    '
    While (akt_position <= Len(eingabe_string))
    
        akt_zeichen = Mid(eingabe_string, akt_position, 1)
        '
        '######################################################################
        '
        ' Auswertung des aktuellen Zeichens
        '
        ' Zeichen ohne Spezialbehandlung sind in der Konstanten "GUELTIGE_ZEICHEN_DATEI_NAME"
        ' hinterlegt.
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
            ElseIf (akt_zeichen = "ä") Then
                akt_zeichen = "ae"
            ElseIf (akt_zeichen = "Ä") Then
                akt_zeichen = "Ae"
            ElseIf (akt_zeichen = "ö") Then
                akt_zeichen = "oe"
            ElseIf (akt_zeichen = "Ö") Then
                akt_zeichen = "Oe"
            ElseIf (akt_zeichen = "ü") Then
                akt_zeichen = "ue"
            ElseIf (akt_zeichen = "Ü") Then
                akt_zeichen = "Ue"
            ElseIf (akt_zeichen = "ß") Then
                akt_zeichen = "ss"
            ElseIf (akt_zeichen = "´" Or akt_zeichen = "`") Then
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
        '######################################################################
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
        '
        If (akt_zeichen <> "_") Then
        
            ausgabe_string = ausgabe_string & akt_zeichen
            
            str_unterstrich = "_"
            zaehler_unterstrich = 0
        '
        ' Pruefung: aktuelles Zeichen ist Unterstrich?
        ' Um fuehrende Unterstriche zu vermeiden wird durch die Variable "str_unterstrich"
        ' in den ersten Laeufen ein anderes Zeichen hinterlegt. Somit kann die
        ' Abfrage nicht greifen. Wird ein gueltiges Startzeichen gefunden,
        ' wird gleichzeitig in der Variablen "str_unterstrich" ein Unterstrich
        ' hinterlegt und schaltet somit diese Abfrage erst frei.
        '
        ' Pfruefung: Zaehler fuer Unterstrich kleiner 2 ?
        ' Es duerfen maximal 2 Unterstriche hintereinander stehen. Dieses wird
        ' durch eine Zaehlvariable fuer hinzugefuegte Unterstriche gemacht. Fuer
        ' jeden hinzugefuegten Unterstrich wird der Zaehler erhoeht.
        '
        ElseIf (akt_zeichen = str_unterstrich) Then
        
            If (zaehler_unterstrich < 2) Then
                
                ausgabe_string = ausgabe_string & akt_zeichen
                
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
    If (pos_letzter_punkt > 4) And (pos_letzter_punkt > Len(eingabe_string) - 6) Then
        '
        ' Die Dateierweiterung wird aus der Eingabezeichenfolge gelesen und in
        ' Kleinbuchstaben konvertiert. Der Punkt wird nicht mit aufgenommen.
        '
        erweiterung_str = LCase(Mid(eingabe_string, pos_letzter_punkt + 1))
        '
        ' Der Ausgabestring wird bis zur Laenge der Erweiterung (bis zur Punktposition)
        ' gelesen.
        '
        ausgabe_string = Left(ausgabe_string, (Len(ausgabe_string) - Len(erweiterung_str)) - 1) & "."
        '
        ' Falsche Erweiterungen werden nach AVI konvertiert, aus MPEG wird MPG.
        ' Die Dateierweiterung soll aus 3 Zeichen und aus Kleinbuchstaben bestehen.
        '
        If (erweiterung_str = "xvid") Then
            
            erweiterung_str = "avi"
        
        ElseIf (erweiterung_str = "divx") Then
            
            erweiterung_str = "avi"
        
        ElseIf (erweiterung_str = "mpeg") Then
            
            erweiterung_str = "mpg"
        
        End If
        '
        ' Die Erweiterung wird dem Ausgabestring hinzugefuegt.
        '
        ausgabe_string = ausgabe_string & erweiterung_str
        '
        ' Unterstriche vor der Erweiterung werden geloescht.
        '
        erweiterung_str = "." & erweiterung_str
        
        ausgabe_string = Replace(ausgabe_string, "__" & erweiterung_str, erweiterung_str)
        ausgabe_string = Replace(ausgabe_string, "_" & erweiterung_str, erweiterung_str)
        ausgabe_string = Replace(ausgabe_string, "_" & erweiterung_str, erweiterung_str)
    
    End If
   
    ausgabe_string = Replace(ausgabe_string, "_is_", "_Is_")
    ausgabe_string = Replace(ausgabe_string, "_it_", "_It_")
    ausgabe_string = Replace(ausgabe_string, "_to_", "_To_")
    ausgabe_string = Replace(ausgabe_string, "_are_", "_Are_")
    ausgabe_string = Replace(ausgabe_string, "_of_", "_Of_")
    ausgabe_string = Replace(ausgabe_string, "_from_", "_From_")
    ausgabe_string = Replace(ausgabe_string, "_the_", "_The_")
    ausgabe_string = Replace(ausgabe_string, "_REPACK", "")
    ausgabe_string = Replace(ausgabe_string, "_XviD", "")
    ausgabe_string = Replace(ausgabe_string, "_DivX", "")
    ausgabe_string = Replace(ausgabe_string, "_GERMAN", "")
    ausgabe_string = Replace(ausgabe_string, "_HDTVRiP", "")
    ausgabe_string = Replace(ausgabe_string, "_HDRiP", "")
    ausgabe_string = Replace(ausgabe_string, "_DVDRiP", "")
    ausgabe_string = Replace(ausgabe_string, "_VHSRiP", "")
    ausgabe_string = Replace(ausgabe_string, "_SatRip", "")
    ausgabe_string = Replace(ausgabe_string, "_x264", "")
    ausgabe_string = Replace(ausgabe_string, "_H264", "")
    ausgabe_string = Replace(ausgabe_string, "_HDTV", "")
    ausgabe_string = Replace(ausgabe_string, "_H264", "")
    ausgabe_string = Replace(ausgabe_string, "_BluRay", "")
    ausgabe_string = Replace(ausgabe_string, "_DVDR", "")
    ausgabe_string = Replace(ausgabe_string, "_BDRiP", "")
    ausgabe_string = Replace(ausgabe_string, "_TVP", "")
    ausgabe_string = Replace(ausgabe_string, "_DVD9", "")
    ausgabe_string = Replace(ausgabe_string, "_DVD5", "")
    ausgabe_string = Replace(ausgabe_string, "_DVD1", "")
    ausgabe_string = Replace(ausgabe_string, "_DVD2", "")
    ausgabe_string = Replace(ausgabe_string, "_DVD4", "")
    ausgabe_string = Replace(ausgabe_string, "_DVD5", "")
    ausgabe_string = Replace(ausgabe_string, "_AC3", "")
    ausgabe_string = Replace(ausgabe_string, "_1080p", "")
    ausgabe_string = Replace(ausgabe_string, "_1080i", "")
    ausgabe_string = Replace(ausgabe_string, "_720p", "")
    ausgabe_string = Replace(ausgabe_string, "_720i", "")

    ausgabe_string = Replace(ausgabe_string, "_Avi", ".avi")
    
    renameDatei = ausgabe_string

End Function

'################################################################################
'
Public Function startGetHexDump(pString As String, pZahlenJeZeile As Integer) As String

On Error GoTo errStartGetHexDump

    m_toggle_mr_stringer_fkt = Not m_toggle_mr_stringer_fkt

Dim akt_position            As Integer
Dim akt_zeichen             As String
Dim my_cr                   As String
Dim str_ergebnis            As String
Dim str_pos                 As String
Dim str_vorlaufende_zahlen  As String
Dim str_zahlen              As String
Dim str_zeichen             As String
Dim zahlen_je_zeile_anzahl  As Integer
Dim zahlen_je_zeile_zaehler As Integer
Dim knz_hex_zahl            As Boolean
    
    my_cr = vbCrLf
    
    str_vorlaufende_zahlen = "0000000"
    
    str_ergebnis = LEER_STRING
    
    zahlen_je_zeile_anzahl = pZahlenJeZeile
    
    zahlen_je_zeile_zaehler = 0
    
    akt_position = 1
    
    str_pos = Right(str_vorlaufende_zahlen & akt_position, 6) & " "
    
    knz_hex_zahl = m_toggle_mr_stringer_fkt
    
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
        If (knz_hex_zahl) Then
        
            str_zahlen = str_zahlen & " " & Right(str_vorlaufende_zahlen & Hex(Asc(akt_zeichen)), 2) & " "
            
        Else
        
            str_zahlen = str_zahlen & Right(str_vorlaufende_zahlen & Asc(akt_zeichen), 3) & " "
            
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
        
            str_zeichen = str_zeichen & "."
            
        Else
        
            str_zeichen = str_zeichen & akt_zeichen
            
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
            
            str_ergebnis = str_ergebnis & my_cr & str_pos & str_zahlen & str_zeichen
            
            str_pos = Right(str_vorlaufende_zahlen & (akt_position + 1), 6) & " "
            
            str_zahlen = LEER_STRING
            
            str_zeichen = LEER_STRING
            
            zahlen_je_zeile_zaehler = 0
        
        End If
        
        akt_position = akt_position + 1
    
    Wend

    If (zahlen_je_zeile_zaehler > 0) Then
        
        While (zahlen_je_zeile_zaehler < zahlen_je_zeile_anzahl)
            
            str_zahlen = str_zahlen & "   " & " "
            
            zahlen_je_zeile_zaehler = zahlen_je_zeile_zaehler + 1
        
        Wend
        
        str_ergebnis = str_ergebnis & my_cr & str_pos & str_zahlen & str_zeichen
        
    End If

EndFunktion:

    On Error Resume Next

    str_zahlen = LEER_STRING
    
    str_zeichen = LEER_STRING

    DoEvents

    startGetHexDump = str_ergebnis

    Exit Function

errStartGetHexDump:

    str_ergebnis = str_ergebnis & "Fehler: errStartGetHexDump: " & Err & " " & Error & " " & Erl

    Resume EndFunktion

End Function


Public Function startGetAsciiPrint(pString As String) As String

'
' ? startGetAsciiPrint( "!#$%&'*+-/=?^_`{|}~" )
'
On Error GoTo errStartGetAsciiPrint

Dim akt_position            As Integer
Dim akt_zeichen             As String
Dim my_cr                   As String
Dim str_ergebnis            As String
Dim str_vorlaufende_zahlen  As String
    
    my_cr = vbCrLf
    
    str_vorlaufende_zahlen = "0000000"
    
    str_ergebnis = LEER_STRING
    
    akt_position = 1
    
    While ((akt_position <= Len(pString)) And (akt_position <= 900))
    
        akt_zeichen = Mid(pString, akt_position, 1)
            
        If (Asc(akt_zeichen) < 31) Then
        
            str_ergebnis = str_ergebnis & my_cr & TRENN_STRING_6 & " asc(   ) = " & Right(str_vorlaufende_zahlen & Asc(akt_zeichen), 3) & " "
            
        Else
        
            str_ergebnis = str_ergebnis & my_cr & TRENN_STRING_6 & " asc(""" & akt_zeichen & """) = " & Right(str_vorlaufende_zahlen & Asc(akt_zeichen), 3) & " "
            
        End If
        
        
        akt_position = akt_position + 1
    
    Wend

EndFunktion:

    On Error Resume Next

    DoEvents

    startGetAsciiPrint = str_ergebnis

    Exit Function

errStartGetAsciiPrint:

    str_ergebnis = str_ergebnis & "Fehler: errStartGetAsciiPrint: " & Err & " " & Error & " " & Erl

    Resume EndFunktion

End Function

'################################################################################
'
Public Function startGrepSuchWorte(pSuchWorte As String, pString As String, pKnzArt As Integer) As String

On Error GoTo errStartGrepSuchWorte

Dim cls_string_array   As clsStringArray
Dim str_markierung     As String
Dim akt_zaehler        As Long
Dim akt_anzahl_zeilen  As Long
Dim aktuelle_zeile     As String

    startGrepSuchWorte = LEER_STRING
    
    If (Trim(pSuchWorte) <> "") Then
        
        str_markierung = "#M#A#R#K#"
        
        Set cls_string_array = startMultiline(pSuchWorte)
        
        If (cls_string_array Is Nothing) = False Then
            
            akt_anzahl_zeilen = cls_string_array.getAnzahlStrings
            akt_zaehler = 0
            
            While (akt_zaehler <= akt_anzahl_zeilen)
            
                aktuelle_zeile = cls_string_array.getString(akt_zaehler)
            
                If (aktuelle_zeile <> "") Then
                
                    pString = Replace(pString, aktuelle_zeile, str_markierung & aktuelle_zeile)
                    
                End If

                akt_zaehler = akt_zaehler + 1
                
            Wend
            
            Set cls_string_array = Nothing
            '
            '#######################################################################
            '
            Set cls_string_array = startMultiline(pString)
            
            If (cls_string_array Is Nothing) = False Then
                
                akt_anzahl_zeilen = cls_string_array.getAnzahlStrings
                akt_zaehler = 0
                
                While (akt_zaehler <= akt_anzahl_zeilen)
                
                    aktuelle_zeile = cls_string_array.getString(akt_zaehler)
                
                    If (aktuelle_zeile <> "") Then
                        
                        If (pKnzArt = 1) Then ' 1 = Positiv Grep +
                        
                            If (InStr(aktuelle_zeile, str_markierung) = 0) Then
                                '
                                ' Alle markierten Zeilen bleiben drin. Alle Zeilen ohne Markierung
                                ' werden geloescht.
                                '
                                Call cls_string_array.setString(akt_zaehler, "")
                            
                            End If
                            
                        Else ' 0 oder andere Zahl = Negativ Grep -
                            
                            If (InStr(aktuelle_zeile, str_markierung) > 0) Then
                                '
                                ' Alle Zeilen, in denen ein Suchwort gefunden wurde, werden ausgenullt
                                ' Es bleiben am Ende nur diejenigen Zeilen uebrig, bei welchem kein
                                ' Suchwort vorkam.
                                '
                                Call cls_string_array.setString(akt_zaehler, "")
                            
                            End If
                            
                        End If
                        
                    End If
    
                    akt_zaehler = akt_zaehler + 1
                    
                Wend
                
                startGrepSuchWorte = Replace(cls_string_array.toString(Chr(13) & Chr(10), True), str_markierung, "")
                
                Set cls_string_array = Nothing
            
            End If
        End If
    End If
    
EndFunktion:
    
    Set cls_string_array = Nothing
    
    Exit Function
    
errStartGrepSuchWorte:
    Resume EndFunktion
    
End Function

'################################################################################
'
Public Function startCsvKonstanten(pString As String, pTrennzeichen As String) As String

On Error GoTo errStartCsvKonstanten

Dim cls_string_array     As clsStringArray
Dim cls_erg              As clsStringArray
Dim zeilen_zaehler       As Long
Dim anzahl_zeilen        As Long
Dim suche_was            As String
Dim akt_zeile            As String
Dim trenn_pos            As Integer
Dim temp_string_1a       As String
Dim str_2                As String
Dim pos_start            As Integer
Dim akt_konstanten_name  As String
Dim akt_konstanten_wert  As String
Dim akt_konstanten_dekl  As String
    
    startCsvKonstanten = LEER_STRING
    
    Set cls_string_array = startMultiline(pString)
    
    If (cls_string_array Is Nothing) = False Then
        
        Set cls_erg = New clsStringArray
        
        If (cls_string_array Is Nothing) = False Then
            
            m_csv_feld_nummer = m_csv_feld_nummer + 1
    
            anzahl_zeilen = cls_string_array.getAnzahlStrings
            
            zeilen_zaehler = 1
            
            While (zeilen_zaehler <= anzahl_zeilen)
                
                akt_zeile = Trim(cls_string_array.getString(zeilen_zaehler))
                
                trenn_pos = InStr(akt_zeile, pTrennzeichen)
                
                If (trenn_pos > 0) Then
                    
                    pos_start = trenn_pos + Len(pTrennzeichen)
                    
                    temp_string_1a = Left(akt_zeile, trenn_pos - 1)
                    
                    str_2 = Mid(akt_zeile, pos_start, Len(akt_zeile))
                    
                    akt_konstanten_name = "CONST_" & UCase(getKlartext(temp_string_1a, "_"))
                    
                    akt_konstanten_wert = str_2
                    
                    akt_konstanten_dekl = "Const " & akt_konstanten_name & " = """ & akt_konstanten_wert & """ "
                    
                    Call cls_erg.addString(akt_konstanten_dekl)
                
                End If
                
                zeilen_zaehler = zeilen_zaehler + 1
            
            Wend
            
            Call cls_erg.addString("")
            Call cls_erg.addString("")
            
            zeilen_zaehler = 1
            
            While (zeilen_zaehler <= anzahl_zeilen)
                
                akt_zeile = Trim(cls_string_array.getString(zeilen_zaehler))
                
                trenn_pos = InStr(akt_zeile, pTrennzeichen)
                
                If (trenn_pos > 0) Then
                    
                    pos_start = trenn_pos + Len(pTrennzeichen)
                    
                    temp_string_1a = Left(akt_zeile, trenn_pos - 1)
                    
                    str_2 = Mid(akt_zeile, pos_start, Len(akt_zeile))
                    
                    akt_konstanten_name = "STR_" & UCase(getKlartext(temp_string_1a, "_"))
                    
                    akt_konstanten_wert = str_2
                    
                    akt_konstanten_dekl = "public static final String " & akt_konstanten_name & " = """ & akt_konstanten_wert & """; "
                    
                    Call cls_erg.addString("/** Konstante fuer " & getKlartext(temp_string_1a, " ") & " """ & akt_konstanten_wert & """ */")
                    
                    Call cls_erg.addString(akt_konstanten_dekl)
                
                End If
                
                zeilen_zaehler = zeilen_zaehler + 1
            
            Wend
            
            Call cls_erg.addString("")
            Call cls_erg.addString("")
            
            startCsvKonstanten = cls_erg.toString(Chr(13) & Chr(10))

        End If
    End If
    
EndFunktion:
    
    Set cls_string_array = Nothing
    
    Exit Function
    
errStartCsvKonstanten:
    Resume EndFunktion
    
End Function

'################################################################################
'
' ? getStringAb( "ABC.DEF.GHI.JKL.MNO", -6, 8 )
'
Public Function getStringAbBis(pEingabe As String, pAbPosition As Long, pBisPosition As Long) As String

    If (pAbPosition > 0) And (Len(pEingabe) >= pAbPosition) And (pAbPosition <= pBisPosition) Then

        getStringAbBis = Mid(pEingabe, pAbPosition, (pBisPosition - pAbPosition) + 1)
    
    Else
    
        getStringAbBis = LEER_STRING

    End If

End Function

'################################################################################
'
' ? getRemoveAbBis( "1234567890", 4, 8    ) = 12390
' ? getRemoveAbBis( "1234567890", 4, 18   ) = 123
' ? getRemoveAbBis( "1234567890", 4, -8   ) =
' ? getRemoveAbBis( "1234567890", 0, 8    ) = 90
'
Public Function getRemoveAbBis(pEingabe As String, pAbPosition As Long, pBisPosition As Long) As String

Dim ergebnis_str As String

    ergebnis_str = LEER_STRING
    '
    If (pAbPosition > 1) And (pAbPosition <= pBisPosition) Then
        
        ergebnis_str = Left(pEingabe, pAbPosition - 1)
            
    End If

    If (pBisPosition > 0) And (pAbPosition <= pBisPosition) And (pBisPosition < Len(pEingabe)) Then
        
         ergebnis_str = ergebnis_str & Right(pEingabe, Len(pEingabe) - pBisPosition)
         
    End If
    
    getRemoveAbBis = ergebnis_str

End Function

'################################################################################
'
' ? replaceSubstringAbBis( "1234567890", 4,  8 , "ABC" ) 123ABC90
' ? replaceSubstringAbBis( "1234567890", 4, 18 , "ABC" ) 123ABC
'
Public Function replaceSubstringAbBis(pEingabe As String, pAbPosition As Long, pBisPosition As Long, pReplaceWith As String) As String

Dim ergebnis_str As String

    ergebnis_str = LEER_STRING
    
    If (pAbPosition > 1) And (pAbPosition <= pBisPosition) Then
        
        ergebnis_str = Left(pEingabe, pAbPosition - 1)
            
    End If
    
    ergebnis_str = ergebnis_str & pReplaceWith

    If (pBisPosition > 0) And (pAbPosition <= pBisPosition) And (pBisPosition < Len(pEingabe)) Then
        
         ergebnis_str = ergebnis_str & Right(pEingabe, Len(pEingabe) - pBisPosition)
         
    End If
     
    replaceSubstringAbBis = ergebnis_str

End Function

'################################################################################
'
Public Function getBenutztesChr13(pString As String) As String
    
    If InStr(1, pString, Chr(13) & Chr(10), vbBinaryCompare) > 0 Then
    
        getBenutztesChr13 = Chr(13) & Chr(10)
    
    ElseIf InStr(1, pString, Chr(13), vbBinaryCompare) > 0 Then
        
        getBenutztesChr13 = Chr(13)
        
    Else
    
        getBenutztesChr13 = Chr(13) & Chr(10)
    
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
' PARAMETER: pEingabe       = der zu untersuchende String
' PARAMETER: pSuchString    = der zu suchende Trennstring
'
' RETURN  : die Position des letzten Vorkommens vom Suchstring, bzw. -1 wenn dieser nicht gefunden wurde
'
Public Function getLetztePositionVorPos(ByVal pEingabe As String, ByVal pSuchString As String, ByVal pEndPosition As Long) As Long

    Dim akt_position As Long ' Speichert die aktuell gefundene Position

    getLetztePositionVorPos = -1
    
    If (pSuchString <> "") And (pEndPosition > 0) Then

        akt_position = InStr(pEingabe, pSuchString)

        While (akt_position > 0)
        
            If (pEndPosition > 0) Then ' gibt es eine StoppPosition
            
                If (akt_position <= pEndPosition) Then
                    getLetztePositionVorPos = akt_position
                Else
                    Exit Function
                End If
            
            Else
                 getLetztePositionVorPos = akt_position
            End If

            akt_position = InStr(akt_position + Len(pSuchString), pEingabe, pSuchString)

        Wend
        
    End If

End Function

'################################################################################
'
Public Function startErstelleKonstantenEinfach(pString As String, pTrennzeichen1 As String, pTrennzeichen2 As String, pTrennzeichen3 As String, pFunktion As Integer, pSelStart As Long, pSelLength As Long) As String

On Error GoTo errStartKonstantenEinfach

Dim ab_position                     As Long
Dim akt_string                      As String
Dim aktuelle_zeile                  As String
Dim bis_position                    As Long
Dim cls_string_array                As clsStringArray
Dim ersatz_gatter_1                 As String
Dim ersatz_gatter_2                 As String
Dim ersatz_gatter_3                 As String
Dim ersatz_gatter_4                 As String
Dim knz_benutze_markierung          As Boolean
Dim pos_letztes_chr_vor_sel_start   As Long
Dim pos_trennzeichen                As String
Dim temp_trennzeichen               As String
Dim zeilen_anzahl                   As Long
Dim zeilen_zaehler                  As Long

Dim pTrennzeichen4 As String

    temp_trennzeichen = "##1##2KO"
    
    pTrennzeichen1 = "##K_1"
    pTrennzeichen2 = "##K_2"
    pTrennzeichen3 = "##K_3"
    pTrennzeichen4 = "##K_4"
    
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
        
        Dim inhalt_spalte_1 As String
        Dim inhalt_spalte_2 As String
        
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
            If (aktuelle_zeile <> "") Then
            
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
                inhalt_spalte_1 = UCase(getKlartext(inhalt_spalte_1, "_"))
                
                '
                ' Erstellung Konstanten-Wert
                ' Der Konstantenwert darf selber keine Anfuehrungszeichen enthalten.
                ' Es werden alle Anfuehrungszeichen entfernt und das Ergebnis getrimmt.
                '
                inhalt_spalte_2 = Trim(Replace(inhalt_spalte_2, """", ""))
                
                Call cls_string_array.setString(zeilen_zaehler, pTrennzeichen1 & inhalt_spalte_1 & pTrennzeichen2 & inhalt_spalte_2 & pTrennzeichen3 & inhalt_spalte_1 & pTrennzeichen4)
            
            End If
            
            zeilen_zaehler = zeilen_zaehler + 1
        
        Wend

        If (pFunktion = 1) Then
        
            ersatz_gatter_1 = "public static final String "
            ersatz_gatter_2 = " " & AUSRICHT_STRING_TEMP_1 & "= """
            ersatz_gatter_3 = """; // "" + "
            ersatz_gatter_4 = " + """

        ElseIf (pFunktion = 2) Then
        
            ersatz_gatter_1 = "public const "
            ersatz_gatter_2 = " " & AUSRICHT_STRING_TEMP_1 & "= """
            ersatz_gatter_3 = """ ' "" & "
            ersatz_gatter_4 = " & """
        
        ElseIf (pFunktion = 3) Then
        
            ersatz_gatter_1 = "prop_inst.setProperty( """
            ersatz_gatter_2 = """, " & AUSRICHT_STRING_TEMP_1 & " """
            ersatz_gatter_3 = """ );"
        
        End If
        
        akt_string = cls_string_array.toString(Chr(13) & Chr(10))
        
        ' ##K_1NOTES_ITEM_FELD##K_2"notes_item_feld"##K_3NOTES_ITEM_FELD##K_4
        '
        ' Eliminierung von bestehenden Anfuehrungszeichen
        '
        'akt_string = Replace(aktuelle_zeile, """", "")

        akt_string = Replace(akt_string, pTrennzeichen1, ersatz_gatter_1)
        
        akt_string = Replace(akt_string, pTrennzeichen2, ersatz_gatter_2)
        
        akt_string = Replace(akt_string, pTrennzeichen3, ersatz_gatter_3)
        
        akt_string = Replace(akt_string, pTrennzeichen4, ersatz_gatter_4)
          
        startErstelleKonstantenEinfach = akt_string

    End If
    
EndFunktion:
    
    Set cls_string_array = Nothing
    
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

Dim cls_string_array           As clsStringArray
Dim knz_umbruch_vorhanden      As Boolean
Dim zeichen_zeilenumbruch      As String
Dim zeilen_zaehler             As Long
Dim akt_zeile                  As String
Dim akt_position               As Long
Dim letzte_position            As Long
Dim ergebnis                   As String

    Set cls_string_array = New clsStringArray

    ergebnis = LEER_STRING
    '
    ' Ermittlung welches Zeilenumbruchzeichen verwendet in der Eingabe verwendet wird
    '
    zeichen_zeilenumbruch = Chr(13) & Chr(10)

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

Dim ergebnis_str_buf As String
Dim letztes_zeichen  As String
Dim akt_zeichen      As String
Dim akt_position     As Long

    trimX = ""
    
    If (pString <> "") Then
        
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
        ' ueber eine For-Schleife wird jedes Zeichen der Eingabe geprueft.
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
                
                    ergebnis_str_buf = ergebnis_str_buf & akt_zeichen
                    
                    letztes_zeichen = akt_zeichen
                 
                End If
            
            Else
            
                '
                ' Zeichen ungleich Leerzeichen
                ' Alle anderen Zeichen werden dem Ergebnisstring hinzugefuegt. Das aktuelle
                ' Zeichen wird in der Variablen "letztes_zeichen" gespeichert.
                '
                ergebnis_str_buf = ergebnis_str_buf & akt_zeichen
                
                letztes_zeichen = akt_zeichen
            
            End If
        
        Next
        
    '
    ' Abschlusspruefung
    ' Pruefung, ob das Ergebnis auf ein Leerzeichen endet.
    ' Das Ergebnis endet auf ein Leerzeichen, wenn das letzte hinzugefuegte Zeichen ein Leerzeichen war.
    '
    If (letztes_zeichen = " ") Then
    
        If (Len(ergebnis_str_buf) < 2) Then
      
            ergebnis_str_buf = LEER_STRING
        
        Else
        
            ergebnis_str_buf = Left(ergebnis_str_buf, Len(ergebnis_str_buf) - 1)
        
        End If
      
    End If
    
    trimX = ergebnis_str_buf
        
    End If
    
End Function
  
'################################################################################
'
Public Function trimT(pString As String) As String

Dim akt_position  As Long

    trimT = LEER_STRING
    
    If (pString <> "") Then
        
        For akt_position = Len(pString) To 1 Step -1
        
            If (Mid(pString, akt_position, 1) <> " ") Then
            
                Exit For
            
            End If
            
        Next
        
        If (akt_position > 0) Then
            
            trimT = Left(pString, akt_position)
        
        End If
        
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
Public Function replaceIgnoreCase(pQuellString As String, pSuchString As String, pStringNeu As String) As String

Dim such_string_ucase      As String  ' Suchtext in Grossbuchstaben
Dim quell_string_ucase     As String  ' durchsuchter Text in Grossbuchstaben
Dim ergebnis_string        As String  ' text fuer die Rueckgabe
Dim position_such_string   As Long ' die aktuell gefundene Startposition des Suchstrings
Dim position_such_prozess  As Long ' die aktuelle AB-Position fuer die Suche in quell_string_ucase
Dim zaehler                As Long ' ein Zaehler zur Vermeidung von Endlossschleifen
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
    ' ... der Zaehler noch kleiner 500 ist ( Vermeidung Endlossschleife )
    '
    While (position_such_string > 0) And (zaehler < 500)
        '
        ' Pruefung: Suchstring gefunden ?
        ' Das ist immer dann der Fall, wenn die Positon einen Wert groesser 0 hat.
        '
        If (position_such_string > 0) Then
            '
            ' Ergebnisstring aufbauen
            ' Aus dem Parameter-Quellstring wird von der letzten Position bis zur aktuellen
            ' Position des Suchstrings die Zeichen kopiert. Anschliessend wird die neue
            ' Zeichenfolge aus "pStringNeu" dem Ergebnis hinzugefuegt.
            '
            ' Ist "pStringNeu" ein Leerstring, wird eben der Suchstring aus dem
            ' Quellstring entfernt ( es gibt keinen Ersatzstring ).
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
    'pString ==> wenn null oder keine Zeichen vorhanden sind, wird null zurueckgegeben
    '
    If (Len(Trim(pString)) = 0) Then

        getKlartext = LEER_STRING

        Exit Function

    End If

    Dim ergebnis_str                           As String
    Dim akt_zeichen                            As String
    Dim letztes_zeichen                        As String
    Dim trenn_zeichen                          As String
    Dim knz_letztes_zeichen_war_grossbuchstabe As Boolean
    Dim knz_forciere_kleinbuchstabe            As Boolean
    Dim knz_trennzeichen_einfuegen             As Boolean
    Dim knz_trennzeichen_erlaubt               As Boolean
    Dim knz_next_zeichen_gross                 As Boolean
    Dim knz_hinzfuegen                         As Boolean
    Dim zaehler_schleife                       As Integer

    akt_zeichen = " "
    letztes_zeichen = " "
    trenn_zeichen = pTrennzeichen
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
            ' Die Forcierung von Kleinbuchstaben wird aufgehoben ( = Flag auf FALSE gesetzt ).
            '
            If ((akt_zeichen = "_") Or (akt_zeichen = " ") Or (akt_zeichen = "-") Or (akt_zeichen = "( ") Or (akt_zeichen = " )")) Then

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
                If ((knz_letztes_zeichen_war_grossbuchstabe) And ((akt_zeichen = "_") Or (akt_zeichen = " ") Or (akt_zeichen = "-"))) Then
                
                     ergebnis_str = ergebnis_str & trenn_zeichen
                     
                     knz_letztes_zeichen_war_grossbuchstabe = False
                     
                     knz_trennzeichen_erlaubt = False
                
                End If

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

                ergebnis_str = ergebnis_str & trenn_zeichen

            End If

            '
            ' Aufbau Ergebnis
            ' Das Zeichen aus der Variablen "akt_zeichen" wird dem Ergebnis-String hinzugefuegt.
            '
            ergebnis_str = ergebnis_str & akt_zeichen

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

    getKlartext = ergebnis_str

End Function

Public Function getStringLit(pString As String) As String

Dim ergebnis_str_buf As String
Dim akt_zeichen      As String
Dim akt_position     As Long
Dim knz_in_string    As Boolean

    getStringLit = LEER_STRING
    
    If (pString <> "") Then
        
        knz_in_string = False

        akt_position = 1
        
        '
        ' Schleife Zeichenpruefung
        ' ueber eine For-Schleife wird jedes Zeichen der Eingabe geprueft.
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
                        
                        ' akt_position = Anfuehrungszeichen == geprueft
                        ' akt_position +1 = Anfuehrungszeichen == geprueft
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
                
                    ergebnis_str_buf = ergebnis_str_buf & "|"
                
                Else
                    
                    'ergebnis_str_buf = ergebnis_str_buf & ">"
                
                End If
                
                knz_in_string = Not knz_in_string
                
                akt_zeichen = LEER_STRING
            
            End If
            
            If (knz_in_string) And (akt_zeichen <> "") Then
                
                ergebnis_str_buf = ergebnis_str_buf & akt_zeichen
                
            End If
            
            akt_position = akt_position + 1
            
        Wend
    
        getStringLit = ergebnis_str_buf
        
    End If
    
End Function

Private Function getStringLitKonst(pString As String) As String

Dim ergebnis_str_buf     As String
Dim akt_zeichen          As String
Dim akt_position         As Long
Dim knz_in_string        As Boolean
Dim akt_literal          As String
Dim trennzeichen         As String
Dim trennzeichen_to_set  As String
Dim zeichen_anf_maskierung As String
Dim letztes_zeichen As String
    zeichen_anf_maskierung = """"

    getStringLitKonst = ""
    '
    ' Pruefung: ist pString gesetzt?
    '
    If (pString <> "") Then
        
        knz_in_string = False
        
        letztes_zeichen = "!"
        
        trennzeichen = LEER_STRING
        
        trennzeichen_to_set = getBenutztesChr13(pString)
        
        akt_position = 1
        
        '
        ' Schleife Zeichenpruefung
        ' ueber eine For-Schleife wird jedes Zeichen der Eingabe geprueft.
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
                    
                    ergebnis_str_buf = ergebnis_str_buf & trennzeichen & akt_literal
                    
                    trennzeichen = trennzeichen_to_set
                    
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
            If (knz_in_string) And (akt_zeichen <> "") Then
                
                akt_literal = akt_literal & akt_zeichen
                
            End If
            
            '
            ' Am Ende der While-Schleife wird der Leseprozess um eine Position weitergeschaltet
            akt_position = akt_position + 1
        Wend
    
        getStringLitKonst = ergebnis_str_buf
        
    End If
    
End Function

'################################################################################
'
' Ermittelt wie of die Zeichenfolge pSuchString in der Zeichenfolge pEingabeString vorkommt.
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
Public Function getAnzahlVorkommen(ByVal pEingabeString As String, ByVal pSuchString As String) As Long

Dim zaehler_vorkommen     As Long
Dim aktuelle_position     As Long
Dim laenge_such_string    As Integer

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
    ' Wenn die geschachtelten Vorkommen gezaehlt werden sollen, muss
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
Public Function getPosWortende(ByVal pString As String, ByVal pPositionStart As Integer) As Integer

On Local Error Resume Next

Dim akt_position        As Integer ' aktuelle Leseposition
Dim anzahl_zeichen      As Integer ' Laenge des Eingabestrings
Dim zeichen_wortbestandteil    As String
Dim knz_weiterer_schleifendurchlauf As Boolean

    zeichen_wortbestandteil = "enirstaudhgolcmfbkVvwz1paeSDA0E2RBGueMIPKF9UNW3L78oeH4T5CZJy6ssxjOueYXqQae_"

    knz_weiterer_schleifendurchlauf = True
    
    anzahl_zeichen = Len(pString)
    
    akt_position = pPositionStart

    While (akt_position <= anzahl_zeichen) And (knz_weiterer_schleifendurchlauf)

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

    While (akt_position > 0) And (knz_weiterer_schleifendurchlauf)

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


Public Function getBlockZufall(pString As String, pIndexAb As Long, pIndexBis As Long) As String

On Error GoTo errGetBlockZufall

Dim str_ergebnis           As String  ' Speichert das Funktionsergebnis
Dim akt_zeichen_str        As String  ' das aktuelle Zeichen an der Leseposition
Dim akt_zeichen_ascii_wert As Long ' ASCII-Wert des Zeichens an der Leseposition
Dim akt_index              As Long ' Aktuelle Leseposition der Eingabe
Dim index_ab               As Long ' Startindex fuer Umtauschvorgaenge
Dim index_bis              As Long ' Endindex fuer Umtauschvorgaenge
Dim index_zufall           As Long ' Zufallsindex fuer das neue Zeichen
Dim laenge_eingabe         As Long ' Laenge der Eingabe
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
            str_ergebnis = str_ergebnis & akt_zeichen_str
        
            akt_index = akt_index + 1
        
        Wend
    
    End If

EndFunktion:

    On Error Resume Next

    DoEvents

    getBlockZufall = str_ergebnis

    Exit Function

errGetBlockZufall:

    Debug.Print ("Fehler: errGetBlockZufall: " & Err & " " & Error & " " & Erl)

    Resume EndFunktion

End Function

'################################################################################
'
Public Function getSwitchCase(pString As String, pIndexAb As Long, pIndexBis As Long) As String

On Error GoTo errGetSwitchCase

Dim str_ergebnis           As String  ' Speichert das Funktionsergebnis
Dim akt_zeichen_str        As String  ' das aktuelle Zeichen an der Leseposition
Dim akt_zeichen_ascii_wert As Long ' ASCII-Wert des Zeichens an der Leseposition
Dim akt_index              As Long ' Aktuelle Leseposition der Eingabe
Dim index_ab               As Long ' Startindex fuer Umtauschvorgaenge
Dim index_bis              As Long ' Endindex fuer Umtauschvorgaenge
Dim laenge_eingabe         As Long ' Laenge der Eingabe

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
            str_ergebnis = str_ergebnis & akt_zeichen_str
        
            akt_index = akt_index + 1
        
        Wend
    
    End If

EndFunktion:

    On Error Resume Next

    DoEvents

    getSwitchCase = str_ergebnis

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

Dim akt_position_wort_start   As Long   ' aktuelle Startposition ( gleichzeitig aktuelle Position des Leseprozesses )
Dim akt_position_wort_ende    As Long   ' aktuelle Endposition
Dim laenge_eingabe            As Long   ' Laenge des Eingabestrings
Dim laenge_such_string        As Long   ' Laenge des Suchstrings
Dim akt_trennzeichen          As String
Dim ergebnis_str              As String
Dim knz_wortende_gefunden     As Integer
Dim zeichen_wortbestandteil   As String ' Gueltige Zeichen fuer ein Wort

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
    akt_trennzeichen = LEER_STRING
    '
    ' Ergebnisstring mit einem Leerstring vorbelegen
    '
    ergebnis_str = LEER_STRING
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
        ' ( Das hinzuaddieren der Suchstringlaenge koennte auch seperat ausserhalb
        ' der Funktion "InStr" gemacht werden. )
        '
        akt_position_wort_start = 1

        Do
            '
            ' Schritt 1: Startposition ermitteln
            ' Die ( naechste ) Startposition fuer das Suchwort wird ab der aktuellen
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
                ' Kennzeichen fuer "Wortende gefunden" auf 0 stellen ( = noch nicht erreicht )
                '
                knz_wortende_gefunden = 0
                '
                ' Die Suchschleife wird solange durchlaufen wie
                ' ... die aktuelle Wortendeposition noch kleiner als die Laenge der Eingabe ist
                ' ... die Flagvariable einen weiteren Schleifendurchlauf erzwingt.
                '
                While (akt_position_wort_ende <= laenge_eingabe) And (knz_wortende_gefunden = 0)
        
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
                ergebnis_str = ergebnis_str & akt_trennzeichen & Mid(pEingabeString, akt_position_wort_start, (akt_position_wort_ende - akt_position_wort_start) + 1)
                '
                ' Erst nach dem ersten ermittelten Wort wird das aktuelle Trennzeichen auf
                ' das uebergebene Trennzeichen gesetzt. Dieses verhindert ein Trennzeichen
                ' vor dem ersten Wort.
                '
                akt_trennzeichen = pTrennzeichen
        
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
    ' Der aufgebaute ergebnis-String wird zurueckgegeben.
    ' Eine explizite Fehlerbehandlung wird nicht gemacht.
    '
    getGrepSuchwort = ergebnis_str

End Function

'################################################################################
'
' http://de.wikipedia.org/wiki/ROT13
'
' ROT13 ( engl. rotate by 13 places, zu Deutsch in etwa "rotiere um 13 Stellen" )
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

    Dim akt_index As Long
    Dim ergebnis  As String

    ergebnis = LEER_STRING

    For akt_index = 1 To Len(pString)

        Select Case UCase(Mid(pString, akt_index, 1))

            Case "A" To "M"
                
                ergebnis = ergebnis & Chr(Asc(Mid(pString, akt_index, 1)) + 13)

            Case "N" To "Z"
                
                ergebnis = ergebnis & Chr(Asc(Mid(pString, akt_index, 1)) - 13)

            Case Else
                
                ergebnis = ergebnis & Mid(pString, akt_index, 1)

        End Select

    Next

    rot13 = ergebnis

End Function

'############################################################################################
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

Dim zeichen_zeilenumbruch              As String  ' das ermittelte Trennzeichen (bzw. eben Zeilenumbruchzeichen)
Dim zeilen_zaehler                     As Long    ' Zaehler fuer Vermeidung von Endlosschleifen
Dim aktuelle_zeile                     As String  ' die aktuell gefundene Zeile aus der Eingabe
Dim aktuelle_startposition             As Long    ' die akutelle Start-Leseposition
Dim naechste_position                  As Long    ' Position des naechsen gefundenen Trennzeichens
Dim knz_weiterer_schleifendurchlauf    As Boolean ' Kennzeichen ob ein weiterer Schleifendurchlauf notwendig ist
Dim aktuelles_such_wort                As String  ' der aktuell zu suchende String in der Eingabe
Dim aktuelles_ersetz_wort              As String  ' der aktuelle Ersatzstring
Dim pos_trennzeichen                   As Long
    '
    ' Ermittlung welches Zeilenumbruchzeichen in der Eingabe verwendet wird
    '
    zeichen_zeilenumbruch = Chr(13) & Chr(10)

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
    While (knz_weiterer_schleifendurchlauf) And (zeilen_zaehler < 32220)
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
            ' Laenge des Trennzeichens ( hier = Zeilenumbruchzeichen )
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
                If (aktuelles_such_wort <> "") Then
                    
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

'############################################################################################
'
Public Function placeStringX(pStringA As String, pStringB As String, pFunktion As Integer, pSelStart As Long, pSelLength As Long) As String
                
    placeStringX = LEER_STRING
                
                
On Error GoTo errStartPlaceStringX
         
Dim zeichen_zeilenumbruch          As String
Dim zeichen_zeilenumbruch_t  As String
Dim cls_string_array               As clsStringArray
Dim ergebnis_fkt                   As String
Dim aktuelle_zeile                 As String
Dim temp_string_1                  As String
Dim temp_string_2                  As String
Dim temp_string_3                  As String
Dim ab_position                    As Long
Dim zeilen_zaehler                 As Long
Dim zeilen_anzahl                  As Long
Dim temp_long_1                    As Long
Dim knz_benutze_markierung         As Boolean
Dim knz_schleifen_durchlauf        As Boolean
Dim cls_string_array_a                As clsStringArray
Dim cls_string_array_b                As clsStringArray
Dim akutelle_zeile                    As String
Dim zeilen_zaehler_b                  As Long
Dim zeilen_anzahl_b                   As Long

    Set cls_string_array_a = startMultiline(pStringA)
    Set cls_string_array_b = startMultiline(pStringB)
    
    If ((cls_string_array_a Is Nothing) Or (cls_string_array_b Is Nothing)) Then
    
    ' keine Aktionen machen
    
    ElseIf (cls_string_array_a Is Nothing) Then

        placeStringX = pStringB

    ElseIf (cls_string_array_b Is Nothing) Then

        placeStringX = pStringA
    
    Else

        zeichen_zeilenumbruch_t = getBenutztesChr13(pStringA)

        If (pSelStart >= 0) Then

            temp_long_1 = getLetztePositionVorPos(pStringA, zeichen_zeilenumbruch, pSelStart)
            
            If (temp_long_1 > 0) Then
                ab_position = (pSelStart - temp_long_1)
            Else
                ab_position = pSelStart + 1
            End If
            
        Else
        
            ab_position = 0
        
        End If
        
        zeilen_anzahl = cls_string_array_a.getAnzahlStrings
        zeilen_zaehler = 1
        zeilen_anzahl_b = cls_string_array_b.getAnzahlStrings
        zeilen_zaehler_b = 1
        
        While ((zeilen_zaehler <= zeilen_anzahl) And (zeilen_zaehler_b <= zeilen_anzahl_b))
        
            aktuelle_zeile = (cls_string_array_a.getString(zeilen_zaehler))
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
            
            ergebnis_fkt = ergebnis_fkt & zeichen_zeilenumbruch & temp_string_1 & temp_string_3 & temp_string_2
            
            zeichen_zeilenumbruch = zeichen_zeilenumbruch_t
            
            zeilen_zaehler = zeilen_zaehler + 1
            
            zeilen_zaehler_b = zeilen_zaehler_b + 1
        
        Wend
            
    End If
     
    placeStringX = ergebnis_fkt

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

Public Function generatorVbNachJava(pString As String) As String

    generatorVbNachJava = LEER_STRING
    
    If (Trim(pString) = LEER_STRING) Then
        Exit Function
    End If

' VB   aktuelle_zeile = Mid( pString, aktuelle_startposition, ( FkString.len( pString ) - aktuelle_startposition ) + 1 );
' Java aktuelle_zeile = pString.substring( aktuelle_startposition, pString.length() );

Dim cls_z                  As clsStringArray
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

    Set cls_z = startMultiline(vb_str)

pos_kommentar = -1
    akt_index = 1
    anzahl_zeilen = cls_z.getAnzahlStrings()
    zaehler_kommentar = 0
    
    While (akt_index <= anzahl_zeilen)
        
        append_to_akt_zeile = LEER_STRING
        knz_bearbeiten = True
        
        akt_zeile = cls_z.getString(akt_index)
        
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
                
                temp_string_1 = cls_z.getString(akt_index + 1)
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

    Call cls_z.setString(akt_index, akt_zeile & append_to_akt_zeile)

    akt_index = akt_index + 1

Wend


    vb_str = cls_z.toString(zeichen_zeilenumbruch)
    
    Set cls_z = Nothing
    
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

Dim buf As String
Dim aktuelles_zeichen As String
Dim trennzeichen_nk As String
Dim knz_negativ As Boolean
Dim knz_nk_aktiv As Integer
Dim zaehler As Integer
Dim zaehler_nk As Integer
Dim ziffern_zaehler As Integer
    
    trennzeichen_nk = ","
    knz_negativ = False
    knz_nk_aktiv = 0
    zaehler = 1
    zaehler_nk = 0
    ziffern_zaehler = 0
    buf = LEER_STRING


    If (Len(pString) > 0) Then
        '
        ' Hier wird ermittelt, ob das Nachkommatrennzeichen auf einen Punkt
        ' geaendert werden muss. Per Vorgabe wird das Komma als Trennzeichen
        ' genommen. Wenn sich in der Eingabe kein Komma findet, dann wird
        ' der Punkt genommen.
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
                ' Wenn der Zaehler fuer die Nachkommastellen kleiner als die gewuenschte
                ' Anzahl der Nachkommastellen ist, wird die aktuelle Zahl dem Ergebnis
                ' hinzugefuegt.
                '
                ' Dieses wird auch dann gemacht, wenn die Anzahl der gewuenschten
                ' Nachkommastellen 0 ist. In einem solchen Fall wird dem Aufrufer
                ' die Eingabe nur in eine Zahl konvertiert.
                '
                If ((zaehler_nk < pAnzahlNachkommaStellen) Or (pAnzahlNachkommaStellen < 0)) Then
                
                    buf = buf & aktuelles_zeichen
                
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
                
                    buf = buf & "."
                    
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
        
            buf = LEER_STRING
            
        End If

    End If
    '
    ' Wenn die Eingabe null, ein Leerstring, oder durch vorangegangene
    ' Abfragen wieder ausgenullt worden ist, ist die Laenge des String-
    ' Buffers 0. Damit hier jetzt eine korrekte Zahl erstellt werden
    ' kann, wird eine fuehrende 0 hinzugefuegt.
    '
    If (Len(buf) = 0) Then
    
        buf = "0"
        
    End If

    '
    ' Hier werden die gewuenschten Anzahl der Nachkommastellen hinzugefuegt.
    '
    ' Wenn der Nachkommastellenzaehler noch 0 ist, muss noch ein Punkt hinzugefuegt werden
    '
    While (zaehler_nk < pAnzahlNachkommaStellen)
    
        If ((zaehler_nk = 0) And (knz_nk_aktiv = 0)) Then
        
            buf = buf & "."
            
        End If
        
        buf = buf & "0"
        
        zaehler_nk = zaehler_nk + 1
    Wend

    '
    ' Bei der Rueckgabe wird noch das Kennzeichen fuer einen negativen Betrag
    ' ausgewertet und gegebenenfalls ein Bindestrich dem Ergebnis hinzugefuegt.
    '
    getzahl = IIf(knz_negativ, "-", "") + buf

End Function


'################################################################################
'
Private Function getStringMaxCols(pEingabe As String, pMaxAnzahlSpalten As Long, pEinzug As String, pNewLineZeichen As String) As String

Dim ergebnis             As String
Dim my_cr                As String
Dim neue_zeile           As String
Dim trenn_position_ab    As Long
Dim trenn_position_bis   As Long
Dim trenn_position_temp  As Long
Dim laenge_eingabe       As Long
Dim zaehler              As Long

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

        ergebnis = pEingabe

    Else
        '
        ' Ist die Eingabe laenger als die maximale Spaltenanzahl wird die
        ' Verkleinerungsschleife gestartet.
        '
        ergebnis = LEER_STRING

        laenge_eingabe = Len(pEingabe)
        trenn_position_ab = 0
        trenn_position_bis = 0
        zaehler = 0

        '
        ' Die Schleife laeuft solange wie
        ' ... die aktuelle Bis-Position noch kleiner der Laenge der Eingabe ist.
        ' ... der Endlosschleifenverhinderungszaehler kleiner 32123 ist.
        '
        While ((trenn_position_bis < laenge_eingabe) And (zaehler < 32123))
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
            ' Jetzt ist es jedeoch denkbar, dass eventuell X-Zeichen ( hier 3 Zeichen )
            ' vor dem eigentlichen Abschneide-End ein Leerzeichen liegt, welches
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

            If (neue_zeile <> "") Then

                ergebnis = ergebnis & my_cr & neue_zeile

            End If

            my_cr = pNewLineZeichen & pEinzug

            zaehler = zaehler + 1

        Wend

    End If

    getStringMaxCols = ergebnis

End Function

'################################################################################
'
Private Function extrahiereWoerter(pText As String, pTrennzeichen As String, pMaxLaenge As Long) As String

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
    
    If (pMaxLaenge < 0) Then
    
        ergebnis_max_zeilenlaenge = 150
    
    Else
    
        ergebnis_max_zeilenlaenge = pMaxLaenge
        
    End If
    
    '
    ' Unterstrich kein Trennzeichen, da in Variablennamen benutzt
    '
    parser_wort_trennzeichen = "[]{}()-+:=\/:*?!#<> |.,;&""" & vbCr & vbTab & vbLf
    
    parser_akt_position = 1

    While (parser_akt_position <= Len(pText))
        
        parser_akt_zeichen = Mid(pText, parser_akt_position, 1)
        
        If (InStr(parser_wort_trennzeichen, parser_akt_zeichen) = 0) And (Asc(parser_akt_zeichen) > 30) Then
        
           parser_temp_wort = parser_temp_wort + parser_akt_zeichen
        
        Else
        
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
    fkt_ergebnis = Replace(fkt_ergebnis, "%DF", "ß")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E4", "ä")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F6", "ö")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FC", "ü")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C4", "Ä")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D6", "Ö")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DC", "Ü")
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
    fkt_ergebnis = Replace(fkt_ergebnis, "%5F", "_")
    fkt_ergebnis = Replace(fkt_ergebnis, "%60", "`")
    fkt_ergebnis = Replace(fkt_ergebnis, "%7B", "{")
    fkt_ergebnis = Replace(fkt_ergebnis, "%7C", "|")
    fkt_ergebnis = Replace(fkt_ergebnis, "%7D", "}")
    fkt_ergebnis = Replace(fkt_ergebnis, "%7E", "~")
    fkt_ergebnis = Replace(fkt_ergebnis, "%7F", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "%80", "`")
    fkt_ergebnis = Replace(fkt_ergebnis, "%81", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "%82", "‚")
    fkt_ergebnis = Replace(fkt_ergebnis, "%83", "ƒ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%84", "„")
    fkt_ergebnis = Replace(fkt_ergebnis, "%85", "…")
    fkt_ergebnis = Replace(fkt_ergebnis, "%86", "†")
    fkt_ergebnis = Replace(fkt_ergebnis, "%87", "‡")
    fkt_ergebnis = Replace(fkt_ergebnis, "%88", "ˆ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%89", "‰")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8A", "Š")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8B", "‹")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8C", "Œ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8D", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8E", "Ž")
    fkt_ergebnis = Replace(fkt_ergebnis, "%8F", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "%90", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "%91", "‘")
    fkt_ergebnis = Replace(fkt_ergebnis, "%92", "’")
    fkt_ergebnis = Replace(fkt_ergebnis, "%93", "“")
    fkt_ergebnis = Replace(fkt_ergebnis, "%94", "”")
    fkt_ergebnis = Replace(fkt_ergebnis, "%95", "•")
    fkt_ergebnis = Replace(fkt_ergebnis, "%96", "–")
    fkt_ergebnis = Replace(fkt_ergebnis, "%97", "—")
    fkt_ergebnis = Replace(fkt_ergebnis, "%98", "˜")
    fkt_ergebnis = Replace(fkt_ergebnis, "%99", "™")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9A", "š")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9B", "›")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9C", "œ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9D", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9E", "ž")
    fkt_ergebnis = Replace(fkt_ergebnis, "%9F", "Ÿ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A0", "")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A1", "¡")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A2", "¢")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A3", "£")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A4", "¤")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A5", "¥")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A6", "¦")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A7", "§")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A8", "¨")
    fkt_ergebnis = Replace(fkt_ergebnis, "%A9", "©")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AA", "ª")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AB", "«")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AC", "¬")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AD", "­")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AE", "®")
    fkt_ergebnis = Replace(fkt_ergebnis, "%AF", "¯")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B0", "°")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B1", "±")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B2", "²")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B3", "³")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B4", "´")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B5", "µ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B6", "¶")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B7", "·")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B8", "¸")
    fkt_ergebnis = Replace(fkt_ergebnis, "%B9", "¹")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BA", "º")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BB", "»")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BC", "¼")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BD", "½")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BE", "¾")
    fkt_ergebnis = Replace(fkt_ergebnis, "%BF", "¿")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C0", "À")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C1", "Á")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C2", "Â")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C3", "Ã")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C4", "Ä")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C5", "Å")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C6", "Æ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C7", "Ç")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C8", "È")
    fkt_ergebnis = Replace(fkt_ergebnis, "%C9", "É")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CA", "Ê")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CB", "Ë")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CC", "Ì")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CD", "Í")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CE", "Î")
    fkt_ergebnis = Replace(fkt_ergebnis, "%CF", "Ï")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D0", "Ð")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D1", "Ñ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D2", "Ò")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D3", "Ó")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D4", "Ô")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D5", "Õ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D6", "Ö")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D7", "×")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D8", "Ø")
    fkt_ergebnis = Replace(fkt_ergebnis, "%D9", "Ù")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DA", "Ú")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DB", "Û")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DC", "Ü")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DD", "Ý")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DE", "Þ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%DF", "ß")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E0", "à")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E1", "á")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E2", "â")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E3", "ã")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E4", "ä")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E5", "å")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E6", "æ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E7", "ç")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E8", "è")
    fkt_ergebnis = Replace(fkt_ergebnis, "%E9", "é")
    fkt_ergebnis = Replace(fkt_ergebnis, "%EA", "ê")
    fkt_ergebnis = Replace(fkt_ergebnis, "%EB", "ë")
    fkt_ergebnis = Replace(fkt_ergebnis, "%EC", "ì")
    fkt_ergebnis = Replace(fkt_ergebnis, "%ED", "í")
    fkt_ergebnis = Replace(fkt_ergebnis, "%EE", "î")
    fkt_ergebnis = Replace(fkt_ergebnis, "%EF", "ï")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F0", "ð")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F1", "ñ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F2", "ò")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F3", "ó")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F4", "ô")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F5", "õ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F6", "ö")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F7", "÷")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F8", "ø")
    fkt_ergebnis = Replace(fkt_ergebnis, "%F9", "ù")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FA", "ú")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FB", "û")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FC", "ü")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FD", "ý")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FE", "þ")
    fkt_ergebnis = Replace(fkt_ergebnis, "%FF", "ÿ")

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

Dim str_fkt_ergebnis     As String    ' Ergebnisstring fuer die Rueckgabe
Dim str_gueltige_zeichen  As String ' Liste der gueltigen Zeichen
Dim akt_zeichen      As String    ' aktuelle Zeichen in der While-Schleife
Dim akt_position     As Integer   ' aktuelle Leseposition der While-Schleife

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

        ElseIf (akt_zeichen = "ß") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DF"

        ElseIf (akt_zeichen = "ä") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E4"

        ElseIf (akt_zeichen = "ö") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F6"

        ElseIf (akt_zeichen = "ü") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FC"

        ElseIf (akt_zeichen = "Ä") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C4"

        ElseIf (akt_zeichen = "Ö") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D6"

        ElseIf (akt_zeichen = "Ü") Then

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

        ElseIf (akt_zeichen = "_") Then

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

        ElseIf (akt_zeichen = "") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%81"

        ElseIf (akt_zeichen = "‚") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%82"

        ElseIf (akt_zeichen = "ƒ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%83"

        ElseIf (akt_zeichen = "„") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%84"

        ElseIf (akt_zeichen = "…") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%85"

        ElseIf (akt_zeichen = "†") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%86"

        ElseIf (akt_zeichen = "‡") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%87"

        ElseIf (akt_zeichen = "ˆ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%88"

        ElseIf (akt_zeichen = "‰") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%89"

        ElseIf (akt_zeichen = "Š") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8A"

        ElseIf (akt_zeichen = "‹") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8B"

        ElseIf (akt_zeichen = "Œ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8C"

        ElseIf (akt_zeichen = "") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8D"

        ElseIf (akt_zeichen = "Ž") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8E"

        ElseIf (akt_zeichen = "") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%8F"

        ElseIf (akt_zeichen = "") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%90"

        ElseIf (akt_zeichen = "‘") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%91"

        ElseIf (akt_zeichen = "’") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%92"

        ElseIf (akt_zeichen = "“") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%93"

        ElseIf (akt_zeichen = "”") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%94"

        ElseIf (akt_zeichen = "•") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%95"

        ElseIf (akt_zeichen = "–") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%96"

        ElseIf (akt_zeichen = "—") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%97"

        ElseIf (akt_zeichen = "˜") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%98"

        ElseIf (akt_zeichen = "™") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%99"

        ElseIf (akt_zeichen = "š") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9A"

        ElseIf (akt_zeichen = "›") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9B"

        ElseIf (akt_zeichen = "œ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9C"

        ElseIf (akt_zeichen = "") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9D"

        ElseIf (akt_zeichen = "ž") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9E"

        ElseIf (akt_zeichen = "Ÿ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%9F"

        ElseIf (akt_zeichen = "") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A0"

        ElseIf (akt_zeichen = "¡") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A1"

        ElseIf (akt_zeichen = "¢") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A2"

        ElseIf (akt_zeichen = "£") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A3"

        ElseIf (akt_zeichen = "¤") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A4"

        ElseIf (akt_zeichen = "¥") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A5"

        ElseIf (akt_zeichen = "¦") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A6"

        ElseIf (akt_zeichen = "§") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A7"

        ElseIf (akt_zeichen = "¨") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A8"

        ElseIf (akt_zeichen = "©") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%A9"

        ElseIf (akt_zeichen = "ª") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AA"

        ElseIf (akt_zeichen = "«") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AB"

        ElseIf (akt_zeichen = "¬") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AC"

        ElseIf (akt_zeichen = "­") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AD"

        ElseIf (akt_zeichen = "®") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AE"

        ElseIf (akt_zeichen = "¯") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%AF"

        ElseIf (akt_zeichen = "°") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B0"

        ElseIf (akt_zeichen = "±") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B1"

        ElseIf (akt_zeichen = "²") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B2"

        ElseIf (akt_zeichen = "³") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B3"

        ElseIf (akt_zeichen = "´") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B4"

        ElseIf (akt_zeichen = "µ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B5"

        ElseIf (akt_zeichen = "¶") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B6"

        ElseIf (akt_zeichen = "·") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B7"

        ElseIf (akt_zeichen = "¸") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B8"

        ElseIf (akt_zeichen = "¹") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%B9"

        ElseIf (akt_zeichen = "º") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BA"

        ElseIf (akt_zeichen = "»") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BB"

        ElseIf (akt_zeichen = "¼") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BC"

        ElseIf (akt_zeichen = "½") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BD"

        ElseIf (akt_zeichen = "¾") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BE"

        ElseIf (akt_zeichen = "¿") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%BF"

        ElseIf (akt_zeichen = "À") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C0"

        ElseIf (akt_zeichen = "Á") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C1"

        ElseIf (akt_zeichen = "Â") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C2"

        ElseIf (akt_zeichen = "Ã") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C3"

        ElseIf (akt_zeichen = "Ä") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C4"

        ElseIf (akt_zeichen = "Å") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C5"

        ElseIf (akt_zeichen = "Æ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C6"

        ElseIf (akt_zeichen = "Ç") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C7"

        ElseIf (akt_zeichen = "È") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C8"

        ElseIf (akt_zeichen = "É") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%C9"

        ElseIf (akt_zeichen = "Ê") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CA"

        ElseIf (akt_zeichen = "Ë") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CB"

        ElseIf (akt_zeichen = "Ì") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CC"

        ElseIf (akt_zeichen = "Í") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CD"

        ElseIf (akt_zeichen = "Î") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CE"

        ElseIf (akt_zeichen = "Ï") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%CF"

        ElseIf (akt_zeichen = "Ð") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D0"

        ElseIf (akt_zeichen = "Ñ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D1"

        ElseIf (akt_zeichen = "Ò") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D2"

        ElseIf (akt_zeichen = "Ó") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D3"

        ElseIf (akt_zeichen = "Ô") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D4"

        ElseIf (akt_zeichen = "Õ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D5"

        ElseIf (akt_zeichen = "Ö") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D6"

        ElseIf (akt_zeichen = "×") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D7"

        ElseIf (akt_zeichen = "Ø") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D8"

        ElseIf (akt_zeichen = "Ù") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%D9"

        ElseIf (akt_zeichen = "Ú") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DA"

        ElseIf (akt_zeichen = "Û") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DB"

        ElseIf (akt_zeichen = "Ü") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DC"

        ElseIf (akt_zeichen = "Ý") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DD"

        ElseIf (akt_zeichen = "Þ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DE"

        ElseIf (akt_zeichen = "ß") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%DF"

        ElseIf (akt_zeichen = "à") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E0"

        ElseIf (akt_zeichen = "á") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E1"

        ElseIf (akt_zeichen = "â") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E2"

        ElseIf (akt_zeichen = "ã") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E3"

        ElseIf (akt_zeichen = "ä") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E4"

        ElseIf (akt_zeichen = "å") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E5"

        ElseIf (akt_zeichen = "æ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E6"

        ElseIf (akt_zeichen = "ç") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E7"

        ElseIf (akt_zeichen = "è") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E8"

        ElseIf (akt_zeichen = "é") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%E9"

        ElseIf (akt_zeichen = "ê") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%EA"

        ElseIf (akt_zeichen = "ë") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%EB"

        ElseIf (akt_zeichen = "ì") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%EC"

        ElseIf (akt_zeichen = "í") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%ED"

        ElseIf (akt_zeichen = "î") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%EE"

        ElseIf (akt_zeichen = "ï") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%EF"

        ElseIf (akt_zeichen = "ð") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F0"

        ElseIf (akt_zeichen = "ñ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F1"

        ElseIf (akt_zeichen = "ò") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F2"

        ElseIf (akt_zeichen = "ó") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F3"

        ElseIf (akt_zeichen = "ô") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F4"

        ElseIf (akt_zeichen = "õ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F5"

        ElseIf (akt_zeichen = "ö") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F6"

        ElseIf (akt_zeichen = "÷") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F7"

        ElseIf (akt_zeichen = "ø") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F8"

        ElseIf (akt_zeichen = "ù") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%F9"

        ElseIf (akt_zeichen = "ú") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FA"

        ElseIf (akt_zeichen = "û") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FB"

        ElseIf (akt_zeichen = "ü") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FC"

        ElseIf (akt_zeichen = "ý") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FD"

        ElseIf (akt_zeichen = "þ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FE"

        ElseIf (akt_zeichen = "ÿ") Then

            str_fkt_ergebnis = str_fkt_ergebnis & "%FF"

        Else
        
            '
            ' Pruefung: Gueltiges Zeichen
            ' Befindet sich das aktuelle Zeichen in dem String der gueltigen
            ' Zeichen, wird das Zeichen in den Ergebnisstring uebernommen.
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

    If (pHtmlString <> "") Then
    
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
        While (akt_index < anzahl_zeichen)
        
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
            
                ElseIf (akt_zeichen = "¡") Then
            
                    akt_zeichen = "&iexcl;"
            
                ElseIf (akt_zeichen = "¢") Then
            
                    akt_zeichen = "&cent;"
            
                ElseIf (akt_zeichen = "£") Then
            
                    akt_zeichen = "&pound;"
            
                ElseIf (akt_zeichen = "¤") Then
            
                    akt_zeichen = "&curren;"
            
                ElseIf (akt_zeichen = "¥") Then
            
                    akt_zeichen = "&yen;"
            
                ElseIf (akt_zeichen = "¦") Then
            
                    akt_zeichen = "&brvbar;"
            
                ElseIf (akt_zeichen = "§") Then
            
                    akt_zeichen = "&sect;"
            
                ElseIf (akt_zeichen = "¨") Then
            
                    akt_zeichen = "&uml;"
            
                ElseIf (akt_zeichen = "©") Then
            
                    akt_zeichen = "&copy;"
            
                ElseIf (akt_zeichen = "ª") Then
            
                    akt_zeichen = "&ordf;"
            
                ElseIf (akt_zeichen = "«") Then
            
                    akt_zeichen = "&laquo;"
            
                ElseIf (akt_zeichen = "¬") Then
            
                    akt_zeichen = "&not;"
            
                ElseIf (akt_zeichen = "­") Then
            
                    akt_zeichen = "&shy;"
            
                ElseIf (akt_zeichen = "®") Then
            
                    akt_zeichen = "&reg;"
            
                ElseIf (akt_zeichen = "¯") Then
            
                    akt_zeichen = "&macr;"
            
                ElseIf (akt_zeichen = "°") Then
            
                    akt_zeichen = "&deg;"
            
                ElseIf (akt_zeichen = "±") Then
            
                    akt_zeichen = "&plusmn;"
            
                ElseIf (akt_zeichen = "²") Then
            
                    akt_zeichen = "&sup2;"
            
                ElseIf (akt_zeichen = "³") Then
            
                    akt_zeichen = "&sup3;"
            
                ElseIf (akt_zeichen = "´") Then
            
                    akt_zeichen = "&acute;"
            
                ElseIf (akt_zeichen = "µ") Then
            
                    akt_zeichen = "&micro;"
            
                ElseIf (akt_zeichen = "¶") Then
            
                    akt_zeichen = "&para;"
            
                ElseIf (akt_zeichen = "·") Then
            
                    akt_zeichen = "&middot;"
            
                ElseIf (akt_zeichen = "¸") Then
            
                    akt_zeichen = "&cedil;"
            
                ElseIf (akt_zeichen = "¹") Then
            
                    akt_zeichen = "&sup1;"
            
                ElseIf (akt_zeichen = "º") Then
            
                    akt_zeichen = "&ordm;"
            
                ElseIf (akt_zeichen = "»") Then
            
                    akt_zeichen = "&raquo;"
            
                ElseIf (akt_zeichen = "¼") Then
            
                    akt_zeichen = "&frac14;"
            
                ElseIf (akt_zeichen = "½") Then
            
                    akt_zeichen = "&frac12;"
            
                ElseIf (akt_zeichen = "¾") Then
            
                    akt_zeichen = "&frac34;"
            
                ElseIf (akt_zeichen = "¿") Then
            
                    akt_zeichen = "&iquest;"
            
                ElseIf (akt_zeichen = "À") Then
            
                    akt_zeichen = "&Agrave;"
            
                ElseIf (akt_zeichen = "Á") Then
            
                    akt_zeichen = "&Aacute;"
            
                ElseIf (akt_zeichen = "Â") Then
            
                    akt_zeichen = "&Acirc;"
            
                ElseIf (akt_zeichen = "Ã") Then
            
                    akt_zeichen = "&Atilde;"
            
                ElseIf (akt_zeichen = "Ä") Then
            
                    akt_zeichen = "&Auml;"
            
                ElseIf (akt_zeichen = "Å") Then
            
                    akt_zeichen = "&Aring;"
            
                ElseIf (akt_zeichen = "Æ") Then
            
                    akt_zeichen = "&AElig;"
            
                ElseIf (akt_zeichen = "Ç") Then
            
                    akt_zeichen = "&Ccedil;"
            
                ElseIf (akt_zeichen = "È") Then
            
                    akt_zeichen = "&Egrave;"
            
                ElseIf (akt_zeichen = "É") Then
            
                    akt_zeichen = "&Eacute;"
            
                ElseIf (akt_zeichen = "Ê") Then
            
                    akt_zeichen = "&Ecirc;"
            
                ElseIf (akt_zeichen = "Ë") Then
            
                    akt_zeichen = "&Euml;"
            
                ElseIf (akt_zeichen = "Ì") Then
            
                    akt_zeichen = "&Igrave;"
            
                ElseIf (akt_zeichen = "Í") Then
            
                    akt_zeichen = "&Iacute;"
            
                ElseIf (akt_zeichen = "Î") Then
            
                    akt_zeichen = "&Icirc;"
            
                ElseIf (akt_zeichen = "Ï") Then
            
                    akt_zeichen = "&Iuml;"
            
                ElseIf (akt_zeichen = "Ð") Then
            
                    akt_zeichen = "&ETH;"
            
                ElseIf (akt_zeichen = "Ñ") Then
            
                    akt_zeichen = "&Ntilde;"
            
                ElseIf (akt_zeichen = "Ò") Then
            
                    akt_zeichen = "&Ograve;"
            
                ElseIf (akt_zeichen = "Ó") Then
            
                    akt_zeichen = "&Oacute;"
            
                ElseIf (akt_zeichen = "Ô") Then
            
                    akt_zeichen = "&Ocirc;"
            
                ElseIf (akt_zeichen = "Õ") Then
            
                    akt_zeichen = "&Otilde;"
            
                ElseIf (akt_zeichen = "Ö") Then
            
                    akt_zeichen = "&Ouml;"
            
                ElseIf (akt_zeichen = "×") Then
            
                    akt_zeichen = "&times;"
            
                ElseIf (akt_zeichen = "Ø") Then
            
                    akt_zeichen = "&Oslash;"
            
                ElseIf (akt_zeichen = "Ù") Then
            
                    akt_zeichen = "&Ugrave;"
            
                ElseIf (akt_zeichen = "Ú") Then
            
                    akt_zeichen = "&Uacute;"
            
                ElseIf (akt_zeichen = "Û") Then
            
                    akt_zeichen = "&Ucirc;"
            
                ElseIf (akt_zeichen = "Ü") Then
            
                    akt_zeichen = "&Uuml;"
            
                ElseIf (akt_zeichen = "Ý") Then
            
                    akt_zeichen = "&Yacute;"
            
                ElseIf (akt_zeichen = "Þ") Then
            
                    akt_zeichen = "&THORN;"
            
                ElseIf (akt_zeichen = "ß") Then
            
                    akt_zeichen = "&szlig;"
            
                ElseIf (akt_zeichen = "à") Then
            
                    akt_zeichen = "&agrave;"
            
                ElseIf (akt_zeichen = "á") Then
            
                    akt_zeichen = "&aacute;"
            
                ElseIf (akt_zeichen = "â") Then
            
                    akt_zeichen = "&acirc;"
            
                ElseIf (akt_zeichen = "ã") Then
            
                    akt_zeichen = "&atilde;"
            
                ElseIf (akt_zeichen = "ä") Then
            
                    akt_zeichen = "&auml;"
            
                ElseIf (akt_zeichen = "å") Then
            
                    akt_zeichen = "&aring;"
            
                ElseIf (akt_zeichen = "æ") Then
            
                    akt_zeichen = "&aelig;"
            
                ElseIf (akt_zeichen = "ç") Then
            
                    akt_zeichen = "&ccedil;"
            
                ElseIf (akt_zeichen = "è") Then
            
                    akt_zeichen = "&egrave;"
            
                ElseIf (akt_zeichen = "é") Then
            
                    akt_zeichen = "&eacute;"
            
                ElseIf (akt_zeichen = "ê") Then
            
                    akt_zeichen = "&ecirc;"
            
                ElseIf (akt_zeichen = "ë") Then
            
                    akt_zeichen = "&euml;"
            
                ElseIf (akt_zeichen = "ì") Then
            
                    akt_zeichen = "&igrave;"
            
                ElseIf (akt_zeichen = "í") Then
            
                    akt_zeichen = "&iacute;"
            
                ElseIf (akt_zeichen = "î") Then
            
                    akt_zeichen = "&icirc;"
            
                ElseIf (akt_zeichen = "ï") Then
            
                    akt_zeichen = "&iuml;"
            
                ElseIf (akt_zeichen = "ð") Then
            
                    akt_zeichen = "&eth;"
            
                ElseIf (akt_zeichen = "ñ") Then
            
                    akt_zeichen = "&ntilde;"
            
                ElseIf (akt_zeichen = "ò") Then
            
                    akt_zeichen = "&ograve;"
            
                ElseIf (akt_zeichen = "ó") Then
            
                    akt_zeichen = "&oacute;"
            
                ElseIf (akt_zeichen = "ô") Then
            
                    akt_zeichen = "&ocirc;"
            
                ElseIf (akt_zeichen = "õ") Then
            
                    akt_zeichen = "&otilde;"
            
                ElseIf (akt_zeichen = "ö") Then
            
                    akt_zeichen = "&ouml;"
            
                ElseIf (akt_zeichen = "÷") Then
            
                    akt_zeichen = "&divide;"
            
                ElseIf (akt_zeichen = "ø") Then
            
                    akt_zeichen = "&oslash;"
            
                ElseIf (akt_zeichen = "ù") Then
            
                    akt_zeichen = "&ugrave;"
            
                ElseIf (akt_zeichen = "ú") Then
            
                    akt_zeichen = "&uacute;"
            
                ElseIf (akt_zeichen = "û") Then
            
                    akt_zeichen = "&ucirc;"
            
                ElseIf (akt_zeichen = "ü") Then
            
                    akt_zeichen = "&uuml;"
            
                ElseIf (akt_zeichen = "ý") Then
            
                    akt_zeichen = "&yacute;"
            
                ElseIf (akt_zeichen = "þ") Then
            
                    akt_zeichen = "&thorn;"
            
                ElseIf (akt_zeichen = "ÿ") Then
            
                    akt_zeichen = "&yuml;"
            
                ElseIf (akt_zeichen = "‘") Then
            
                    akt_zeichen = "&lsquo;"
            
                ElseIf (akt_zeichen = "’") Then
            
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



