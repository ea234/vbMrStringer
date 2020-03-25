VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Vektor fuer die Speicherung der Strings
'
Private m_array_strings()          As String

'
' Zaehler fuer die Anzahl der gespeicherten Strings. Die Vektorausdehnung
' kann groesser als die Anzahl der gespeicherten Strings sein.
'
Private m_anzahl_daten_strings     As Long

'
' Speichert die Anzahl von neuen Speicherplaetzen beim Redim
'
Private m_vektor_anzahl_erweiterung  As Long

'
' Kennzeichen fuer die Sortierung bei Spezial-Funktionen
'
Private m_knz_benutze_markierung   As Boolean
Private m_knz_umstellung_datum_machen As Boolean
Private m_ab_position              As Long
Private m_bis_position             As Long

'################################################################################
'
Private Sub Class_Initialize()
    
    Call initArrayStrings

End Sub

'################################################################################
'
Private Sub Class_Terminate()
    
    Call initArrayStrings

End Sub

'################################################################################
'
Public Sub initArrayStrings()

On Error Resume Next
    
    ReDim m_array_strings(1)
    
    m_anzahl_daten_strings = 0
    
    m_knz_benutze_markierung = False
    
    m_vektor_anzahl_erweiterung = 40
    
End Sub

'################################################################################
'
' Fuegt den String dem Vektor hinzu. Der Vektor wird gegebenenfalls erweitert,
' um den neuen String speichern zu koennen.
'
' PARAMETER: pString der zu speichernde String
'
Public Sub addString(ByVal pString As String)

    '
    ' Vektorvergroesserung und Speicheranzahl
    '
    ' Im Vektor soll eine neuer String gespeichert werden.
    '
    ' Die Anzahl der gespeicherten Strings wird um 1 erhoeht.
    '
    ' Hat der Vektor nicht mehr genug Speicherplaetze, wird der
    ' Vektor eventuell vergroessert.
    '
    ' Da die Speicherung im Vektor bei Indexposition 0 beginnt, ist
    ' der Speicherindex des neuen Elementes gleich der Anzahl - 1.
    '

    m_anzahl_daten_strings = m_anzahl_daten_strings + 1

    If UBound(m_array_strings, 1) < m_anzahl_daten_strings Then
        
        ReDim Preserve m_array_strings(m_anzahl_daten_strings + m_vektor_anzahl_erweiterung)
    
    End If
    
    m_array_strings(m_anzahl_daten_strings) = Replace(pString, Chr(9), "    ")

End Sub

'################################################################################
'
' Gibt die Anzahl der gespeicherten Strings zurueck.
'
' RETURN : die Anzahl der Strings
'
Public Function getAnzahlStrings() As Long

    getAnzahlStrings = m_anzahl_daten_strings
    
End Function

'################################################################################
'
' Liefert den String an dem uebergebenen Index zurueck.
' Liegt der Index ausserhalb der Vektoranzahl wird ein Leerstring zurueck gegeben.
'
' PARAMETER: pIndex = der Index im Vektor
'
' RETURN  : der String am am Index, oder ein Leerstring wenn der Index nicht existiert
'
Public Function getString(pIndex As Long) As String

    '
    ' Pruefung: pIndex gueltig ?
    '
    ' Der Index aus dem Parameter "pIndex" muss im Bereich
    ' von 0 bis zur aktuellen Vektoranzahl liegen.
    '
    ' Existiert der Index, wird der dort gespeicherte String zurueck gegeben
    '
    ' Existiert der Index nicht, wird ein Leerstring zurueck gegeben.
    '
    If ((pIndex > 0) And (pIndex <= m_anzahl_daten_strings)) Then
        
        getString = m_array_strings(pIndex)
    
    Else
    
        getString = ""
        
    End If
    
End Function

'################################################################################
'
Public Sub setString(pIndex As Long, pString As String)

    '
    ' Pruefung: pIndex gueltig ?
    '
    If ((pIndex > 0) And (pIndex <= m_anzahl_daten_strings)) Then
    
        '
        ' Ist der Index gueltig, wird der String an der Indexposition gesetzt.
        '
        m_array_strings(pIndex) = pString
    
    End If
    
End Sub

'################################################################################
'
' Liefert einen String der gespeicherten Strings zurueck.
'
' Mit dem Parameter "pKnzLeerzeilenEntfernen" koenen Leerzeilen uberlesen werden.
'
Public Function toString(pZeichenZeilenumbruch As String, Optional pKnzLeerzeilenEntfernen As Boolean = False) As String

Dim str_fkt_ergebnis As String ' Speichert den Rueckgabestring
Dim akt_index        As Long   ' der Index fuer die While-Schleife

    str_fkt_ergebnis = ""
    
    akt_index = 1
    
    While (akt_index <= m_anzahl_daten_strings)
        
        If (pKnzLeerzeilenEntfernen) Then
            
            If (Trim(m_array_strings(akt_index)) <> LEER_STRING) Then
                
                str_fkt_ergebnis = str_fkt_ergebnis & m_array_strings(akt_index) & pZeichenZeilenumbruch
            
            End If
        
        Else
            
            str_fkt_ergebnis = str_fkt_ergebnis & m_array_strings(akt_index) & pZeichenZeilenumbruch
        
        End If
        
        akt_index = akt_index + 1
            
    Wend

    toString = str_fkt_ergebnis
    
End Function

'################################################################################
'
Private Function getStringSortierung(pIndex As Long, pSortierkriterium As Long) As String
    
    If (pSortierkriterium = 234) Then
        '
        ' Sortierung nach Zeilenlänge
        '
        If (m_knz_benutze_markierung) Then
            
            getStringSortierung = Right("0000000000" & Len(Trim(getStringAbBis(m_array_strings(pIndex), m_ab_position, m_bis_position))), 10) & " " & m_array_strings(pIndex)
            
        Else
            
            getStringSortierung = Right("0000000000" & Len(m_array_strings(pIndex)), 10) & " " & m_array_strings(pIndex)
        
        End If
        
    Else
        '
        ' Prüfung: Soll die Markierung benutzt werden?
        '
        If (m_knz_benutze_markierung) Then
        
            If (m_knz_umstellung_datum_machen) Then
                '
                ' Markierung nach Datum
                ' Bei der Markierung nach Datum ist das Problem, dass bei der normalen Darstellung
                ' von TT.MM.JJJJ keine korrekte Sortierung möglich ist. Dieses Format muss so umge-
                ' stellt werden, dass zuerst das Jahr, dannn der Monat und dann der Tag kommt.
                '
                ' Es wird lediglich eine Umstellung der Zeichen nach einer Vorschrift gemacht.
                ' Die Vorschrift ist "aus 12.34.5678 wird 56783412"
                '
                ' Es wird keine Prüfung auf "Datum" gemacht, sondern lediglich Umstellung
                ' von TT.MM.JJJJ nach JJJJMMTT gemacht. Das ist kein Problem, da bei einer
                ' Sortierung nach Datum eigentlich nur die Zeilen mit einer korrekten Datumsangabe
                ' in Reihe gebracht werden sollen. Die anderen Zeilen sind nicht weiter von
                ' Bedeutung (Sortieurng von Datei-Namenslisten vom DIR-Befehl)
                '
                ' Der aktuelle String wird dem Datum dann hinten angestellt.
                '
                getStringSortierung = Mid(m_array_strings(pIndex), m_ab_position + 6, 4) & Mid(m_array_strings(pIndex), m_ab_position + 3, 2) & Mid(m_array_strings(pIndex), m_ab_position, 2) & m_array_strings(pIndex)
            
            Else
                '
                ' Sortierung nach dem Inhalt der Markierung
                ' Der aktuelle String wird auch hier hinten an gestellt, damit praktisch eine
                ' Untersortierung stattfindet.
                '
                getStringSortierung = getStringAbBis(m_array_strings(pIndex), m_ab_position, m_bis_position) & m_array_strings(pIndex)
            
            End If
        
        Else
            '
            ' normale Sortierung nach der Zeichenfolge
            '
            getStringSortierung = m_array_strings(pIndex)
        
        End If
    
    End If

End Function

'################################################################################
'
Private Sub QuickSort(ByVal pIndexStart As Long, ByVal pIndexEnde As Long, pSortierkriterium As Long, pKnzAufsteigend As Boolean)

Dim vergleichs_string_mitte As String
Dim temp_string             As String
Dim index_start             As Long
Dim index_ende              As Long

    index_start = pIndexStart
    index_ende = pIndexEnde
    vergleichs_string_mitte = getStringSortierung(Int((index_start + index_ende) * 0.5), pSortierkriterium)

    Do
        If (pKnzAufsteigend) Then

            While (getStringSortierung(index_start, pSortierkriterium) < vergleichs_string_mitte)
                
                index_start = index_start + 1
            
            Wend

            While (getStringSortierung(index_ende, pSortierkriterium) > vergleichs_string_mitte)
                
                index_ende = index_ende - 1
            
            Wend

        Else

            While (getStringSortierung(index_start, pSortierkriterium) > vergleichs_string_mitte)
                
                index_start = index_start + 1
            
            Wend

            While (getStringSortierung(index_ende, pSortierkriterium) < vergleichs_string_mitte)
                
                index_ende = index_ende - 1
            
            Wend
            
        End If

        If (index_start <= index_ende) Then
            
            temp_string = m_array_strings(index_start)
            
            m_array_strings(index_start) = m_array_strings(index_ende)
            
            m_array_strings(index_ende) = temp_string

            index_start = index_start + 1
            
            index_ende = index_ende - 1

        End If

    Loop Until (index_start > index_ende)

    If (pIndexStart < index_ende) Then
    
        Call QuickSort(pIndexStart, index_ende, pSortierkriterium, pKnzAufsteigend)
        
    End If
    
    If (index_start < pIndexEnde) Then
    
        Call QuickSort(index_start, pIndexEnde, pSortierkriterium, pKnzAufsteigend)
        
    End If

End Sub

'################################################################################
'
Public Sub startSortierung(pSortierkriterium As Long, pKnzAufsteigend As Boolean, pKnzMarkIstDatum As Boolean, pKnzBenutzeMarkierung As Boolean, pAbPosition As Long, pBisPosition As Long)

On Error GoTo errStartSortierung

    m_knz_umstellung_datum_machen = pKnzMarkIstDatum
    
    m_knz_benutze_markierung = pKnzBenutzeMarkierung
    
    m_ab_position = pAbPosition
    
    m_bis_position = pBisPosition

    Call QuickSort(1, m_anzahl_daten_strings, pSortierkriterium, pKnzAufsteigend)

EndFunktion:

    Exit Sub

errStartSortierung:

    Debug.Print Err, Error$

    Resume EndFunktion

End Sub

'################################################################################
'
' Vertauscht die gespeicherten Strings zufaellig.
'
Public Sub startZufallsUmsortierung()

Dim temp_string   As String ' temporaerer String fuer die Vertauschung
Dim index_aktuell As Long   ' der Index der While-Schleife
Dim index_zufall  As Long   ' der zufaellige Index, fuer die neue Position

On Error Resume Next

    index_aktuell = 0
    
    Randomize

    '
    ' Ueber eine While-Schleife werden alle gesetzten Vektor-Strings einmal vertauscht
    '
    While (index_aktuell < m_anzahl_daten_strings)
    
        '
        ' Der Index der While-Schleife wird erhoeht
        '
        index_aktuell = index_aktuell + 1
           
        '
        ' Der zweite Index wird per Zufall bestimmt.
        ' Die Anzahl der gespeicherten Strings bestimmt dabei
        ' die Obergrenze des Zufallswertes.
        '
        index_zufall = Int(1 + m_anzahl_daten_strings * Rnd)
        
        '
        ' Pruefungen:
        ' Die beiden Indexwerte muessen sich unterscheiden.
        ' Der Zufallsindex darf nicht groesser als die anzahl der gespeicherten Strings sein.
        '
        ' Sind die beiden Bedingungen erfuellt, werden die Array-Inhalte an den Positionen getauscht.
        '
        If ((index_zufall <> index_aktuell) And (index_zufall <= m_anzahl_daten_strings)) Then
            
            temp_string = m_array_strings(index_aktuell)
            
            m_array_strings(index_aktuell) = m_array_strings(index_zufall)
            
            m_array_strings(index_zufall) = temp_string

        End If
    
    Wend

End Sub
