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

Private m_array_strings()          As String
Private m_anzahl_daten_strings     As Long
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
    
End Sub

'################################################################################
'
Public Sub addString(ByVal pString As String)

    m_anzahl_daten_strings = m_anzahl_daten_strings + 1

    If UBound(m_array_strings, 1) < m_anzahl_daten_strings Then
        
        ReDim Preserve m_array_strings(m_anzahl_daten_strings + 3)
    
    End If
    
    m_array_strings(m_anzahl_daten_strings) = Replace(pString, Chr(9), "    ")

End Sub

'################################################################################
'
Public Function getAnzahlStrings() As Long

    getAnzahlStrings = m_anzahl_daten_strings
    
End Function

'################################################################################
'
Public Function getString(pIndex As Long) As String

    If ((pIndex > 0) And (pIndex <= m_anzahl_daten_strings)) Then
        
        getString = m_array_strings(pIndex)
    
    Else
    
        getString = ""
        
    End If
    
End Function

'################################################################################
'
Public Sub setString(pIndex As Long, pString As String)

    If ((pIndex > 0) And (pIndex <= m_anzahl_daten_strings)) Then
        
        m_array_strings(pIndex) = pString
    
    End If
    
End Sub

'################################################################################
'
Public Function toString(pZeichenZeilenumbruch As String, Optional pKnzLeerzeilenEntfernen As Boolean = False) As String

Dim ergebnis_str   As String
Dim zeilen_zaehler As Long

    ergebnis_str = ""
    
    zeilen_zaehler = 1
    
    While (zeilen_zaehler <= m_anzahl_daten_strings)
        
        If (pKnzLeerzeilenEntfernen) Then
            
            If (Trim(m_array_strings(zeilen_zaehler)) <> "") Then
                
                ergebnis_str = ergebnis_str & m_array_strings(zeilen_zaehler) & pZeichenZeilenumbruch
            
            End If
        
        Else
            
            ergebnis_str = ergebnis_str & m_array_strings(zeilen_zaehler) & pZeichenZeilenumbruch
        
        End If
        
        zeilen_zaehler = zeilen_zaehler + 1
            
    Wend

    toString = ergebnis_str
    
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

    If (pIndexStart < index_ende) Then Call QuickSort(pIndexStart, index_ende, pSortierkriterium, pKnzAufsteigend)
    
    If (index_start < pIndexEnde) Then Call QuickSort(index_start, pIndexEnde, pSortierkriterium, pKnzAufsteigend)

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
Public Sub startZufallsUmsortierung()

Dim temp_string   As String
Dim index_aktuell As Long
Dim index_zufall  As Long

On Error Resume Next

    index_aktuell = 0
    
    Randomize

    While (index_aktuell < m_anzahl_daten_strings)
    
        index_aktuell = index_aktuell + 1
            
        index_zufall = Int(1 + m_anzahl_daten_strings * Rnd)
        
        If (index_zufall <> index_aktuell) And (index_zufall <= m_anzahl_daten_strings) Then
            
            temp_string = m_array_strings(index_aktuell)
            
            m_array_strings(index_aktuell) = m_array_strings(index_zufall)
            
            m_array_strings(index_zufall) = temp_string

        End If
    
    Wend

End Sub

'    Dim zeilen_anzahl  As Long
'    Dim zeilen_zaehler As Long
'    Dim aktuelle_zeile As String
'
'    '
'    ' Anzahl der insgesamt vorhandenen Zeilen aus dem String-Array ermitteln
'    '
'    zeilen_anzahl = pStringArray.getAnzahlStrings
'
'    '
'    ' Zeilenzaehler auf 1 stellen.
'    '
'    zeilen_zaehler = 1
'
'    '
'    ' Schleifendurchlauf von 1 bis zu der Anzahl der vorhandenen Zeilen.
'    ' Es ist kein Endlossschleifen-Verhinderungszaehler vorhanden.
'    '
'    While (zeilen_zaehler <= zeilen_anzahl)
'        '
'        ' Aktuelle Zeile
'        ' Die aktuelle Zeile wird per Index aus dem Zeilenobjekt gelesen
'        ' und in der Variablen "aktuelle_zeile" gespeichert.
'        '
'        aktuelle_zeile = cls_string_array.getString(zeilen_zaehler)
'
'
'        '
'        ' Zeilenzaehler erhoehen
'        ' Am Ende der IF-Kaskade wird der Zeilenzaehler fuer den naechsten Durchgang um 1 erhoeht.
'        '
'        zeilen_zaehler = zeilen_zaehler + 1
'
'    Wend

