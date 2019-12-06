Attribute VB_Name = "fkCsvExport2Excel"
Option Explicit

Private knz_export_laeuft As Boolean

'################################################################################
'
Public Function startCsv2Excel(pString As String, pTrennzeichen As String) As Boolean

On Error GoTo errStartCsv2Excel

Dim pExcelObjekt      As clsExcel  ' Instanz mit den Excel-Funktionen

Dim excel_sheet_name  As String    ' Name des zu erstellenden Tabellenblattes
Dim knz_fehlerfrei    As Boolean   ' Kennzeichen, ob die Erstellung fehlerfrei gemacht wurde

Dim cls_string_array As clsStringArray

    startCsv2Excel = False
    
    If (pTrennzeichen = "") Then
    
        Exit Function
        
    End If

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
        ' Initialisierung der internen Variablen
        '
        knz_fehlerfrei = True
    
        excel_sheet_name = "Export" & Format(Now, "dd.mm.yyyy")
    
        If ((knz_fehlerfrei) And (istAbbruchVerarbeitung() = False)) Then
    
            Set pExcelObjekt = New clsExcel
    
            knz_fehlerfrei = pExcelObjekt.initExcelObjekt()
    
        End If
    
        If ((knz_fehlerfrei) And (istAbbruchVerarbeitung() = False)) Then
            
            knz_fehlerfrei = pExcelObjekt.addWorkbook("ExcelExport")
    
        End If

        If ((knz_fehlerfrei) And (istAbbruchVerarbeitung() = False)) Then

            knz_fehlerfrei = pExcelObjekt.addSheet(excel_sheet_name)

        End If

        If ((knz_fehlerfrei) And (istAbbruchVerarbeitung() = False)) Then

            knz_fehlerfrei = pExcelObjekt.selectSheet(excel_sheet_name)
        
        End If

        If ((knz_fehlerfrei) And (istAbbruchVerarbeitung() = False)) Then

            Dim start_zeile As Long
            
            Dim zeilen_anzahl  As Long
            Dim zeilen_zaehler As Long
            Dim aktuelle_zeile As String
        
            start_zeile = 3
            
            '
            ' Anzahl der insgesamt vorhandenen Zeilen aus dem String-Array ermitteln
            '
            zeilen_anzahl = cls_string_array.getAnzahlStrings
        
            '
            ' Zeilenzaehler auf 1 stellen.
            '
            zeilen_zaehler = 1
        
            '
            ' Schleifendurchlauf von 1 bis zu der Anzahl der vorhandenen Zeilen.
            '
            While (zeilen_zaehler <= zeilen_anzahl)
                '
                ' Aktuelle Zeile
                ' Die aktuelle Zeile wird per Index aus dem Zeilenobjekt gelesen
                ' und in der Variablen "aktuelle_zeile" gespeichert.
                '
                aktuelle_zeile = cls_string_array.getString(zeilen_zaehler)
        
                If (pExcelObjekt.setZeileByString(start_zeile + zeilen_zaehler, aktuelle_zeile, pTrennzeichen) = False) Then
                
                    zeilen_zaehler = zeilen_anzahl + 1
                
                End If
        
                '
                ' Zeilenzaehler erhoehen
                ' Am Ende der IF-Kaskade wird der Zeilenzaehler fuer den naechsten Durchgang um 1 erhoeht.
                '
                zeilen_zaehler = zeilen_zaehler + 1
        
            Wend
    
        End If
    
    End If

EndFunktion:

    On Error Resume Next

    Set pExcelObjekt = Nothing
    Set cls_string_array = Nothing

    '
    ' Verarbeitungskennzeichen (Export OK oder nicht) dem Aufrufer zurueckgeben
    '
    startCsv2Excel = knz_fehlerfrei

    '
    ' DoEvents aufrufen
    '
    DoEvents

    '
    ' Funktion verlassen
    '
    Exit Function

errStartCsv2Excel:

    '##sss("Fehler: errStartCsv2Excel: " & Err & " " & Error & " " & Erl)

    knz_fehlerfrei = False

    Resume EndFunktion

End Function

'################################################################################
'
Private Function istAbbruchVerarbeitung() As Boolean

    istAbbruchVerarbeitung = False
    
End Function