VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private BOOLEAN_TRUE    As Integer
Private BOOLEAN_FALSE   As Integer
Private BOOLEAN_UNKNOWN As Integer

Private m_excel_objekt  As Object
Private m_max_zeile     As Integer
Private m_max_spalte    As Integer

'################################################################################
'
Private Sub wlExcel(pString As String)

    'print pString ' Lotus Notes
    'debug.Print pString

End Sub

'################################################################################
'
Sub Delete()

    If (m_excel_objekt Is Nothing) = False Then
        'm_excel_objekt.Quit
    End If

    Set m_excel_objekt = Nothing

End Sub

'################################################################################
'
Public Function istExcelExcelObjektGesetzt() As Boolean

On Error Resume Next

    istExcelExcelObjektGesetzt = (m_excel_objekt Is Nothing) = False

End Function

'################################################################################
'
Public Function initExcelObjekt() As Boolean

On Error GoTo errInitExcelObjekt

    BOOLEAN_TRUE = 1
    BOOLEAN_FALSE = 2
    BOOLEAN_UNKNOWN = 0

    If (m_excel_objekt Is Nothing) Then

        Set m_excel_objekt = CreateObject("Excel.application")

    End If

    If (m_excel_objekt Is Nothing) Then

        initExcelObjekt = False

    Else

        initExcelObjekt = True

        m_excel_objekt.Visible = True

    End If

    Exit Function

errInitExcelObjekt:

    'writeLogExcel("Fehler: errInitExcelObjekt: " & Err & " " & Error$ & " " & Erl)

    initExcelObjekt = False

    Exit Function

End Function

'################################################################################
'
Public Sub closeExcelObjekt()

On Error Resume Next

    m_excel_objekt.Visible = False
    m_excel_objekt.Quit

    Set m_excel_objekt = Nothing

End Sub

'################################################################################
'
Public Function openWorkbook(ByVal pDateiName As String, ByVal pKnzMessageboxAnzeigen As Boolean) As Boolean

On Error GoTo errOpenWorkbook

labelRetry:

    m_excel_objekt.Workbooks.Open FileName:=pDateiName

    'LNOTES m_excel_objekt.Workbooks.Open( pDateiName )

    openWorkbook = True

    Exit Function

errOpenWorkbook:

    'writeLogExcel("Fehler: errOpenWorkbook: " & Err & " " & Error$ & " " & Erl)

    If (pKnzMessageboxAnzeigen) Then
        '
        ' vbOKCancel =  1
        ' vbOK       =  1 = Rueckgabe
        '
        'If ( MsgBox( "Fehler: " & Err & " " & Error$ & " " & Erl, vbOKCancel, "Fehler openWorkbook" ) = vbOK ) Then
        If (MsgBox("Fehler: " & Err & " " & Error$ & " " & Erl, 1, "Fehler openWorkbook") = 1) Then     ' = das geht auch unter Lotus-Notes

            Resume labelRetry

        End If

    End If

    openWorkbook = False

    Exit Function

End Function

'################################################################################
'
Public Function closeWorkbook(ByVal pDateiName As String) As Boolean

On Error Resume Next

    Dim knz_save_display_alerts As Boolean

    knz_save_display_alerts = m_excel_objekt.DisplayAlerts

    m_excel_objekt.DisplayAlerts = False

    m_excel_objekt.Workbooks.Close FileName:=pDateiName

    m_excel_objekt.DisplayAlerts = knz_save_display_alerts

    closeWorkbook = True

End Function

'################################################################################
'
Public Function addWorkbook(ByVal pDateiName As String) As Boolean

On Error GoTo errAddWorkbook

    m_excel_objekt.Workbooks.Add

    'm_excel_objekt.ActiveWorkbook.Name = pDateiName

    addWorkbook = True

    Exit Function

errAddWorkbook:

    'writeLogExcel("Fehler: errAddWorkbook: " & Err & " " & Error$ & " " & Erl)

    addWorkbook = False

    Exit Function

End Function

'################################################################################
'
Public Function selectWorksheet(pWorkSheetName As String) As Boolean

On Error GoTo errSelectWorksheet

    Call m_excel_objekt.Worksheets(pWorkSheetName).Select

    selectWorksheet = True

    Exit Function

errSelectWorksheet:

    'writeLogExcel("Fehler: errSelectWorksheet: " & Err & " " & Error$ & " " & Erl)

    selectWorksheet = False

    Exit Function

End Function

'################################################################################
'
Public Function addSheet(ByVal pSheetName As String) As Boolean

    Dim new_excel_sheet As Object

    On Error GoTo errAddSheet:

    Dim knz_save_display_alerts As Boolean

    'set new_excel_sheet = m_excel_objekt.Sheets.add( after:=m_excel_objekt.Sheets( m_excel_objekt.Sheets.Count ) ) ' = Nur Visual-Basic

    Set new_excel_sheet = m_excel_objekt.Sheets().Add() ' = Lotus Notes

    On Error Resume Next

    new_excel_sheet.Name = pSheetName

    addSheet = True

    Exit Function

errAddSheet:

On Error Resume Next

    'writeLogExcel("Fehler: errAddSheet: " & Err & " " & Error$ & " " & Erl)

    If (Err.Number > 0) Then

        knz_save_display_alerts = m_excel_objekt.DisplayAlerts

        m_excel_objekt.DisplayAlerts = False

        new_excel_sheet.Delete

        m_excel_objekt.DisplayAlerts = knz_save_display_alerts

        addSheet = False

    Else

        addSheet = True

    End If

End Function

'################################################################################
'
Public Function selectSheet(ByVal pSheetName As String) As Boolean

On Error GoTo errSelectSheet

    m_excel_objekt.Sheets(pSheetName).Select

    selectSheet = True

    Exit Function

errSelectSheet:

    'writeLogExcel("Fehler: errSelectSheet1: " & Err & " " & Error$ & " " & Erl)

    selectSheet = False

    Exit Function

End Function

'################################################################################
'
Public Sub setExcelExcelObjektVisible(ByVal pVisible As Boolean)

    m_excel_objekt.Visible = pVisible

End Sub

'################################################################################
'
Public Sub setCursor(ByVal pZeile As Integer, ByVal pSpalte As Integer)

On Error Resume Next

    m_excel_objekt.Cells(pZeile, pSpalte).Select

End Sub

'################################################################################
'
Public Function setRangeByString(ByVal pStartZeile As Integer, ByVal pEingabe As String, ByVal pTrennzeichen As String, ByVal pZeilenTrennzeichen As String) As Integer

On Error GoTo errSetRangeByString

    setRangeByString = pStartZeile

    If (pEingabe = "") Then

        Exit Function

    End If

Dim akt_position          As Long
Dim akt_zeile             As String
Dim zeichen_zeilenumbruch As String
Dim ergebnis              As String
Dim letzte_position       As Long
Dim zeilen_zaehler        As Long
Dim knz_erste_zeile       As Boolean
Dim knz_umbruch_vorhanden As Boolean

    knz_erste_zeile = True

    ergebnis = ""
    '
    ' Ermittlung welches Zeilenumbruchzeichen verwendet in der Eingabe verwendet wird
    '
    If (pZeilenTrennzeichen <> LEER_STRING) Then

        zeichen_zeilenumbruch = pZeilenTrennzeichen

        knz_umbruch_vorhanden = (InStr(1, pEingabe, zeichen_zeilenumbruch) > 0)

    Else

        zeichen_zeilenumbruch = MY_CHR_13_10

        knz_umbruch_vorhanden = (InStr(1, pEingabe, zeichen_zeilenumbruch) > 0)

        If (knz_umbruch_vorhanden = False) Then

            zeichen_zeilenumbruch = Chr(13)

            knz_umbruch_vorhanden = (InStr(1, pEingabe, zeichen_zeilenumbruch) > 0)

        End If

    End If
    '
    ' Wenn Zeilenumbrueche vorhanden sind, wird eine Schleife gestartet.
    ' Sind keine Zeilenumbrueche vorhanden, gibt es nur einen Aufruf.
    ' Desweiteren wuerde sich die Schleifenkonstruktion verkomplizieren,
    ' wenn diese auch Zeichenketten ohne Zeilenumbruch verarbeiten sollte.
    '
    If (knz_umbruch_vorhanden) Then

        letzte_position = 1

        akt_position = InStr(letzte_position, pEingabe, zeichen_zeilenumbruch)

        While (akt_position > 0) And (zeilen_zaehler < 120)

            akt_zeile = Mid(pEingabe, letzte_position, akt_position - letzte_position)

            Call setZeileByString(pStartZeile, akt_zeile, pTrennzeichen)

            pStartZeile = pStartZeile + 1

            letzte_position = akt_position + Len(zeichen_zeilenumbruch)

            akt_position = InStr(letzte_position, pEingabe, zeichen_zeilenumbruch)

            zeilen_zaehler = zeilen_zaehler + 1

        Wend

        If (letzte_position < Len(pEingabe)) Then

            akt_zeile = Mid(pEingabe, letzte_position, (Len(pEingabe) - letzte_position) + 1)

            If (Trim(akt_zeile) <> LEER_STRING) Then

                ergebnis = ergebnis & akt_zeile & zeichen_zeilenumbruch

                Call setZeileByString(pStartZeile, akt_zeile, pTrennzeichen)

                pStartZeile = pStartZeile + 1

            End If

        End If

    Else

        Call setZeileByString(pStartZeile, pEingabe, pTrennzeichen)

        pStartZeile = pStartZeile + 1

    End If

    setRangeByString = pStartZeile

    Exit Function

errSetRangeByString:

    'writeLogExcel("Fehler: errSetRangeByString: " & Err & " " & Error$ & " " & Erl)

    setRangeByString = -1

    Exit Function

End Function

'################################################################################
'
Public Function setZeileByString(ByVal pZeile As Integer, ByVal pEingabe As String, ByVal pTrennzeichen As String) As Boolean

On Error GoTo errSetZeileByString

 'Call pExcelObjekt.setZeileByString(start_zeile + akt_zeile, "a,b,b,d,e,f,,e", ",")

Dim aktueller_string     As String
Dim index_end_position   As Integer
Dim index_start_position As Integer
Dim zaehler              As Integer

    zaehler = 0

    index_start_position = 1
    index_end_position = InStr(index_start_position, pEingabe, pTrennzeichen)

    While (index_start_position > 0) And (zaehler < 32123)

        '
        ' den Zaehler erhoehen
        '
        zaehler = zaehler + 1
        '
        ' Den aktuellen String aus der Eingabe ermitteln.
        '
        ' Wenn das Trennzeichen gefunden werden konnte, is index_end_position groesser als 0 und auch
        ' groesser als index_end_position. Es wird ein Teilstring ab index_start_position mit der Laenge index_end_position - index_start_position
        ' aus der Eingabe gelesen.
        '
        ' Wenn das Trennzeichen nicht ( mehr ) gefunden werden konnte, ist index_end_position gleich 0.
        ' Es wird ein Teilstring ab index_start_position bis zum Ende des Eingabestrings gelesen.
        '
        If (index_end_position > 0) Then

            aktueller_string = Mid(pEingabe, index_start_position, index_end_position - index_start_position)

            index_end_position = index_end_position + 1 ' neue Startposition liegt hinter dem Trennzeichen

        Else

            aktueller_string = Mid(pEingabe, index_start_position)

        End If
        '
        ' irgendwas mit dem String machen
        '
        'If IsNumeric(aktueller_string) Then

        '    m_excel_objekt.Cells(pZeile, zaehler).Value = CDbl(aktueller_string)

        'Else
          
          m_excel_objekt.Cells(pZeile, zaehler).NumberFormat = "@"
          
            m_excel_objekt.Cells(pZeile, zaehler).Value = aktueller_string

        'End If

        Call checkMaxZeileSpalte(pZeile, zaehler)
        '
        ' Trennzeichen ueberlesen
        '
        index_start_position = index_end_position
        
        '
        ' naechstes Trennzeichen in der Eingabe suchen
        '
        If (index_start_position > 0) Then

            index_end_position = InStr(index_start_position, pEingabe, pTrennzeichen)

        Else

            index_end_position = 0

        End If

    Wend

    setZeileByString = True

    Exit Function

errSetZeileByString:

    'writeLogExcel("Fehler: errSetZeileByString: " & Err & " " & Error$ & " " & Erl)

    setZeileByString = False

    Exit Function

End Function

'################################################################################
'
Public Function getMaxZeile() As Integer

    getMaxZeile = m_max_zeile

End Function

'################################################################################
'
Public Function getMaxSpalte() As Integer

    getMaxSpalte = m_max_spalte

End Function

'################################################################################
'
Public Sub resetMaxZeileSpalte()

    m_max_zeile = -1
    
    m_max_spalte = -1

End Sub

'################################################################################
'
Private Sub checkMaxZeileSpalte(ByVal pZeile As Integer, ByVal pSpalte As Integer)

     If (pZeile > m_max_zeile) Then
     
         m_max_zeile = pZeile
     
     End If

     If (pSpalte > m_max_spalte) Then
     
         m_max_spalte = pSpalte
     
     End If

End Sub

