VERSION 5.00
Begin VB.Form frmMrStringer 
   Caption         =   "vbMrStringer"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   27420
   LinkTopic       =   "Form1"
   ScaleHeight     =   726
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1828
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton m_startHtmlUrlDecoder 
      Caption         =   "URLD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   24750
      TabIndex        =   141
      ToolTipText     =   "Url-Decoded"
      Top             =   2625
      Width           =   1065
   End
   Begin VB.CommandButton m_btnFormatJavaLeerzeilen 
      Caption         =   "Fjava"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   24750
      TabIndex        =   140
      ToolTipText     =   "Format Java Leerzeilen"
      Top             =   2100
      Width           =   1065
   End
   Begin VB.CommandButton m_btnSetCsvZeichen 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   23700
      TabIndex        =   139
      ToolTipText     =   "Setzt das aktuelle CSV-Zeichen vorne oder hinten"
      Top             =   60
      Width           =   465
   End
   Begin VB.CommandButton m_btnStartMaskiereAnfuehrungszeichen 
      Caption         =   "MANF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   138
      ToolTipText     =   "Anfuehrungszeichen maskieren"
      Top             =   9975
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartJsp2Java 
      Caption         =   "JSPJV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   137
      ToolTipText     =   "JSP nach Java Funktion"
      Top             =   9450
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartJavaProperties 
      Caption         =   "JPROP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   136
      ToolTipText     =   "Java-Properties setzen. Trennung bei Markierung"
      Top             =   8925
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartHtmlTabelleVar 
      Caption         =   "TABL D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   135
      ToolTipText     =   "Debugausgabe als HTML-Tabelle "
      Top             =   6300
      Width           =   1140
   End
   Begin VB.CommandButton m_startHtmlUrlEncoded 
      Caption         =   "URLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   134
      ToolTipText     =   "Url-Encoding"
      Top             =   8400
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartHtmlTabelleCsv 
      Caption         =   "<TABL>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   133
      ToolTipText     =   "Erstellt eine HTML-Tabelle mit Selektion als Trennzeichen"
      Top             =   5775
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartHtmlJoinTabelle 
      Caption         =   "HTML J"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   132
      ToolTipText     =   "Join als HTML-Tabelle"
      Top             =   5250
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartBlockZufall 
      Caption         =   "BZUF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   131
      ToolTipText     =   "Zufallswerte auf Grundlage des Textes"
      Top             =   7875
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartGroup 
      Caption         =   "GRP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   130
      Top             =   7350
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartGetAscii 
      Caption         =   "ASCI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   129
      Top             =   6825
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartGetHexDump 
      Caption         =   "HEX D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   24750
      TabIndex        =   128
      ToolTipText     =   "Erstellt einen Hex-Dump"
      Top             =   1575
      Width           =   1065
   End
   Begin VB.CommandButton m_btnStartHtmlGeneratorLink 
      Caption         =   "HTML L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   127
      ToolTipText     =   "Erstellt HTML-Link Elemente"
      Top             =   4200
      Width           =   1140
   End
   Begin VB.CommandButton m_btnCsvExcel 
      Caption         =   "EXCL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   126
      ToolTipText     =   "Eingabe in ein Excel-Sheet schreiben"
      Top             =   3675
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartXmlNrJava 
      Caption         =   "XMLnr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   17775
      TabIndex        =   125
      Top             =   1095
      Width           =   945
   End
   Begin VB.CommandButton m_btnStartReplace4 
      Caption         =   "RPL4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   124
      ToolTipText     =   "http-Parameter auslesen und in bean setzen"
      Top             =   3150
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartGeneriereHtmlTabelle 
      Caption         =   "Gtab"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   123
      Top             =   2625
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartReplace3 
      Caption         =   "RPL3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   122
      Top             =   2100
      Width           =   1140
   End
   Begin VB.CommandButton m_btnZeilenBoolean 
      Caption         =   "ZB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   23325
      TabIndex        =   121
      Top             =   1095
      Width           =   615
   End
   Begin VB.CommandButton m_btnCsvReplaceMarkierung 
      Caption         =   "CSV R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   22650
      TabIndex        =   120
      ToolTipText     =   "CSV-Trennzeichen an Cursorpos setzen"
      Top             =   60
      Width           =   1035
   End
   Begin VB.CommandButton m_btnDupliziereMarkZeilen 
      Caption         =   "DM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   119
      Top             =   1095
      Width           =   585
   End
   Begin VB.CommandButton m_btnMarkiereWort 
      Caption         =   "MW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4125
      TabIndex        =   118
      Top             =   1095
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartReplace2 
      Caption         =   "RPL2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   117
      Top             =   1575
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartStrLitKonstanten 
      Caption         =   "StrLitKo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25050
      TabIndex        =   116
      Top             =   600
      Width           =   1140
   End
   Begin VB.CommandButton m_btnUmlaute 
      Caption         =   "&Umlaute"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   23400
      TabIndex        =   115
      ToolTipText     =   "Ersetzt deutsche Umlaute"
      Top             =   1575
      Width           =   1260
   End
   Begin VB.CommandButton m_btnSetCsvTrennzeichen 
      Caption         =   "CSV TZ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   21450
      TabIndex        =   114
      ToolTipText     =   "CSV-Trennzeichen an Cursorpos setzen"
      Top             =   60
      Width           =   1110
   End
   Begin VB.CommandButton m_btnErstelleKonstantenToProp 
      Caption         =   "KOP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   24150
      TabIndex        =   113
      ToolTipText     =   "Konstanten in Properties "
      Top             =   600
      Width           =   795
   End
   Begin VB.CommandButton m_btnStartLeerzeilenEinfuegen 
      Caption         =   "Lz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8250
      TabIndex        =   112
      ToolTipText     =   "Leerzeilen einfuegen"
      Top             =   1095
      Width           =   765
   End
   Begin VB.CommandButton m_btnStartHtmlQuotes 
      Caption         =   "&HTML Q"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25905
      TabIndex        =   111
      ToolTipText     =   "Html-Quotes"
      Top             =   4725
      Width           =   1140
   End
   Begin VB.CommandButton m_btnStartClrTxt 
      Caption         =   "Clr TXT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   14130
      TabIndex        =   110
      Top             =   1575
      Width           =   1080
   End
   Begin VB.CommandButton m_btnMarkiereDoppeltePlusMinus 
      Caption         =   "+ -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2190
      TabIndex        =   107
      Top             =   1095
      Width           =   555
   End
   Begin VB.CommandButton m_btnStartCheckLeerstring 
      Caption         =   "Chk LS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15255
      TabIndex        =   106
      Top             =   1575
      Width           =   1080
   End
   Begin VB.CommandButton m_btnStartNotesDebugFeldWerte 
      Caption         =   "D. Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   24990
      TabIndex        =   105
      ToolTipText     =   "Debug NotesFelder"
      Top             =   1095
      Width           =   1215
   End
   Begin VB.CommandButton m_btnStartFormatJson 
      Caption         =   "Frmt JSON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11430
      TabIndex        =   104
      Top             =   1575
      Width           =   1380
   End
   Begin VB.CommandButton m_btnMarkiereDoppeltePlus 
      Caption         =   "MD +"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1395
      TabIndex        =   103
      Top             =   1095
      Width           =   735
   End
   Begin VB.CommandButton m_btnStartCalcExe 
      Caption         =   "Calc.exe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   16380
      TabIndex        =   102
      ToolTipText     =   "Startet Calc.exe"
      Top             =   1575
      Width           =   1230
   End
   Begin VB.CommandButton m_btnStartCmdExe 
      Caption         =   "CMD.exe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   17655
      TabIndex        =   101
      Top             =   1575
      Width           =   1230
   End
   Begin VB.CommandButton m_btnStrgVIbmLog 
      Caption         =   "Ibm Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   18930
      TabIndex        =   100
      Top             =   1575
      Width           =   1155
   End
   Begin VB.CommandButton m_btnStartFormatTxt 
      Caption         =   "Frmt TXT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12855
      TabIndex        =   99
      Top             =   1575
      Width           =   1230
   End
   Begin VB.CommandButton m_btnCamelCase 
      Caption         =   "CamelCase"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   21675
      TabIndex        =   98
      Top             =   1575
      Width           =   1605
   End
   Begin VB.CommandButton m_btnZeilenAdd 
      Caption         =   "Zeilen Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   20130
      TabIndex        =   97
      Top             =   1575
      Width           =   1455
   End
   Begin VB.CommandButton m_btnGrepZahl 
      Caption         =   "GZ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3675
      TabIndex        =   96
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton m_btnTestDivers 
      Caption         =   "Gsin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   26250
      TabIndex        =   95
      Top             =   600
      Width           =   795
   End
   Begin VB.CommandButton m_btnStartGetterSetterJavaScript 
      Caption         =   "GJS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9105
      TabIndex        =   94
      Top             =   1095
      Width           =   795
   End
   Begin VB.CommandButton m_btnStartGrepMarkPlus 
      Caption         =   "M +"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   75
      TabIndex        =   93
      Top             =   1095
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartGrepMarkMinus 
      Caption         =   "M -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   735
      TabIndex        =   92
      Top             =   1095
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartTestUmdrehen 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   26865
      TabIndex        =   91
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton m_btnStartTrimLeerzeilen 
      Caption         =   "Trim L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7290
      TabIndex        =   90
      Top             =   1095
      Width           =   915
   End
   Begin VB.CommandButton m_btnMakeLongDatum 
      Caption         =   "Long"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   22455
      TabIndex        =   89
      ToolTipText     =   "Long Datum erstellen"
      Top             =   600
      Width           =   870
   End
   Begin VB.CommandButton m_btnStartGetterSetterVb 
      Caption         =   "GSv"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9945
      TabIndex        =   88
      Top             =   1095
      Width           =   675
   End
   Begin VB.CommandButton m_btnStartGetterSetter 
      Caption         =   "GS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10665
      TabIndex        =   87
      Top             =   1095
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartSetNull 
      Caption         =   "null"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11340
      TabIndex        =   86
      Top             =   1095
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartSumme 
      Caption         =   "SUM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12030
      TabIndex        =   85
      Top             =   1095
      Width           =   915
   End
   Begin VB.CommandButton m_btnStartSortDatum 
      Caption         =   "Sort D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9195
      TabIndex        =   84
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton m_btnCopyEingabe 
      Caption         =   "Copy &Eing."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7695
      TabIndex        =   83
      Top             =   1575
      Width           =   1755
   End
   Begin VB.CommandButton m_btnStartDoppelteVorkommen 
      Caption         =   "V D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11775
      TabIndex        =   82
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartEinmaligeVorkommen 
      Caption         =   "V 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11115
      TabIndex        =   81
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartRot13 
      Caption         =   "ROT13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   18810
      TabIndex        =   80
      Top             =   1095
      Width           =   1080
   End
   Begin VB.CommandButton m_strBlock 
      Caption         =   "BLK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13095
      TabIndex        =   79
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton m_btnStartMove 
      Caption         =   "MV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3435
      TabIndex        =   78
      Top             =   1095
      Width           =   615
   End
   Begin VB.CommandButton m_startCsvCase 
      Caption         =   "CSV Case"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   19995
      TabIndex        =   77
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton m_btnStartGrepSuchworteNegativ 
      Caption         =   "G2 -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2175
      TabIndex        =   76
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton m_btnSwitchPfad 
      Caption         =   "Pfad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   26250
      TabIndex        =   75
      Top             =   1095
      Width           =   795
   End
   Begin VB.CommandButton m_btnStartVbToJava 
      Caption         =   "VB->Java"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   19980
      TabIndex        =   74
      Top             =   1095
      Width           =   1230
   End
   Begin VB.CommandButton m_btnStartSortZufall 
      Caption         =   "Sort Z"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8235
      TabIndex        =   73
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton m_btnStartPlaceX 
      Caption         =   "Plce. X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8430
      TabIndex        =   72
      ToolTipText     =   "Text 2 an Cursor"
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton m_btnStartReplaceX 
      Caption         =   "Repl. X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9525
      TabIndex        =   71
      ToolTipText     =   "Replace Suchworte"
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton m_btnTrimX 
      Caption         =   "Trim X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6285
      TabIndex        =   70
      Top             =   1095
      Width           =   915
   End
   Begin VB.CommandButton m_btnSetGatter0Ende 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3135
      TabIndex        =   69
      ToolTipText     =   "Setzt ein Suchbegriff am Wortende"
      Top             =   60
      Width           =   375
   End
   Begin VB.CommandButton m_btnSetGatter0Zurueck 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2055
      TabIndex        =   68
      ToolTipText     =   "Setzt am Wortanfang den Suchbegriff"
      Top             =   60
      Width           =   375
   End
   Begin VB.CommandButton m_btnStartChr13Konvertierung 
      Caption         =   "CHR13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   20775
      TabIndex        =   67
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton m_btnStartSortierungLaenge 
      Caption         =   "Sort L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7275
      TabIndex        =   66
      ToolTipText     =   "Sortieren nach Laenge"
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton m_btnStartNotesLesenSchreiben 
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   24030
      TabIndex        =   65
      Top             =   1095
      Width           =   915
   End
   Begin VB.CommandButton m_btnJoinX 
      Caption         =   "Join X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7470
      TabIndex        =   64
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton m_btnStartJSON 
      Caption         =   "JSON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12915
      TabIndex        =   63
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton m_btnStartDirEinlesen 
      Caption         =   "Dir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   23400
      TabIndex        =   62
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartReverse 
      Caption         =   "RVSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5700
      TabIndex        =   61
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton m_btnIf2 
      Caption         =   "IF2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   16935
      TabIndex        =   60
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartAusrichter1 
      Caption         =   "Ausrichter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4395
      TabIndex        =   59
      ToolTipText     =   "Ausrichtung mit der Markierung"
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton m_btnStartUnique 
      Caption         =   "unique"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10155
      TabIndex        =   58
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton m_btnCsvDoppelpunkt 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25785
      TabIndex        =   57
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton m_btnCsvPunkt 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   26145
      TabIndex        =   56
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton m_btnStartNamen 
      Caption         =   "Namen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   18555
      TabIndex        =   55
      Top             =   600
      Width           =   1035
   End
   Begin VB.CommandButton m_btnStartZaehler 
      Caption         =   "NR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   22620
      TabIndex        =   54
      Top             =   1095
      Width           =   615
   End
   Begin VB.CommandButton m_btnGrepWort 
      Caption         =   "GW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2955
      TabIndex        =   53
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartSplit 
      Caption         =   "Split"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3615
      TabIndex        =   52
      ToolTipText     =   "Teilt den Text am Cursor oder an der Markierung"
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton m_btnStartToZeile 
      Caption         =   "toZeile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   19635
      TabIndex        =   51
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton m_btnStartDebugAusgabe 
      Caption         =   "Debug"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10590
      TabIndex        =   50
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton m_btnDeklaration 
      Caption         =   "DIM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15195
      TabIndex        =   49
      Top             =   600
      Width           =   675
   End
   Begin VB.CommandButton m_btnSetGatter0 
      Caption         =   "#0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2475
      TabIndex        =   48
      ToolTipText     =   "Setzt ein Suchwort an der Cursorposition"
      Top             =   75
      Width           =   615
   End
   Begin VB.CommandButton m_btnDuplizierung 
      Caption         =   "D1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2805
      TabIndex        =   47
      Top             =   1095
      Width           =   585
   End
   Begin VB.CommandButton m_btnErstelleKonstantenEinfach 
      Caption         =   "&KO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   21795
      TabIndex        =   46
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartUCaseLCase 
      Caption         =   "UL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12435
      TabIndex        =   45
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton m_btnCsvGleichKomma 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25425
      TabIndex        =   44
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton m_txtCsvPipe 
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   26505
      TabIndex        =   43
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton m_btnCsvSemikolon 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   25065
      TabIndex        =   42
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton m_btnCsvGleich 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   24705
      TabIndex        =   41
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton m_btnStartTrim 
      Caption         =   "Trim"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5505
      TabIndex        =   40
      Top             =   1095
      Width           =   735
   End
   Begin VB.CommandButton m_btnStartSpalte1 
      Caption         =   "#1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   75
      TabIndex        =   39
      ToolTipText     =   "Setzt vorne oder hinten einen Suchbegriff"
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartCsvToZeile 
      Caption         =   "CSV Zeile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   18615
      TabIndex        =   38
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton m_btnStartCsvKonstanten 
      Caption         =   "CSV CONST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   16935
      TabIndex        =   37
      Top             =   60
      Width           =   1635
   End
   Begin VB.CommandButton m_btnStartCsvSwap 
      Caption         =   "CSV SWAP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15255
      TabIndex        =   36
      Top             =   60
      Width           =   1635
   End
   Begin VB.CommandButton m_btnSwitchEingabe 
      Caption         =   "Eing. tauschen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4575
      TabIndex        =   35
      Top             =   1575
      Width           =   1995
   End
   Begin VB.HScrollBar scrollTeiler 
      Height          =   315
      Left            =   2760
      Max             =   90
      Min             =   10
      TabIndex        =   3
      Top             =   6120
      Value           =   50
      Width           =   6510
   End
   Begin VB.CommandButton m_btnStartGrepSuchworteP 
      Caption         =   "G2 +"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1395
      TabIndex        =   34
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton m_cmdToggleEingabe 
      Caption         =   "Eing."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6675
      TabIndex        =   33
      Top             =   1575
      Width           =   915
   End
   Begin VB.TextBox m_txtEingabe2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   32
      Top             =   2460
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.CommandButton m_btnErstelleCsv 
      Caption         =   "Erst. CSV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13935
      TabIndex        =   31
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton m_btnGrepWeglassen 
      Caption         =   "G -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   735
      TabIndex        =   30
      ToolTipText     =   "Aufnahme aller Zeilen ohne dem markiertem Wort"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton m_btnGrepAufnehmen 
      Caption         =   "G +"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   75
      TabIndex        =   29
      ToolTipText     =   "Aufnahme aller Zeilen mit dem markiertem Wort"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton m_btnStartJoin 
      Caption         =   "Join"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6690
      TabIndex        =   28
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton m_btnStartXmlJavaWriter 
      Caption         =   "XMLwr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   16650
      TabIndex        =   27
      Top             =   1095
      Width           =   1020
   End
   Begin VB.TextBox m_txtCsvZeichen 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   24225
      TabIndex        =   26
      Text            =   ","
      Top             =   60
      Width           =   435
   End
   Begin VB.CommandButton m_startGetStringLit 
      Caption         =   "StrLit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   17595
      TabIndex        =   25
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton m_btnStartRemove 
      Caption         =   "RMVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5355
      TabIndex        =   24
      ToolTipText     =   "Entfernt das markierte Wort"
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton m_btnStartClip 
      Caption         =   "Clip"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4395
      TabIndex        =   23
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton m_btnStartSortierung 
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6315
      TabIndex        =   22
      ToolTipText     =   "Sortierung auf- oder absteigend"
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton m_btnCopyToEingabe 
      Caption         =   "Strg + &V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   21
      Top             =   1575
      Width           =   1755
   End
   Begin VB.CommandButton m_btnStartFallunterscheidungVB 
      Caption         =   "IF-VB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15915
      TabIndex        =   20
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton m_btnErstelleXmlFormat2 
      Caption         =   "XML 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   14010
      TabIndex        =   19
      Top             =   1095
      Width           =   975
   End
   Begin VB.CommandButton m_btnErstelleXmlFormat 
      Caption         =   "XML 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12990
      TabIndex        =   18
      Top             =   1095
      Width           =   975
   End
   Begin VB.CommandButton m_btnStringVb 
      Caption         =   "ToString "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11730
      TabIndex        =   17
      Top             =   600
      Width           =   1155
   End
   Begin VB.CommandButton m_btnCopyAusgabe2Eiingabe 
      Caption         =   "Ausgabe als Eingabe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1860
      TabIndex        =   16
      Top             =   1575
      Width           =   2595
   End
   Begin VB.CommandButton m_btnStartFormatXml 
      Caption         =   "Format XML"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15075
      TabIndex        =   15
      Top             =   1095
      Width           =   1515
   End
   Begin VB.CommandButton m_btnStartCmdRename 
      Caption         =   "Rename"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   21300
      TabIndex        =   14
      Top             =   1095
      Width           =   1230
   End
   Begin VB.CommandButton m_btnGeneratorJava 
      Caption         =   "&Gen. Java"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13875
      TabIndex        =   13
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton m_btnCopy 
      Caption         =   "Copy &Ausg."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9555
      TabIndex        =   12
      Top             =   1575
      Width           =   1755
   End
   Begin VB.CommandButton m_btnStartSpalte2 
      Caption         =   "#2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   735
      TabIndex        =   11
      ToolTipText     =   "Setzt vorne und hinten einen Suchbegriff"
      Top             =   60
      Width           =   615
   End
   Begin VB.Frame m_frameEinstellungen 
      BorderStyle     =   0  'Kein
      Caption         =   "m_lblStringBufferStart"
      Height          =   3735
      Left            =   10500
      TabIndex        =   4
      Top             =   4425
      Visible         =   0   'False
      Width           =   11715
      Begin VB.TextBox m_txtTrennzeichen4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1830
         TabIndex        =   108
         Text            =   "#0"
         Top             =   1575
         Width           =   1455
      End
      Begin VB.TextBox m_txtTrennzeichen3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Text            =   "#3"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox m_txtTrennzeichen2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Text            =   "#2"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox m_txtTrennzeichen1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Text            =   "#1"
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Trennzeichen 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   109
         Top             =   1575
         Width           =   1695
      End
      Begin VB.Label m_lblTrennzeichen3 
         Caption         =   "Trennzeichen 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label m_lblTrennzeichen2 
         Caption         =   "Trennzeichen 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label m_lblTrennzeichen1 
         Caption         =   "Trennzeichen 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox m_txtAusgabe 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   6180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   2460
      Width           =   2175
   End
   Begin VB.TextBox m_txtEingabe 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   660
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   2460
      Width           =   2715
   End
   Begin VB.CommandButton m_btnStartSpalte3 
      Caption         =   "#3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1395
      TabIndex        =   0
      ToolTipText     =   "Dupliziert die Eingabe und setzt Suchbegriffe"
      Top             =   60
      Width           =   615
   End
   Begin VB.Line m_lineResize 
      Visible         =   0   'False
      X1              =   42
      X2              =   1040
      Y1              =   144
      Y2              =   577
   End
End
Attribute VB_Name = "frmMrStringer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private knz_togle_form_gen      As Boolean

Private knz_generator_vb        As Boolean
Private knz_pfad                As Boolean
Private knz_eingabe_volle_hoehe As Boolean
Private knz_resize_laeuft       As Boolean
Private m_knz_tz_anfang         As Boolean
Private m_zaehler_chr13         As Integer

Private Sub m_btnFormatJavaLeerzeilen_Click()


Dim ergebnis_fkt   As String
Dim my_cr          As String
Dim text_clip      As String
Dim replace_text_1 As String
Dim replace_text_2 As String

    my_cr = vbCrLf
    
    knz_togle_form_gen = Not knz_togle_form_gen

    replace_text_1 = "FkString.getFeldLinksMin("
    
    replace_text_2 = ", breite_temp" & IIf(knz_togle_form_gen, "_02", "_01") & " )"

    '
    ' Anweisungen fuer den Ausrichter erstellen
    '
    
    ergebnis_fkt = startMrStringer(FKT_ENTFERNE_LEERZEILEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
    
    ergebnis_fkt = startMrStringer(FKT_LEERZEILEN_EINFUEGEN, ergebnis_fkt, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

    ergebnis_fkt = Replace(ergebnis_fkt, vbCrLf, TRENN_STRING_8)
    ergebnis_fkt = Replace(ergebnis_fkt, vbCr, TRENN_STRING_8)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_8, vbCrLf)
    
    ergebnis_fkt = Replace(ergebnis_fkt, "{", vbCrLf & "{")
    
    ergebnis_fkt = Replace(ergebnis_fkt, "{" & vbCrLf, "{")
    
    Dim tab_zaehler As Integer

    While (tab_zaehler < 50)
    
        ergebnis_fkt = Replace(ergebnis_fkt, vbCrLf & text_clip & "{", text_clip & "{")
        
        ergebnis_fkt = Replace(ergebnis_fkt, vbCrLf & text_clip & vbCrLf & "{", vbCrLf & text_clip & "{")
    
        ergebnis_fkt = Replace(ergebnis_fkt, vbCrLf & text_clip & "}", text_clip & "}")
        
        text_clip = text_clip & " "
        
        tab_zaehler = tab_zaehler + 1
    Wend
    
    m_txtAusgabe.Text = ergebnis_fkt

End Sub

Private Sub m_btnSetCsvZeichen_Click()

'    m_txtAusgabe.Text = startMrStringer(FKT_MARKIERE_VORNE_ODER_HINTEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
    m_txtAusgabe.Text = startMrStringer(FKT_MARKIERE_CSV_VORNE_ODER_HINTEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False, m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartJsp2Java_Click()

    m_txtAusgabe.Text = startJsp2Java(m_txtEingabe.Text)

End Sub

'################################################################################
'
Private Sub Form_Load()

m_lineResize.X1 = 4

    scrollTeiler.Top = m_lineResize.Y1

    scrollTeiler.Left = m_lineResize.X1

    m_txtAusgabe.FontSize = 11

    m_txtEingabe.FontSize = m_txtAusgabe.FontSize

    m_txtEingabe2.FontSize = m_txtAusgabe.FontSize

    m_txtEingabe.Top = m_lineResize.Y1 + scrollTeiler.Height + 5

    m_txtEingabe2.Top = m_lineResize.Y1 + scrollTeiler.Height + 5

    m_txtAusgabe.Top = m_txtEingabe.Top

    m_txtEingabe.Left = m_lineResize.X1

    m_txtEingabe2.Left = m_lineResize.X1

    m_frameEinstellungen.Top = m_txtEingabe.Top

    m_frameEinstellungen.Top = m_lineResize.X1

    knz_resize_laeuft = False

    knz_eingabe_volle_hoehe = True

    m_txtEingabe2.Visible = Not knz_eingabe_volle_hoehe

    m_knz_aktiv = False
    
End Sub

'################################################################################
'
Private Sub Form_Resize()

    Dim breite_fenster_gesamt As Double
    Dim breite_scroll_prozent As Double
    
    If (knz_resize_laeuft) Then
        
        Exit Sub
    
    End If
    
    knz_resize_laeuft = True
    
    If (Me.ScaleWidth > m_lineResize.X2) Then
    
        breite_fenster_gesamt = CInt(Me.ScaleWidth - (m_lineResize.X1 * 3))
    
    Else
        
        breite_fenster_gesamt = CInt(m_lineResize.X2)
    
    End If
        
    breite_scroll_prozent = CInt((breite_fenster_gesamt * CInt(scrollTeiler.Value)) * 0.01)

    m_frameEinstellungen.Width = breite_fenster_gesamt
    
    scrollTeiler.Width = breite_fenster_gesamt
    
    m_txtEingabe.Width = breite_scroll_prozent
    
    m_txtEingabe2.Width = breite_scroll_prozent
    
    m_txtAusgabe.Width = breite_fenster_gesamt - breite_scroll_prozent
    
    m_txtAusgabe.Left = breite_scroll_prozent + m_lineResize.X1 * 2
    
    If (Me.ScaleHeight > m_lineResize.Y2) Then
    
        m_txtAusgabe.Height = Me.ScaleHeight - (m_lineResize.Y1 + m_lineResize.X1 + scrollTeiler.Height + 5) ' x1 = Abstand zu unteren Rand
        
        If (knz_eingabe_volle_hoehe) Then
            
            m_txtEingabe.Height = m_txtAusgabe.Height
            
            m_txtEingabe2.Top = m_txtEingabe.Top
            
            m_txtEingabe2.Height = m_txtAusgabe.Height
        
        Else
        
            m_txtEingabe.Height = CInt(m_txtAusgabe.Height * 0.5) - 10
            
            m_txtEingabe2.Top = m_txtEingabe.Top + m_txtEingabe.Height + 10
            
            m_txtEingabe2.Height = m_txtEingabe.Height + 10

        End If
        
    End If

    knz_resize_laeuft = False

End Sub

Private Sub m_btnStartMaskiereAnfuehrungszeichen_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_MASKIERE_ANFZEICHEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False, m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartTestUmdrehen_Click()
    
Dim vb_str As String

    vb_str = vb_str & vbCrLf & """CMD.exe""    = Startet eine Dos-Box"
    vb_str = vb_str & vbCrLf & """Calc.exe""   = Startet den Windows-Taschenrechner"
    vb_str = vb_str & vbCrLf & """#1""         = Suchzeichen vorne oder hinten"
    vb_str = vb_str & vbCrLf & """#2""         = Suchzeichen vorne und hinten"
    vb_str = vb_str & vbCrLf & """#3""         = Stringzeile gedoppelt mit Suchzeichen"
    vb_str = vb_str & vbCrLf & """#0""         = Suchzeichen an Cursorposition"
    vb_str = vb_str & vbCrLf & """>""          = Suchzeichen an Wortanfang ab Cursorposition"
    vb_str = vb_str & vbCrLf & """<""          = Suchzeichen an Wortende ab Cursorposition"
    vb_str = vb_str & vbCrLf & """G +""        = Grep+ Zeilen mit markiertem Wort finden"
    vb_str = vb_str & vbCrLf & """G -""        = Grep- Zeilen ohne markiertes Wort finden"
    vb_str = vb_str & vbCrLf & """M +""        = Mark+ Suchzeichen vorne oder hinten, bei Zeilen mit markiertem Wort"
    vb_str = vb_str & vbCrLf & """M -""        = Mark- Suchzeichen vorne oder hinten, bei Zeilen ohne markiertem Wort"
    vb_str = vb_str & vbCrLf & """+ -""        = Erst Grep+ dann Grep-"
    vb_str = vb_str & vbCrLf & """DM""         = Dupliziert die Zeilen, in welchen das markierte Wort vorkommt"
    vb_str = vb_str & vbCrLf & """Ausrichter"" = Ausrichtung mit Markierung"
    vb_str = vb_str & vbCrLf & """Split""      = Zeilen an Cursorposition oder Markierung teilen"
    vb_str = vb_str & vbCrLf & """Clip""       = Markierung heraustrennen"
    vb_str = vb_str & vbCrLf & """Sort""       = Zeilensortierung (insgesamt oder nach Markierung)"
    vb_str = vb_str & vbCrLf & """Sort L""     = Zeilensortierung nach Zeilenlänge"
    vb_str = vb_str & vbCrLf & """Sort D""     = Zeilensortierung nach Datum (DOS-Dateiauflistung)"
    vb_str = vb_str & vbCrLf & """Sort Z""     = Zufallsumstellung der Zeilen"
    vb_str = vb_str & vbCrLf & """RMVE""       = Markierung wird gelöscht (Remove)"
    vb_str = vb_str & vbCrLf & """Trim""       = Entfernung von führenden und abschliessenden Leerzeichen"
    vb_str = vb_str & vbCrLf & """Trim X""     = Entfernt auch doppelte Leerzeichen zwischen den Wörtern"
    vb_str = vb_str & vbCrLf & """Trim L""     = Entfernung von Leerzeilen "
    vb_str = vb_str & vbCrLf & """Lz""         = Entfernt Leerzeilen"
    vb_str = vb_str & vbCrLf & """UL""         = Uppercase / Lowercase (Zeile oder nach Markierung)"
    vb_str = vb_str & vbCrLf & """unique""     = Löscht doppelte Zeilen"
    vb_str = vb_str & vbCrLf & """Dir""        = Verzeichniseinlesen (Mit und Ohne Pfad)"
    vb_str = vb_str & vbCrLf & """NR""         = Zeilen zählen"
    vb_str = vb_str & vbCrLf & """MV""         = Verschiebt den Markierten Text nach vorne oder hinten"
    vb_str = vb_str & vbCrLf & """MW""         = Markiere das ausgewaehlte Wort"
    vb_str = vb_str & vbCrLf & """G2 +""       = Multigrep + mit der Eingabebox 2"
    vb_str = vb_str & vbCrLf & """G2 -""       = Multigrep - mit der Eingabebox 2"
    vb_str = vb_str & vbCrLf & """GW""         = Liefert alle Wörter welche mit der Makierung anfangen"
    vb_str = vb_str & vbCrLf & "                 Ist keine Makierung vorhanden, werden alle Wörter aufgelistet"
    vb_str = vb_str & vbCrLf & """GZ""         = Grep Zahlen"
    vb_str = vb_str & vbCrLf & """,""          = Trennzeichen Komma setzen"
    vb_str = vb_str & vbCrLf & """.""          = Trennzeichen Punkt setzen"
    vb_str = vb_str & vbCrLf & """:""          = Trennzeichen Doppelpunkt setzen"
    vb_str = vb_str & vbCrLf & """;""          = Trennzeichen Semikolon setzen"
    vb_str = vb_str & vbCrLf & """=""          = Trennzeichen Gleichheitszeichen setzen"
    vb_str = vb_str & vbCrLf & """KO""         = Konstanten in Java / VB "
    vb_str = vb_str & vbCrLf & """KOP""        = Konstanten in Properties (Java)"
    vb_str = vb_str & vbCrLf & """StrLitKo""   = Stringliterale im Text in Konstanten umwandeln"
    vb_str = vb_str & vbCrLf & """StrLit""     = Stringliterale im Text raussuchen"
    vb_str = vb_str & vbCrLf & """CHR13""      = Ersetzt den Zeilenumbruch mit chr(13)"
    vb_str = vb_str & vbCrLf & """CSV CONST""  = Erstellt aus den CSV-Daten Konstanten"
    vb_str = vb_str & vbCrLf & """CSV Case""   = Erstellt aus den CSV-Daten eine Case-Anweisung"
    vb_str = vb_str & vbCrLf & """CSV SWAP""   = Vertauscht die CSV-Daten"
    vb_str = vb_str & vbCrLf & """Notes""      = Generator fuer Lotus-Notes set und Get-Anweisungen"
    vb_str = vb_str & vbCrLf & """Join""       = Fügt die beiden Textinhalte A und B zusammen"
    vb_str = vb_str & vbCrLf & """Plce. X""    = Join mit einer Platzierung des zweiten Textes an der Cursorposition"
    vb_str = vb_str & vbCrLf & """Frmt TXT""   = Formatiert TXT auf eine Breite von 80 Zeichen"
    vb_str = vb_str & vbCrLf & """Pfad""       = Dreht Slasches in Pfadangaben um"
    vb_str = vb_str & vbCrLf & """Rename""     = Erstellt DOS-Rename-Anweisungen"
    vb_str = vb_str & vbCrLf & """Debug""      = Generator fuer Debug-Ausgaben"
    vb_str = vb_str & vbCrLf & """RVSE""       = Umdrehen des Textes"
    vb_str = vb_str & vbCrLf & """XML 2""      = XML-Darstellung 2"
    vb_str = vb_str & vbCrLf & """XML 1""      = XML-Darstellung 1"
    vb_str = vb_str & vbCrLf & """XMLnr""      = XML Schreiber/Parser mit Nummern"
    vb_str = vb_str & vbCrLf & """ROT13""      = Rot13-Algorithmus"
    vb_str = vb_str & vbCrLf & """Frmt JSON""  = Formatiert JSON"
    vb_str = vb_str & vbCrLf & """Format XML"" = Formatiert XML"
    vb_str = vb_str & vbCrLf & """Gen. Java""  = Kapselt jede Zeile in ein ""pBuffer.append()"" ein (Konvertierung in Java-String) "
    vb_str = vb_str & vbCrLf & """Erst. CSV""  = Erstellt eine Zeile aus den Eingabedaten (Verwendung in Funktionsaufrufen)"
    vb_str = vb_str & vbCrLf & """CamelCase""  = Erstellt aus jeder Zeile einen CamelCase-String"
    vb_str = vb_str & vbCrLf & """JSON""       = Eingabe als JSON-String"
    vb_str = vb_str & vbCrLf & """CSV Zeile""  ="
    vb_str = vb_str & vbCrLf & """VB->Java""   = Sehr grobe Konvertierung von Visual-Basic nach Java-Quelltext"
    vb_str = vb_str & vbCrLf & """ToString""   = jede Zeile einem String hinzufügen"
    vb_str = vb_str & vbCrLf & """D. Notes""   = NotesDokument set und get Feldwert"
    vb_str = vb_str & vbCrLf & """Umlaute""    = Konvertiert Umlaute von z.B. Ä nach Ae"
    vb_str = vb_str & vbCrLf & """Repl. X""    ="
    vb_str = vb_str & vbCrLf & """Zeilen Add"" ="
    vb_str = vb_str & vbCrLf & """EXCL""       = Eingabe als CSV-Export nach Excel "
    vb_str = vb_str & vbCrLf & """BLK""        = Blockgenerierung"
    vb_str = vb_str & vbCrLf & """Ibm Log""    = Konvertierung Zeilenumbruch IBM-Log"
    vb_str = vb_str & vbCrLf & """XMLwr""      = Generator XML-Writer"
    vb_str = vb_str & vbCrLf & """IF-VB""      = Generator IF fuer Visual-Basic"
    vb_str = vb_str & vbCrLf & """IF2""        = Genarator IF (Version 2)"
    vb_str = vb_str & vbCrLf & """MD +""       ="
    vb_str = vb_str & vbCrLf & """VD""         = Vorkommen von doppelten Strings suchen"
    vb_str = vb_str & vbCrLf & """V1""         = Vorkommen von einmal vorkommenden Strings suchen "
    vb_str = vb_str & vbCrLf & """HTML Q""     = setzt Html-Qoute-Zeichen"
    vb_str = vb_str & vbCrLf & """HTML L""     = HTML-Link: Erstellt aus Eingabe 1 jeweils einen HTML-Link"
    vb_str = vb_str & vbCrLf & """HTML J""     = HTML-Join: Erstellt aus Eingabe 1 und Eingabe 2 eine HTML-Tabelle"
    vb_str = vb_str & vbCrLf & """TABL D""     = HTML-Table-Debugausgabe: Variablen toString als HTML-Tabelle"
    vb_str = vb_str & vbCrLf & """HEX D""      = erstellt einen Hex-Dump"
    vb_str = vb_str & vbCrLf & """D1""         = Dupliziert die Zeile oder Markierung"
    vb_str = vb_str & vbCrLf & """ZB""         = Zeilen Boolean / Markierung 0 und 1 abwechselnd je Zeile"
    vb_str = vb_str & vbCrLf & """GS""         = Generator Set/Get Java"
    vb_str = vb_str & vbCrLf & """GJS""        = Generator Set/Get Java-Script"
    vb_str = vb_str & vbCrLf & """GSv""        = Generator Set/Get Visual-Basic"
    vb_str = vb_str & vbCrLf & """DIM""        = Variablendeklaration in Visual-Basic"
    vb_str = vb_str & vbCrLf & """SUM""        = grobe Aufsummierung von Werten"
    vb_str = vb_str & vbCrLf & """GRP""        = setzt eine Leerzeile, wenn sich der Text der Markierung aendert"
    vb_str = vb_str & vbCrLf & """BZUF""       = Zufallsgenerator auf Grundlage des aktuellen Zeichens"
    vb_str = vb_str & vbCrLf & """JPROP""      = Java-Properties setzen. Trennung bei Markierung"
    vb_str = vb_str & vbCrLf & """MANF""       = Anfuehrungszeichen maskieren Java und VB"
    vb_str = vb_str & vbCrLf & ""
    vb_str = vb_str & vbCrLf & """?""          = Zeigt den Hilfetext an"
    vb_str = vb_str & vbCrLf & ""
    vb_str = vb_str & vbCrLf & """Strg + V""            = Setzt den Text aus der Zwischenablage in das Eingabefeld"
    vb_str = vb_str & vbCrLf & """Eing.""               = Zeigt die zweite Eingabebox an"
    vb_str = vb_str & vbCrLf & """Eing. tauschen""      = vertauscht die Inhalte der beiden Eingabefelder"
    vb_str = vb_str & vbCrLf & """Ausgabe als Eingabe"" = setzt die Ausgabe als Eingabe"
    vb_str = vb_str & vbCrLf & """Copy Eing.""          = kopiert den Inhalt der ersten Eingabebox in die Zwischenablage"
    vb_str = vb_str & vbCrLf & """Copy Ausg.""          = kopiert den Inhalt der Ausgabebox in die Zwischenablage"
    vb_str = vb_str & vbCrLf & ""
    vb_str = vb_str & vbCrLf & ""
    vb_str = vb_str & vbCrLf & "Beispieltext"
    vb_str = vb_str & vbCrLf & ""
    vb_str = vb_str & vbCrLf & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB"
    vb_str = vb_str & vbCrLf & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB"
    vb_str = vb_str & vbCrLf & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB"
    vb_str = vb_str & vbCrLf & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB"
    vb_str = vb_str & vbCrLf & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB"
    vb_str = vb_str & vbCrLf & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB"
    vb_str = vb_str & vbCrLf & "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB"
    vb_str = vb_str & vbCrLf & "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB"
    vb_str = vb_str & vbCrLf & "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB"
    vb_str = vb_str & vbCrLf & "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB"
    vb_str = vb_str & vbCrLf & "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB"
    vb_str = vb_str & vbCrLf & "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaB"
    vb_str = vb_str & vbCrLf & "00000000000000000000000000000000000000000000000B"
    vb_str = vb_str & vbCrLf & "00000000000000000000000000000000000000000000000B"
    vb_str = vb_str & vbCrLf & "00000000000000000000000000000000000000000000000B"
    vb_str = vb_str & vbCrLf & "00000000000000000000000000000000000000000000000B"
    vb_str = vb_str & vbCrLf & "00000AAA.BBB.CCC0000000000000000000000000000000B"
    vb_str = vb_str & vbCrLf & "00000AAA.BBB.CCC0000000000000000000000000000000B"

    m_txtAusgabe.Text = vb_str

End Sub

'################################################################################
'
Private Sub m_btnCopyEingabe_Click()

On Error GoTo errCopy

    Clipboard.Clear
    
    Clipboard.SetText m_txtEingabe.Text
    
    Exit Sub

errCopy:
    
    MsgBox Error
    
    Exit Sub

End Sub

'################################################################################
'
Private Sub m_btnCopy_Click()

On Error GoTo errCopy

    Clipboard.Clear
    
    Clipboard.SetText m_txtAusgabe.Text
    
    Exit Sub

errCopy:
    
    MsgBox Error
    
    Exit Sub
    
End Sub

'################################################################################
'
Private Sub m_btnCopyAusgabe2Eiingabe_Click()

    If (m_txtAusgabe.SelStart > 0) Then
    
        m_txtEingabe.Text = Mid(m_txtAusgabe.Text, m_txtAusgabe.SelStart, m_txtAusgabe.SelLength)
    
    Else
        
        m_txtEingabe.Text = m_txtAusgabe.Text
    
    End If

End Sub

'################################################################################
'
Private Sub m_btnStartHtmlJoinTabelle_Click()

Dim ergebnis_fkt   As String
Dim my_cr          As String
Dim replace_text_1 As String
Dim replace_text_2 As String
Dim replace_text_3 As String

    my_cr = vbCrLf

    replace_text_1 = "<tr><td>"
    replace_text_2 = "</td><td>"
    replace_text_3 = "</td></tr>"
    
    '
    ' Anweisungen fuer den Ausrichter erstellen
    '
    ergebnis_fkt = startJoin(m_txtEingabe.Text, m_txtEingabe2.Text, TRENN_STRING_6)

    ergebnis_fkt = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN, ergebnis_fkt, -1, -1, False, "")

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_7, replace_text_1)

    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , TRENN_STRING_6)

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6, replace_text_2)

    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , TRENN_STRING_8)

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_8, replace_text_3)

    m_txtAusgabe.Text = "<table>" & my_cr & ergebnis_fkt & my_cr & "</table>"

End Sub

'################################################################################
'
Private Sub m_btnStartHtmlTabelleVar_Click()

Dim ergebnis_fkt   As String
Dim my_cr          As String
Dim replace_text_1 As String
Dim replace_text_2 As String
Dim replace_text_3 As String
Dim str_vorlauf    As String
Dim str_nachlauf   As String

    knz_togle_form_gen = Not knz_togle_form_gen

    my_cr = vbCrLf

    If (knz_togle_form_gen) Then
    
        replace_text_1 = "str_html_tabelle += my_cr + ""<tr><td>"" + """
        
        replace_text_2 = " " & AUSRICHT_STRING_TEMP_1 & """ + ""</td><td>"" + "
        
        replace_text_3 = " " & AUSRICHT_STRING_TEMP_2 & "+ ""</td></tr>"";"
    
        str_vorlauf = "String str_html_tabelle = """";" & my_cr & "String my_cr            = ""\n"";" & my_cr & my_cr & "str_html_tabelle += ""<table>"";" & my_cr & my_cr
                      
        str_nachlauf = my_cr & my_cr & "str_html_tabelle += ""</table>""; "
        
        str_nachlauf = str_nachlauf & my_cr & my_cr & "//" & replace_text_1 & replace_text_2 & """&nbsp;""" & replace_text_3
        
    
    Else
    
        str_vorlauf = "Dim str_html_tabelle As String" & my_cr & "Dim my_cr            As String" & my_cr & my_cr & "str_html_tabelle = ""<table>"" " & my_cr & my_cr
        
        str_nachlauf = my_cr & my_cr & "str_html_tabelle = str_html_tabelle & my_cr & ""</table>"" "
    
        replace_text_1 = "str_html_tabelle = str_html_tabelle & my_cr & ""<tr><td>"" & """
        
        replace_text_2 = " " & AUSRICHT_STRING_TEMP_1 & """ & ""</td><td>"" & "
        
        replace_text_3 = " " & AUSRICHT_STRING_TEMP_2 & "& ""</td></tr>"""
    
        str_nachlauf = str_nachlauf & my_cr & my_cr & "'" & replace_text_1 & replace_text_2 & """&nbsp;""" & replace_text_3
    
    End If
        
    ergebnis_fkt = m_txtEingabe.Text
    
    ergebnis_fkt = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE, ergebnis_fkt, -1, -1)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6 & TRENN_STRING_7 & TRENN_STRING_8, "")

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6, replace_text_1)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_7, replace_text_2)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_8, replace_text_3)
    
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , AUSRICHT_STRING_TEMP_1)
    
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , AUSRICHT_STRING_TEMP_2)
    
    m_txtAusgabe.Text = Replace(Replace(str_vorlauf & ergebnis_fkt & str_nachlauf, AUSRICHT_STRING_TEMP_1, ""), AUSRICHT_STRING_TEMP_2, "")

End Sub

'################################################################################
'
Private Sub m_btnStartJavaProperties_Click()

    checkCsvSelektion

Dim ergebnis_fkt   As String
Dim my_cr          As String
Dim replace_text_1 As String
Dim replace_text_2 As String
Dim replace_text_3 As String
Dim str_vorlauf    As String
Dim str_nachlauf   As String

    my_cr = vbCrLf
    
    str_vorlauf = "Properties " & STR_VAR_NAME_PROPERTIES_LOKAL & " = new Properties();" & my_cr & my_cr
    
    ergebnis_fkt = startMrStringer(FKT_CSV_JAVA_PROP, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False, m_txtCsvZeichen.Text)
    
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , AUSRICHT_STRING_TEMP_1)
    
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , AUSRICHT_STRING_TEMP_2)
    
    m_txtAusgabe.Text = Replace(Replace(str_vorlauf & ergebnis_fkt, AUSRICHT_STRING_TEMP_1, ""), AUSRICHT_STRING_TEMP_2, "")

End Sub

'################################################################################
'
Private Sub m_btnStartReplace2_Click()

Dim ergebnis_fkt   As String
Dim my_cr          As String
Dim text_clip      As String
Dim replace_text_1 As String
Dim replace_text_2 As String

    my_cr = vbCrLf
    
    knz_togle_form_gen = Not knz_togle_form_gen

    replace_text_1 = "FkString.getFeldLinksMin("
    
    replace_text_2 = ", breite_temp" & IIf(knz_togle_form_gen, "_02", "_01") & " )"

    '
    ' Anweisungen fuer den Ausrichter erstellen
    '
    ergebnis_fkt = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_7, replace_text_1)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_8, replace_text_2)

    ergebnis_fkt = Replace(ergebnis_fkt, replace_text_1 + replace_text_2, "")

    '
    ' Anweisungen um die Breite zu ermitteln
    '
    '
    ' 1. Markierungsspalte aus dem Eingabetext raustrennen
    '
    text_clip = startMrStringer(FKT_CLIP_GET_TEXT, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
    '
    ' 2. Entferne doppelte Strings aus dem Clip-String
    '
    text_clip = startMrStringer(FKT_GET_UNIQUE, text_clip, -1, -1)
    
    '
    ' 3. Eine Markierung vorne und hinten anfuegen
    '
    text_clip = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN, text_clip, -1, -1)

    '
    ' 4. Entferne die Stellen, an welchen der vordere und
    '    der hintere Trennstring aneinanderkommen.
    '    (Eleminierung von Leerzeilen aus Schritt 1)
    '
    text_clip = Replace(text_clip, TRENN_STRING_7 + TRENN_STRING_8, "")
    
    text_clip = startMrStringer(FKT_ENTFERNE_LEERZEILEN, text_clip, -1, -1) + TRENN_STRING_8
    
    '
    ' Entferne alle Stellen, an welchen sich der Trennstring 3
    ' zweimal nacheinander kommt.
    text_clip = Replace(text_clip, TRENN_STRING_8 + TRENN_STRING_8, "")

    '
    ' Entferne Trennstring 2, da dieser nicht mehr benoetigt wird
    '
    text_clip = Replace(text_clip, TRENN_STRING_7, "")
    
    '
    ' Ersetze Trennstring 3 mit einem Komma (Parametertrennung)
    '
    text_clip = Replace(text_clip, TRENN_STRING_8, ",")
    
    '
    ' Ergebnis
    '
    ergebnis_fkt = " breite_temp" & IIf(knz_togle_form_gen, "_02", "_01") & " = FkString.getMaxLen( " & text_clip & " ); " & my_cr & my_cr & ergebnis_fkt
    
    m_txtAusgabe.Text = ergebnis_fkt

End Sub

'################################################################################
'
Private Sub m_btnStartHtmlTabelleCsv_Click()

Dim ergebnis_fkt   As String
Dim my_cr          As String
Dim replace_text_1 As String
Dim replace_text_2 As String
Dim replace_text_3 As String
Dim str_markierung As String

    my_cr = vbCrLf

    replace_text_1 = "<tr><td>"
    
    replace_text_2 = "</td><td>"
    
    replace_text_3 = "</td></tr>"
    
    If ((m_txtEingabe.SelStart > 0) And ((m_txtEingabe.SelLength > 0) And (m_txtEingabe.SelLength <= 5))) Then
        
        str_markierung = Mid(m_txtEingabe.Text, m_txtEingabe.SelStart + 1, m_txtEingabe.SelLength)
        
        ergebnis_fkt = m_txtEingabe.Text
        
        ergebnis_fkt = Replace(ergebnis_fkt, str_markierung, replace_text_2)
    
        ergebnis_fkt = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN, ergebnis_fkt, -1, -1)
    
        ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_7, replace_text_1)
    
        ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_8, replace_text_3)
        
    Else
    
        ergebnis_fkt = "Keine Markierung vorhanden"
    
    End If
    
    m_txtAusgabe.Text = "<table>" & my_cr & ergebnis_fkt & my_cr & "</table>"

End Sub

'################################################################################
'
Private Sub m_btnStartHtmlGeneratorLink_Click()

Dim ergebnis_fkt   As String
Dim my_cr          As String
Dim replace_text_1 As String
Dim replace_text_2 As String
Dim replace_text_3 As String

    my_cr = vbCrLf

    replace_text_1 = "<A href="""
    
    replace_text_2 = """ " & AUSRICHT_STRING_TEMP_1 & "target=""_blank"">"
    
    replace_text_3 = "</A><BR>"
    
    '
    ' Anweisungen fuer den Ausrichter erstellen
    '
    ergebnis_fkt = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6 + TRENN_STRING_7 + TRENN_STRING_8, "")

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6, replace_text_1)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_7, replace_text_2)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_8, replace_text_3)

    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , AUSRICHT_STRING_TEMP_1)

    m_txtAusgabe.Text = Replace(ergebnis_fkt, AUSRICHT_STRING_TEMP_1, "")

End Sub

'################################################################################
'
Private Sub m_btnStartGeneriereHtmlTabelle_Click()

Dim ergebnis_fkt    As String
Dim my_cr           As String
Dim replace_text_1  As String
Dim replace_text_2  As String
Dim replace_text_3  As String
Dim str_table_start As String
Dim str_table_end   As String

    my_cr = vbCrLf
    
    knz_togle_form_gen = Not knz_togle_form_gen
    
    If (knz_togle_form_gen) Then
    
        str_table_start = "Dim html_table As String" & my_cr & my_cr & "html_table = html_table & ""<table>""" & my_cr & my_cr
        
        str_table_end = my_cr & "html_table = html_table & ""</table>"""
    
        replace_text_1 = "html_table = html_table & ""<tr><td>"" & """
        
        replace_text_2 = "" & TRENN_STRING_6 & """ & ""</td><td>"" & "
        
        replace_text_3 = "" & TRENN_STRING_7 & " & ""</td></tr>"""

    Else

        str_table_start = "String html_table = """";" & my_cr & my_cr & "html_table += ""<table>"";" & my_cr & my_cr
        
        str_table_end = my_cr & "html_table += ""</table>"";"

        replace_text_1 = "html_table += ""<tr><td>"" + """
        
        replace_text_2 = "" & TRENN_STRING_6 & """ + ""</td><td>"" + "
        
        replace_text_3 = "" & TRENN_STRING_7 & " + ""</td></tr>"";"
    
    End If

    '
    ' Trim auf die Eingabe ausfuehren
    '
    ergebnis_fkt = startMrStringer(FKT_TRIM_X, m_txtEingabe.Text, -1, -1)

    '
    ' Spalten verdoppeln und markieren
    '
    ergebnis_fkt = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE, ergebnis_fkt, -1, -1)

    '
    ' Leerzeilen korrigieren (Leerzeile bleibt Leerzeile)
    '
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6 & TRENN_STRING_7 & TRENN_STRING_8, "")

    '
    ' Ersetzungen fuer die Erstellung der HTML-Tabelle machen
    '
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6, replace_text_1)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_7, replace_text_2)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_8, replace_text_3)
    
    '
    ' Ausrichtung der Spalten
    '
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , TRENN_STRING_6)
    
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , TRENN_STRING_7)
    
    '
    ' Hilfsmarkierungen des Ausrichters entfernen
    '
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6, "")
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_7, "")

    '
    ' Ergebnis ins Ausgabefeld setzen
    '
    m_txtAusgabe.Text = str_table_start & ergebnis_fkt & str_table_end

End Sub

'################################################################################
'
Private Sub m_btnStartReplace3_Click()

Dim ergebnis_fkt   As String
Dim replace_text_1 As String
Dim replace_text_2 As String
Dim replace_text_3 As String

    replace_text_1 = "prop_instanz.setProperty( ModulKonfiguration."
    replace_text_2 = ", """
    replace_text_3 = """ );"

    '
    ' Anweisungen fuer den Ausrichter erstellen
    '
    ergebnis_fkt = startMrStringer(FKT_TRIM_X, m_txtEingabe.Text, -1, -1)

    ergebnis_fkt = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE, ergebnis_fkt, -1, -1)

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6 & TRENN_STRING_7 & TRENN_STRING_8, "")

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6, replace_text_1)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_7, replace_text_2)
    
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_8, replace_text_3)
    
    m_txtAusgabe.Text = ergebnis_fkt

End Sub

'################################################################################
'
Private Sub m_btnStartReplace4_Click()

Dim ergebnis_str_1 As String
Dim ergebnis_fkt   As String
Dim replace_text_1 As String
Dim replace_text_2 As String
Dim replace_text_3 As String

    '
    '#####################################################################################
    '
    
    ergebnis_str_1 = startMrStringer(FKT_TRIM_X, m_txtEingabe.Text, -1, -1)

    ergebnis_str_1 = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE, ergebnis_str_1, -1, -1)

    ergebnis_str_1 = Replace(ergebnis_str_1, TRENN_STRING_6 & TRENN_STRING_7 & TRENN_STRING_8, "")
    
    '
    '#####################################################################################
    '
    
    replace_text_1 = "String "
    replace_text_2 = " = FkHttpServlet.getParameter( pRequest, """
    replace_text_3 = """, null, 15 );"
     
    ergebnis_fkt = ergebnis_str_1

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6, replace_text_1)
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_7, replace_text_2)
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_8, replace_text_3)
    
    '
    '#####################################################################################
    '
    
    replace_text_1 = "if ( "
    replace_text_2 = " != null ) { anw_instanz.set( "
    replace_text_3 = " TRENN_STRING_9); }"
    
    ergebnis_fkt = ergebnis_fkt & vbCrLf & ergebnis_str_1

    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_6, replace_text_1)
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_7, replace_text_2)
    ergebnis_fkt = Replace(ergebnis_fkt, TRENN_STRING_8, replace_text_3)
    
    '
    '#####################################################################################
    '
    
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , " = FkHttpServlet.getPa")
    
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , ", null, 15 );")
    
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , "!= null ) { anw_inst")
    
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , "TRENN_STRING_9); }")
    
    ergebnis_fkt = Replace(ergebnis_fkt, "TRENN_STRING_9); }", "); }")
    
    m_txtAusgabe.Text = ergebnis_fkt

End Sub

'################################################################################
'
Private Sub m_btnErstelleKonstantenEinfach_Click()

    Call fkAppMrStringer.setToggleMrStringerFkt(Not fkAppMrStringer.getToggleMrStringerFkt())

Dim ergebnis_fkt As String
    
    ergebnis_fkt = startErstelleKonstantenEinfach(m_txtEingabe.Text, m_txtTrennzeichen1.Text, m_txtTrennzeichen2.Text, m_txtTrennzeichen3.Text, IIf(fkAppMrStringer.getToggleMrStringerFkt(), 1, 2), m_txtEingabe.SelStart, m_txtEingabe.SelLength)

    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , AUSRICHT_STRING_TEMP_1)

    m_txtAusgabe.Text = startMrStringer(FKT_AUSRICHTER_STRING, Replace(ergebnis_fkt, AUSRICHT_STRING_TEMP_1, ""), -1, -1, , " );")

End Sub

'################################################################################
'
Private Sub m_btnErstelleKonstantenToProp_Click()

Dim ergebnis_fkt As String
    
    ergebnis_fkt = startErstelleKonstantenEinfach(m_txtEingabe.Text, m_txtTrennzeichen1.Text, m_txtTrennzeichen2.Text, m_txtTrennzeichen3.Text, 3, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , AUSRICHT_STRING_TEMP_1)

    m_txtAusgabe.Text = startMrStringer(FKT_AUSRICHTER_STRING, Replace(ergebnis_fkt, AUSRICHT_STRING_TEMP_1, ""), -1, -1, , " );")

End Sub

'################################################################################
'
Private Sub m_btnStartGetAscii_Click()

    m_txtAusgabe.Text = startGetAsciiPrint(m_txtEingabe.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartGetHexDump_Click()

    m_txtAusgabe.Text = startGetHexDump(m_txtEingabe.Text, 14)

End Sub

'################################################################################
'
Private Sub m_btnStartGroup_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_GROUP_NACH_STRING, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartStrLitKonstanten_Click()

Dim cls_string_array      As clsStringArray

Dim akt_konst_name        As String
Dim aktuelle_zeile        As String
Dim ergebnis_fkt          As String
Dim my_cr                 As String
Dim str_lit               As String
Dim text_eingabe          As String
Dim zeichen_zeilenumbruch As String
Dim zeilen_anzahl         As Long
Dim zeilen_zaehler        As Long

    text_eingabe = m_txtEingabe.Text
    
    If (Len(text_eingabe) > 0) Then
        '
        ' 1. Stringliterale finden
        '
        str_lit = startMrStringer(FKT_STRING_LIT, text_eingabe, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
        
        '
        ' 2. Doppelte Strings entfernen
        '
        str_lit = startMrStringer(FKT_GET_UNIQUE, str_lit, -1, 0)
    
        my_cr = vbCrLf
        
        '
        ' 3. Die Strings in einem Stringarray verpacken
        '
        Set cls_string_array = startMultiline(str_lit)
        
        If (cls_string_array Is Nothing) Then
        
            ergebnis_fkt = "Fehler bei Erstellung ClsStringArray"
        
        Else
    
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
            '
            While (zeilen_zaehler <= zeilen_anzahl)
                
                '
                ' Aktuelle Zeile aus dem Zeilenobjekt lesen
                '
                aktuelle_zeile = cls_string_array.getString(zeilen_zaehler)
                
                '
                ' Laengenpruefung machen
                ' Die aktuelle Zeile muss mindestens 2 Stellen haben.
                ' Die aktuelle Zeile darf nicht mehr als 100 Stellen haben.
                '
                If ((Len(aktuelle_zeile) >= 2) And (Len(aktuelle_zeile) <= 100)) Then
                
                    '
                    ' Konstantennamen aus der aktuellen Zeile erstellen
                    '
                    akt_konst_name = "STR_KO_" & UCase(getKlartext(aktuelle_zeile, "_"))
                
                    '
                    ' Konstantendeklaration dem Ergebnisstring hinzufuegen
                    '
                    ergebnis_fkt = ergebnis_fkt & my_cr & "private static final String " & akt_konst_name & " = """ & aktuelle_zeile & """;"
                
                    '
                    ' Im Ausgangstext alle Vorkommen des Strings mit dem
                    ' Konstantennamen ersetzen.
                    '
                    text_eingabe = Replace(text_eingabe, """" & aktuelle_zeile & """", akt_konst_name)
                
                End If
                
                '
                ' Zeilenzaehler erhoehen
                '
                zeilen_zaehler = zeilen_zaehler + 1
                
            Wend
            
            '
            ' Dem bisherigen Ergebnisstring, wird der Ausgangsstring mit den
            ' gemachten Ersetzungen hinzhugefuegt.
            '
            ergebnis_fkt = ergebnis_fkt & my_cr & my_cr & text_eingabe
            
            '
            ' Instanz von Stringarray auf "nothing" setzen.
            '
            Set cls_string_array = Nothing
    
        End If
    
    End If
    
    m_txtAusgabe.Text = ergebnis_fkt

End Sub

'################################################################################
'
Private Sub m_btnStartFormatJson_Click()

Dim ergebnis_fkt As String
Dim my_cr As String

    my_cr = vbCrLf
    
    ergebnis_fkt = AUSRICHT_STRING_TEMP_1 & m_txtEingabe.Text
    ergebnis_fkt = Replace(ergebnis_fkt, "[", my_cr & "[" & my_cr)
    ergebnis_fkt = Replace(ergebnis_fkt, "]", my_cr & "]" & my_cr)
    ergebnis_fkt = Replace(ergebnis_fkt, "{", my_cr & "{" & my_cr)
    ergebnis_fkt = Replace(ergebnis_fkt, "}", my_cr & "}" & my_cr)
    ergebnis_fkt = Replace(ergebnis_fkt, ",", "," & my_cr)
    ergebnis_fkt = Replace(ergebnis_fkt, "]" & my_cr & ",", "],")
    ergebnis_fkt = Replace(ergebnis_fkt, "}" & my_cr & ",", "},")
    ergebnis_fkt = Replace(ergebnis_fkt, my_cr & my_cr, my_cr)
    ergebnis_fkt = Replace(ergebnis_fkt, my_cr & my_cr, my_cr)
    ergebnis_fkt = Replace(ergebnis_fkt, AUSRICHT_STRING_TEMP_1 & my_cr, "")
    ergebnis_fkt = Replace(ergebnis_fkt, AUSRICHT_STRING_TEMP_1, "")
 
    m_txtAusgabe.Text = ergebnis_fkt

End Sub

'################################################################################
'
Private Sub m_btnUmlaute_Click()
    
    Dim str_text As String
    
    str_text = m_txtEingabe.Text
    
    str_text = Replace(str_text, "ä", "ae")
    str_text = Replace(str_text, "Ä", "Ae")
    str_text = Replace(str_text, "ü", "ue")
    str_text = Replace(str_text, "Ü", "Ue")
    str_text = Replace(str_text, "ö", "oe")
    str_text = Replace(str_text, "Ö", "Oe")
    str_text = Replace(str_text, "ß", "ss")
    str_text = Replace(str_text, "€", "EUR")
    
    m_txtAusgabe.Text = str_text

End Sub

'################################################################################
'
Private Sub m_btnStartHtmlQuotes_Click()

    m_txtAusgabe.Text = quoteHtmlCharacter(m_txtEingabe.Text)

End Sub

'################################################################################
'
Private Sub cpy_IBM_Click()

    m_txtEingabe.Text = Replace(Clipboard.GetText, Chr(9), Chr(13) & Chr(10))
    
End Sub

'################################################################################
'
Private Sub m_btnCamelCase_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_CAMEL_CASE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnCsvReplaceMarkierung_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_CSV_REPLACE_MARKIERUNG_MIT_CSV, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, True, m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnDupliziereMarkZeilen_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_GREP_DUPLIZIERE_MARKZEILEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, True)

End Sub

'################################################################################
'
Private Sub m_btnGrepZahl_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_GREP_ZAHLEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnMakeLongDatum_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_MAKE_LONG_DATUM, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnMarkiereDoppeltePlus_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_MARKIERE_DOPPELTE_PLUS, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, True)

End Sub

'################################################################################
'
Private Sub m_btnMarkiereDoppeltePlusMinus_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_MARKIERE_DOPPELTE_PLUS_MINUS, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, True)

End Sub

'################################################################################
'
Private Sub m_btnMarkiereWort_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_MARKIERE_WORT, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, , m_txtCsvZeichen.Text)
    
End Sub

'################################################################################
'
Private Sub m_btnSetCsvTrennzeichen_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SET_TRENNZEICHEN_STR, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, , m_txtCsvZeichen.Text)
    
End Sub

'################################################################################
'
Private Sub m_btnStartCalcExe_Click()

    Shell "calc.exe", vbNormalFocus

End Sub

'################################################################################
'
Private Sub m_btnStartCheckLeerstring_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_CHECK_LEERSTRING, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartClrTxt_Click()

    m_txtAusgabe.Text = getStringGueltigeZeichen(m_txtEingabe.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartCmdExe_Click()

    Shell "cmd.exe", vbNormalFocus
    
End Sub

'################################################################################
'
Private Sub m_btnStartFormatTxt_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_FORMAT_TXT, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartGetterSetter_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_GETTER_SETTER_JAVA, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartGetterSetterJavaScript_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_GETTER_SETTER_JAVA_SCRIPT, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartGetterSetterVb_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_GETTER_SETTER_VB, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartGrepMarkMinus_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_GREP_MARK, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False)

End Sub

'################################################################################
'
Private Sub m_btnStartGrepMarkPlus_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_GREP_MARK, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, True)

End Sub

'################################################################################
'
Private Sub m_btnStartLeerzeilenEinfuegen_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_LEERZEILEN_EINFUEGEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartNotesDebugFeldWerte_Click()

   ' m_txtAusgabe.Text = startMrStringer(FKT_NOTES_LESEN_SCHREIBEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
    m_txtAusgabe.Text = startMrStringer(FKT_NOTES_DEBUG_FELD_WERTE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartSumme_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_CALC_SUMME, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartTrimLeerzeilen_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_ENTFERNE_LEERZEILEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartXmlNrJava_Click()

Dim ergebnis_fkt As String
Dim temp_toggle As Boolean

    ergebnis_fkt = startMrStringer(FKT_JAVA_XML_WRITER_NUMMER, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
    
    temp_toggle = getToggleMrStringerFkt()
    
    ergebnis_fkt = startMrStringer(FKT_AUSRICHTER_STRING, ergebnis_fkt, -1, -1, , AUSRICHT_STRING_TEMP_1)

    m_txtAusgabe.Text = Replace(ergebnis_fkt, AUSRICHT_STRING_TEMP_1, "")

    Call setToggleMrStringerFkt(temp_toggle)

End Sub

'################################################################################
'
Private Sub m_btnStrgVIbmLog_Click()
       
    m_txtEingabe.Text = Replace(Clipboard.GetText, Chr(10), Chr(13))

End Sub

'################################################################################
'
Private Sub m_btnSwitchPfad_Click()

    knz_pfad = Not knz_pfad

    If (knz_pfad) Then
        
        m_txtAusgabe.Text = Replace(m_txtEingabe.Text, "\", "/")
    
    Else
        
        m_txtAusgabe.Text = Replace(m_txtEingabe.Text, "/", "\")
    
    End If

End Sub

'################################################################################
'
Private Sub m_btnStartChr13Konvertierung_Click()

    m_zaehler_chr13 = m_zaehler_chr13 + 1
    
    If (m_zaehler_chr13 = 1) Then
        
        m_txtAusgabe.Text = Replace(m_txtEingabe.Text, Chr(13), Chr(13) & Chr(10))
    
    Else
        
        m_zaehler_chr13 = 0
        
        m_txtAusgabe.Text = Replace(m_txtEingabe.Text, Chr(10), Chr(13) & Chr(10))
    
    End If

End Sub

'################################################################################
'
Private Sub checkCsvSelektion()

    If (m_txtEingabe.SelLength > 0) Then
        
        m_txtCsvZeichen.Text = Mid(m_txtEingabe.Text, m_txtEingabe.SelStart + 1, m_txtEingabe.SelLength)
        
        m_txtEingabe.SelLength = 0
    
    End If

End Sub

'################################################################################
'
Private Sub m_btnJoinX_Click()

    m_txtAusgabe.Text = startJoin(m_txtEingabe.Text, m_txtEingabe2.Text, m_txtTrennzeichen3.Text, True)

End Sub

'################################################################################
'
Private Sub m_btnStartGrepSuchworteP_Click()

    m_txtAusgabe.Text = startGrepSuchWorte(m_txtEingabe.Text, m_txtEingabe2.Text, 1)

End Sub

'################################################################################
'
Private Sub m_btnStartGrepSuchworteNegativ_Click()
 
    m_txtAusgabe.Text = startGrepSuchWorte(m_txtEingabe.Text, m_txtEingabe2.Text, 0)

End Sub

'################################################################################
'
Private Sub m_btnSwitchEingabe_Click()

    Dim temp_str As String
    temp_str = m_txtEingabe.Text
    m_txtEingabe.Text = m_txtEingabe2.Text
    m_txtEingabe2.Text = temp_str

End Sub

'################################################################################
'
Private Sub m_btnTestDivers_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SINGLETON_JAVA, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
 
End Sub

'################################################################################
'
Private Sub m_btnZeilenAdd_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_ZEILEN_ADD, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnZeilenBoolean_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_ZEILEN_BOOLEAN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_cmdToggleEingabe_Click()

    knz_eingabe_volle_hoehe = Not knz_eingabe_volle_hoehe
    
    m_txtEingabe2.Visible = (knz_eingabe_volle_hoehe = False)
    
    Form_Resize

End Sub

'################################################################################
'
Private Sub m_btnCopyToEingabe_Click()

    m_txtEingabe.Text = Replace(Clipboard.GetText, Chr(9), "    ")
    
End Sub

'################################################################################
'
Private Sub m_btnStartCsvKonstanten_Click()

    checkCsvSelektion
    
    m_txtAusgabe.Text = startCsvKonstanten(m_txtEingabe.Text, m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartFormatXml_Click()

    If (m_txtEingabe.SelLength > 0) Then
    
        m_txtAusgabe.Text = formatXML(Mid(m_txtEingabe.Text, m_txtEingabe.SelStart + 1, m_txtEingabe.SelLength))
    
    Else
    
        m_txtAusgabe.Text = formatXML(m_txtEingabe.Text)
    
    End If

End Sub

'################################################################################
'
Private Sub m_btnStartJoin_Click()

    m_txtAusgabe.Text = startJoin(m_txtEingabe.Text, m_txtEingabe2.Text, m_txtTrennzeichen3.Text)

End Sub

Private Sub m_startHtmlUrlDecoder_Click()

    m_txtAusgabe.Text = getUrlDecoded(m_txtEingabe.Text)

End Sub

'################################################################################
'
Private Sub m_startHtmlUrlEncoded_Click()

    knz_togle_form_gen = Not knz_togle_form_gen

    m_txtAusgabe.Text = getUrlEncoded(m_txtEingabe.Text, knz_togle_form_gen)

End Sub

'################################################################################
'
Private Sub scrollTeiler_Change()

    Form_Resize

End Sub

'################################################################################
'
Private Sub scrollTeiler_Scroll()

    Form_Resize

End Sub

'################################################################################
'
Private Sub m_txtCsvPipe_Click()
    
    m_txtCsvZeichen.Text = "|"

End Sub

'################################################################################
'
Private Sub m_btnCsvDoppelpunkt_Click()
    
    m_txtCsvZeichen.Text = ":"

End Sub

'################################################################################
'
Private Sub m_btnCsvGleichKomma_Click()
    
    m_txtCsvZeichen.Text = ","

End Sub

'################################################################################
'
Private Sub m_btnCsvPunkt_Click()
    
    m_txtCsvZeichen.Text = "."

End Sub

'################################################################################
'
Private Sub m_btnCsvSemikolon_Click()
    
    m_txtCsvZeichen.Text = ";"

End Sub

'################################################################################
'
Private Sub m_btnCsvGleich_Click()
    
    m_txtCsvZeichen.Text = "="

End Sub

'################################################################################
'
Private Sub m_btnSetGatter0_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SET_TRENNZEICHEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnSetGatter0Ende_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SET_TRENNZEICHEN_VOR, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnSetGatter0Zurueck_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_SET_TRENNZEICHEN_ZURUECK, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartToZeile_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_TO_ZEILE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, , m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnDuplizierung_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_DUPLIZIERUNG, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnErstelleXmlFormat_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_ERSTELLE_XML, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_startGetStringLit_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_STRING_LIT, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnErstelleXmlFormat2_Click()

Dim str_x As String

    str_x = startMrStringer(FKT_ERSTELLE_XML_2, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
    
    str_x = startMrStringer(FKT_AUSRICHTER_STRING, str_x, -1, -1, , "x_attribut")
    
    str_x = startMrStringer(FKT_AUSRICHTER_STRING, str_x, -1, -1, , "/>")

    m_txtAusgabe.Text = Replace(str_x, "#Xp", "p")

End Sub

'################################################################################
'
Private Sub m_btnGrepWort_Click()

    If (m_txtEingabe.SelLength > 0) Then
        
        m_txtAusgabe.Text = startMrStringer(FKT_GREP_WORT, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
    
    Else
        
        m_txtAusgabe.Text = startMrStringer(FKT_EXTRAHIERE_WORTE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False, m_txtCsvZeichen.Text)
    
    End If

End Sub

'################################################################################
'
Private Sub m_btnSplit_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SET_TRENNZEICHEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
    
End Sub

'################################################################################
'
Private Sub m_btnStartClip_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_CLIP_POSITION, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartXmlJavaWriter_Click()

Dim str_x As String

    str_x = startMrStringer(FKT_JAVA_XML_WRITER_STRING, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)
    
    str_x = startMrStringer(FKT_AUSRICHTER_STRING, str_x, -1, -1, , "#Xp")
    
    str_x = startMrStringer(FKT_AUSRICHTER_STRING, str_x, -1, -1, , "TAG_VOR")

    m_txtAusgabe.Text = Replace(str_x, "#Xp", "p")

End Sub

'################################################################################
'
Private Sub m_btnStartZaehler_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_ZAEHLER, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStrReverse_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_GREP_WORT, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartCsvToZeile_Click()
    
    checkCsvSelektion

    m_txtAusgabe.Text = startMrStringer(FKT_CSV_2_ZEILE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False, m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartDebugAusgabe_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_DEBUG_AUSGABE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartFallunterscheidungVB_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_GENERATOR_IF, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartNamen_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_ERSTELLE_NAMEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartRemove_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_STRING_REMOVE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartSortierung_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SORTIEREN_ABC, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartSortierungLaenge_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SORTIEREN_LAENGE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartSortZufall_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SORTIEREN_ZUFALL, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartUnique_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_GET_UNIQUE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartReverse_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_UMDREHEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartDoppelteVorkommen_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_GET_DOPPELTE_VORKOMMEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartEinmaligeVorkommen_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_GET_EINMALIGE_VORKOMMEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartAusrichter1_Click()
 
    m_txtAusgabe.Text = startMrStringer(FKT_AUSRICHTER_POSITION, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnDeklaration_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_DEKLARATION, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartTrim_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_TRIM, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnTrimX_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_TRIM_X, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartUCaseLCase_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_UCASE_LCASE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartCsvSwap_Click()
    
    checkCsvSelektion

    m_txtAusgabe.Text = startMrStringer(FKT_CSV_SWAP, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False, m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnErstelleCsv_Click()

    checkCsvSelektion
    
    m_txtAusgabe.Text = startMrStringer(FKT_CSV_ERSTELLE_CSV, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False, m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartSplit_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SPLIT, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStringVb_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_STRING_IT, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartUmdrehen_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_UMDREHEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartJSON_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_JSON_LESEN_SCHREIBEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartDirEinlesen_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_GET_DIR, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnGeneratorJava_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_JAVA_GENERATOR, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartNotesLesenSchreiben_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_NOTES_LESEN_SCHREIBEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_startCsvCase_Click()

    checkCsvSelektion
    
    m_txtAusgabe.Text = startMrStringer(FKT_CSV_CASE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False, m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartPlaceX_Click()

    m_txtAusgabe.Text = placeStringX(m_txtEingabe.Text, m_txtEingabe2.Text, FKT_GENERATOR_IF_2, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartMove_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_VERSCHIEBEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_strBlock_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_ERSTELLE_BLOCK, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartRot13_Click()
        
    m_txtAusgabe.Text = rot13(m_txtEingabe.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartSpalte1_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_MARKIERE_VORNE_ODER_HINTEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartCmdRename_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_CMD_RENAME, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartSpalte2_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False, m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartSpalte3_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_MARKIERE_VORNE_UND_HINTEN_UND_DOPPLE, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub
'################################################################################
'
Private Sub m_btnStartReplaceX_Click()

    m_txtAusgabe.Text = startReplaceSuchWorte(m_txtEingabe.Text, m_txtEingabe2.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartVbToJava_Click()
    
    m_txtAusgabe.Text = generatorVbNachJava(m_txtEingabe.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartSetNull_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SET_NULL, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnStartSortDatum_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_SORTIEREN_DATUM, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

'################################################################################
'
Private Sub m_btnGrepAufnehmen_Click()
    
    m_txtAusgabe.Text = startMrStringer(FKT_GREP_PLUS_MINUS, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, True)

End Sub

'################################################################################
'
Private Sub m_btnGrepWeglassen_Click()
    
     m_txtAusgabe.Text = startMrStringer(FKT_GREP_PLUS_MINUS, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength, False)

End Sub

'################################################################################
'
Private Sub m_btnCsvExcel_Click()
    
    Call fkCsvExport2Excel.startCsv2Excel(m_txtEingabe.Text, m_txtCsvZeichen.Text)

End Sub

'################################################################################
'
Private Sub m_btnStartBlockZufall_Click()

    m_txtAusgabe.Text = startMrStringer(FKT_BLOCK_ZUFALL, m_txtEingabe.Text, m_txtEingabe.SelStart, m_txtEingabe.SelLength)

End Sub

