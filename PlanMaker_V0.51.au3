; :wrap=none:noTabs=false:collapseFolds=0:maxLineLen=80:mode=autoitscript:tabSize=4:indentSize=4:noWordSep=_@:deepIndent=false:wordBreakChars=,+-\=<>/?^&*:folding=indent:
#include-once
; _PlanMaker.au3 V0.51
; Thorsten Willert
; Sun Jan 21 13:56:31 CET 2018 @580 /Internet-Zeit/
; Tested with Sofmaker-Office: Planmaker 2016, 2018 and FreeOffice

#cs

V.51
tested with PlanMaker 2018
added: Console-output to __PlanMaker_COMErrFunc

V.50
added: _PlanMaker_RangeFormat
added: _PlanMaker_RangeMethode

V0.40
added: _PlanMaker_PageSetup
added: _PlanMaker_Print
example:

V0.30
added: _PlanMaker_FormatNumber

V0.20
added: _PlanMaker_FormatShading
changed: _PlanMaker_FormatFont (range as object or string)
changed: _PlanMaker_FormatBorder* (range as object or string)
changed: _PlanMaker_SheetFromArray (range as object or string)
changed: _PlanMaker_BookAttach (supports now only the "filename", too)
changed: __IsRange (checks now for object-name "IDocRange", too)
optimized: _PlanMaker_ScreenUpdate

V0.17
added: _PlanMaker_FormatBorder
added: _PlanMaker_FormatBorders
added: _PlanMaker_FormatBorder_All
added: _PlanMaker_FormatBorder_Frame
added: _PlanMaker_FormatBorder_Inner
fixed: _FormatFont(error: ManualApply)

V0.16
fixed: various Au3Check warnings
added: Constants for cell-borders, type and texture

V0.15
added: constants for named color-values
added: function _PlanMaker_Color2SmoColor (converts Int-RGB OR Hex-RGB values to SoftMaker-Office HEX-BGR values)

#ce

#cs
; #CURRENT# ====================================================================
_PlanMaker_BookAttach
_PlanMaker_BookClose
_PlanMaker_BookNew
_PlanMaker_BookOpen
_PlanMaker_BookSave
_PlanMaker_BookSaveAs
_PlanMaker_CellRangeByPosition
_PlanMaker_CellRead
_PlanMaker_CellWrite
_PlanMaker_DocumentPropertyGet
_PlanMaker_DocumentPropertyGetAll
_PlanMaker_DocumentPropertySet
_PlanMaker_FormatBorder
_PlanMaker_FormatBorders
_PlanMaker_FormatBorder_All
_PlanMaker_FormatBorder_Frame
_PlanMaker_FormatBorder_Inner
_PlanMaker_FormatFont
_PlanMaker_FormatNumber
_PlanMaker_FormatShading
_PlanMaker_FormulaRead
_PlanMaker_FormulaWrite
_PlanMaker_PageSetup
_PlanMaker_Quit
_PlanMaker_RGB2SmoColor
_PlanMaker_ScreenUpdate
_PlanMaker_SheetActivate
_PlanMaker_SheetAddNew
_PlanMaker_SheetDelete
_PlanMaker_SheetList
_PlanMaker_SheetFromArray
_PlanMaker_SheetToArray
_PlanMaker_UserPropertyGet
_PlanMaker_UserPropertyGetAll
_PlanMaker_UserPropertySet
__ConvertToLetter
__IsRange
#ce

;Konstanten aus dem BasicMaker Handbuch
Global Enum $_smoDoNotSaveChanges = 0, _
		$_smoPromptToSaveChanges, _
		$_smoSaveChanges
#cs
	WindowState
	smoWindowStateNormal = 1 ' normal
	smoWindowStateMinimize = 2 ' minimiert
	smoWindowStateMaximize = 3 ' maximiert

	Calculation
	pmCalculationAutomatic = 0 ' Berechnungen automatisch aktualisieren
	pmCalculationManual = 1 ' Berechnungen manuell aktualisieren

	DisplayCommentIndicator
	pmNoIndicator = 0 ' Weder Kommentare noch gelbes Dreieck
	pmCommentIndicatorOnly = 1 ' Nur ein gelbes Hinweisdreieck in der Zelle
	pmCommentOnly = 2 ' Kommentare zeigen, aber kein Hinweisdreieck
	pmCommentAndIndicator = 3 ' Sowohl Kommentare als auch Dreieck zeigen

#ce
;Userproperty
Global Enum $_smoUserHomeAddressName = 1, _ 	 		 ; Nachname (privat)
		$_smoUserHomeAddressFirstName, _ 				 ; Vorname (privat)
		$_smoUserHomeAddressStreet, _ 					 ; Straß (privat)
		$_smoUserHomeAddressZip, _ 						 ; Postleitzahl (privat)
		$_smoUserHomeAddressCity, _ 					 ; Stadt (privat)
		$_smoUserHomeAddressPhone1, _ 					 ; Telefon (privat)
		$_smoUserHomeAddressFax, _ 						 ; Telefax (privat)
		$_smoUserHomeAddressEmail, _ 					 ; E-Mail-Adresse (privat)
		$_smoUserHomeAddressPhone2, _ 					 ; Mobiltelefon o.Ä. (privat)
		$_smoUserHomeAddressHomepage, _ 				 ; Homepage (privat)
		$_smoUserBusinessAddressName, _ 				 ; Nachname (geschÄftlich)
		$_smoUserBusinessAddressFirstName, _ 			 ; Vorname (geschÄftlich)
		$_smoUserBusinessAddressCompany, _ 				 ; Firma (geschÄftlich)
		$_smoUserBusinessAddressDepartment, _ 			 ; Abteilung (geschÄftlich)
		$_smoUserBusinessAddressStreet, _ 				 ; Straß (geschÄftlich)
		$_smoUserBusinessAddressZip, _ 					 ; Postleitzahl (geschÄftlich)
		$_smoUserBusinessAddressCity, _ 				 ; Stadt (geschÄftlich)
		$_smoUserBusinessAddressPhone1, _ 				 ; Telefon (geschÄftlich)
		$_smoUserBusinessAddressFax, _ 					 ; Telefax (geschÄftlich)
		$_smoUserBusinessAddressEmail, _ 				 ; E-Mail-Adresse (geschÄftlich)
		$_smoUserBusinessAddressPhone2, _ 				 ; Mobiltelefon o.Ä. (geschÄftlich)
		$_smoUserBusinessAddressHomepage, _ 			 ; Homepage (geschÄftlich)
		$_smoUserHomeAddressInitials, _ 				 ; Initialen des Benutzers (privat)
		$_smoUserBusinessAddressInitials ; Initialen des Benutzers (geschÄftlich)


;FileFormat (UnabhÄngig vom Übergebenen Parameter FileFormat versucht PlanMaker stets, das Dateiformat selbst zu erkennen,
; und ignoriert offensichtlich falsche Angaben)
Global Enum $_pmFormatDocument = 0, _ 	 		 ; 0 PlanMaker-Dokument - dies ist die Standardeinstellung
		$_pmFormatTemplate, _ 					 ; 1 PlanMaker-Dokumentvorlage
		$_pmFormatExcel97, _ 					 ; 2 Excel 97/2000/XP
		$_pmFormatExcel5, _ 					 ; 3 Excel 5.0/7.0
		$_pmFormatExcelTemplate, _ 				 ; 4 Excel-Dokumentvorlage
		$_pmFormatSYLK, _ 						 ; 5 Sylk
		$_pmFormatRTF, _ 						 ; 7 Rich Text Format
		$_pmFormatTextMaker = 7, _				 ; 7 TextMaker (= RTF)
		$_pmFormatHTML, _ 						 ; 8 HTML-Dokument
		$_pmFormatdBaseDOS, _ 					 ; 9 dBASE-Datenbank mit DOS-Zeichensatz
		$_pmFormatdBaseAnsi, _ 					 ; 10 dBASE-Datenbank mit Windows-Zeichensatz
		$_pmFormatDIF, _ 						 ; 11 Textdatei mit Windows-Zeichensatz
		$_pmFormatPlainTextAnsi, _ 				 ; 12 Textdatei mit Windows-Zeichensatz
		$_pmFormatPlainTextDOS, _ 				 ; 13 Textdatei mit DOS-Zeichensatz
		$_pmFormatPlainTextUnix, _ 				 ; 14 Textdatei mit ANSI-Zeichensatz fÜr UNIX, Linux und FreeBSD
		$_pmFormatPlainTextUnicode, _ 			 ; 15 Textdatei mit Unicode-Zeichensatz
		$_pmFormatdBaseUnicode = 19, _			 ; 18 dBASE-Datenbank mit Unicode-Zeichensatz
		$_pmFormatPlainTextUTF8 = 21, _			 ; 21 Textdatei mit UTF8-Zeichensatz
		$_pmFormatMSXML = 23, _					 ; 23 Excel ab 2007 (.xlsx)
		$_pmFormatPM2008 = 26, _				 ; 26 PlanMaker 2008-Dokument
		$_pmFormatPM2010 ; 27 PlanMaker 2010-Dokument

;TextMarker (Gibt bei den Textdatei-Formaten an, mit welchem Zeichen Textfelder umgeben sind)
Global Enum $_pmImportTextMarkerNone = 0, _	 	 ; Text ist nicht speziell markiert
		$_pmImportTextMarkerApostrophe, _ 		 ; Apostrophe
		$_pmImportTextMarkerQmark ; AnfÜhrungszeichen


;DocumentProperty
Global Enum $_smoPropertyTitle = 1, _ 				 ; "Title"
		$_smoPropertySubject, _ 					 ; "Subject"
		$_smoPropertyAuthor, _ 						 ; "Author"
		$_smoPropertyKeywords, _ 					 ; "Keywords"
		$_smoPropertyComments, _ 					 ; "Comments"
		$_smoPropertyAppName, _ 					 ; "Application name"
		$_smoPropertyTimeLastPrinted, _ 			 ; "Last print date"
		$_smoPropertyTimeCreated, _ 				 ; "Creation date"
		$_smoPropertyTimeLastSaved, _ 				 ; "Last save time"
		$_smoPropertyPages = 18, _ 				 ; "Number of pages"
		$_smoPropertyCells, _ 						 ; "Number of cells"
		$_smoPropertyTextCells, _ 					 ; "Number of cells with text"
		$_smoPropertyNumericCells, _ 				 ; "Number of cells with numbers"
		$_smoPropertyFormulaCells, _ 				 ; "Number of cells with formulas"
		$_smoPropertyNotes, _ 						 ; "Number of comments"
		$_smoPropertySheets, _ 						 ; "Number of worksheets"
		$_smoPropertyCharts, _ 						 ; "Number of charts"
		$_smoPropertyPictures, _ 					 ; "Number of pictures"
		$_smoPropertyOLEObjects, _ 					 ; "Number of OLE objects"
		$_smoPropertyDrawings, _ 					 ; "Number of drawings"
		$_smoPropertyTextFrames ; "Number of text frames"

;smoPropertyKeystrokes = 10 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertyCharacters = 11 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertyWords = 12 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertySentences = 13 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertyParas = 14 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertyChapters = 15 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertySections = 16 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertyLines = 17 ' - (bei PlanMaker nicht verfÜgbar)

;smoPropertyTables = 30 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertyFootnotes = 31 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertyAvgWordLength = 32 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertyAvgCharactersSentence = 33 ' - (bei PlanMaker nicht verfÜgbar)
;smoPropertyAvgWordsSentence = 34 ' - (bei PlanMaker nicht verfÜgbar)

#cs
	Type
	smoPropertyTypeBoolean = 0 ' Boolean
	smoPropertyTypeDate = 1 ' Datum
	smoPropertyTypeFloat = 2 ' Fliesskommawert
	smoPropertyTypeNumber = 3 ' Ganzzahl
	smoPropertyTypeString = 4 ' Zeichenkette
#ce
;	Orientation
Global Enum $_smoOrientLandscape = 0, _ ; Querformat
	$_smoOrientPortrait;  Hochformat

;	PaperSize
Global Const $_smoPaperCustom = -1
Global Enum $_smoPaperLetter = 1, _
		$_smoPaperLetterSmall, _
		$_smoPaperTabloid, _
		$_smoPaperLedger, _
		$_smoPaperLegal, _
		$_smoPaperStatement, _
		$_smoPaperExecutive, _
		$_smoPaperA3, _
		$_smoPaperA4, _
		$_smoPaperA4Small, _
		$_smoPaperA5, _
		$_smoPaperB4, _
		$_smoPaperB5, _
		$_smoPaperFolio, _
		$_smoPaperQuarto, _
		$_smoPaper10x14, _
		$_smoPaper11x17, _
		$_smoPaperNote, _
		$_smoPaperEnvelope9, _
		$_smoPaperEnvelope10, _
		$_smoPaperEnvelope11, _
		$_smoPaperEnvelope12, _
		$_smoPaperEnvelope14, _
		$_smoPaperCSheet, _
		$_smoPaperDSheet, _
		$_smoPaperESheet, _
		$_smoPaperEnvelopeDL, _
		$_smoPaperEnvelopeC5, _
		$_smoPaperEnvelopeC3, _
		$_smoPaperEnvelopeC4, _
		$_smoPaperEnvelopeC6, _
		$_smoPaperEnvelopeC65, _
		$_smoPaperEnvelopeB4, _
		$_smoPaperEnvelopeB5, _
		$_smoPaperEnvelopeB6, _
		$_smoPaperEnvelopeItaly, _
		$_smoPaperEnvelopeMonarch, _
		$_smoPaperEnvelopePersonal, _
		$_smoPaperFanfoldUS, _
		$_smoPaperFanfoldStdGerman, _
		$_smoPaperFanfoldLegalGerman

;	PrintComments
Global Enum $_pmPrintNoComments = 0, _ ; Keine Kommentare drucken
		$_pmPrintInPlace ; Kommentare ausdrucken

;	Order
Global Enum $_pmOverThenDown = 0, _ ; Von links nach rechts
		$_pmDownThenOver = 1 ; Von oben nach unten

;	HorizontalAlignment
Global Enum $_pmHAlignGeneral = 0, _ ; Standard
		$_pmHAlignLeft, _ ; Linksbuendig
		$_pmHAlignRight, _ ; Rechtsbuendig
		$_pmHAlignCenter, _ ; Zentriert
		$_pmHAlignJustify, _ ; Blocksatz
		$_pmHAlignCenterAcrossSelection ; Zentriert Ueber Spalten

;	VerticalAlignment
Global Enum $_pmVAlignTop = 0, _ ; Oben
		$_pmVAlignCenter, _ ; Zentriert
		$_pmVAlignBottom, _ ; Unten
		$_pmVAlignJustify ; Vertikaler Blocksatz

;	Type
Global Enum $_pmNumberGeneral = 0, _ ; Standard
		$_pmNumberDecimal, _ ; Zahl
		$_pmNumberScientific, _ ; Wissenschaftlich
		$_pmNumberFraction, _ ; Bruch (fuer den Nenner siehe Digits-Eigenschaft)
		$_pmNumberDate, _ ; Datum/Uhrzeit (siehe Hinweis)
		$_pmNumberPercentage, _ ; Prozent
		$_pmNumberCurrency, _ ; Waehrung (siehe Hinweis)
		$_pmNumberBoolean, _ ; Wahrheitswert
		$_pmNumberCustom, _ ; Benutzerdefiniert (siehe Hinweis)
		$_pmNumberText, _ ; Text
		$_pmNumberAccounting ; Buchhaltung (siehe Hinweis)

#cs
	Currency
	EUR Euro
	USD US Dollar
	CAD Kanadische Dollar
	AUD Australische Dollar
	JPY Japanische Yen
	RUB Russische Rubel
	BEF Belgische Francs
	CHF Schweizer Franken
	DEM Deutsche Mark
	ESP Spanische Peseten
	FRF Franzoesische Francs
	LUF Luxemburgische Francs
	NLG Niederlaendische Gulden
	PTE Portugiesische Escudos
#ce

; Color
Global Const $_smoColorAutomatic = -1 ; Automatisch (siehe BasicMaker-Manual)
Global Const $_smoColorTransparent = -1 ; Transparent (siehe BasicMaker-Manual)
Global Const $_smoColorBlack = "&h0"
Global Const $_smoColorBlue = "&hFF0000"
Global Const $_smoColorBrightGreen = "&h00FF00"
Global Const $_smoColorRed = "&h0000FF"
Global Const $_smoColorYellow = "&h00FFFF"
Global Const $_smoColorTurquoise = "&hFFFF00"
Global Const $_smoColorViolet = "&h800080"
Global Const $_smoColorWhite = "&hFFFFFF"
Global Const $_smoColorIndigo = "&h993333"
Global Const $_smoColorOliveGreen = "&h003333"
Global Const $_smoColorPaleBlue = "&hFFCC99"
Global Const $_smoColorPlum = "&h663399"
Global Const $_smoColorRose = "&hCC99FF"
Global Const $_smoColorSeaGreen = "&h669933"
Global Const $_smoColorSkyBlue = "&hFFCC00"
Global Const $_smoColorTan = "&h99CCFF"
Global Const $_smoColorTeal = "&h808000"
Global Const $_smoColorAqua = "&hCCCC33"
Global Const $_smoColorBlueGray = "&h996666"
Global Const $_smoColorBrown = "&h003399"
Global Const $_smoColorGold = "&h00CCFF"
Global Const $_smoColorGreen = "&h008000"
Global Const $_smoColorLavender = "&hFF99CC"
Global Const $_smoColorLime = "&h00CC99"
Global Const $_smoColorOrange = "&h0066FF"
Global Const $_smoColorPink = "&hFF00FF"
Global Const $_smoColorLightBlue = "&hFF6633"
Global Const $_smoColorLightOrange = "&h0099FF"
Global Const $_smoColorLightYellow = "&h99FFFF"
Global Const $_smoColorLightGreen = "&hCCFFCC"
Global Const $_smoColorLightTurquoise = "&hFFFFCC"
Global Const $_smoColorDarkBlue = "&h800000"
Global Const $_smoColorDarkGreen = "&h003300"
Global Const $_smoColorDarkRed = "&h000080"
Global Const $_smoColorDarkTeal = "&h663300"
Global Const $_smoColorDarkYellow = "&h008080"
Global Const $_smoColorGray05 = "&hF3F3F3"
Global Const $_smoColorGray10 = "&hE6E6E6"
Global Const $_smoColorGray125 = "&hE0E0E0"
Global Const $_smoColorGray15 = "&hD9D9D9"
Global Const $_smoColorGray20 = "&hCCCCCC"
Global Const $_smoColorGray25 = "&hC0C0C0"
Global Const $_smoColorGray30 = "&hB3B3B3"
Global Const $_smoColorGray35 = "&hA6A6A6"
Global Const $_smoColorGray375 = "&hA0A0A0"
Global Const $_smoColorGray40 = "&h999999"
Global Const $_smoColorGray45 = "&h8C8C8C"
Global Const $_smoColorGray50 = "&h808080"
Global Const $_smoColorGray55 = "&h737373"
Global Const $_smoColorGray60 = "&h666666"
Global Const $_smoColorGray625 = "&h606060"
Global Const $_smoColorGray65 = "&h595959"
Global Const $_smoColorGray75 = "&h404040"
Global Const $_smoColorGray85 = "&h262626"
Global Const $_smoColorGray90 = "&h191919"
Global Const $_smoColorGray70 = "&h4C4C4C"
Global Const $_smoColorGray80 = "&h333333"
Global Const $_smoColorGray875 = "&h202020"
Global Const $_smoColorGray95 = "&hC0C0C0"

;ColorIndex
Global Enum $_smoColorIndexAutomatic = -1, _ 		 ; Automatisch (siehe unten)
		$_smoColorIndexTransparent = -1, _ 		 	 ; Transparent (siehe unten)
		$_smoColorIndexBlack = 0, _ 				 ; Schwarz
		$_smoColorIndexBlue, _ 						 ; Blau
		$_smoColorIndexCyan, _ 						 ; Zyanblau
		$_smoColorIndexGreen, _ 					 ; GrÜn
		$_smoColorIndexMagenta, _ 					 ; Magenta
		$_smoColorIndexRed, _ 						 ; Rot
		$_smoColorIndexYellow, _ 					 ; Gelb
		$_smoColorIndexWhite, _ 					 ; Weiß¸
		$_smoColorIndexDarkBlue, _ 					 ; Dunkelblau
		$_smoColorIndexDarkCyan, _ 					 ; Dunkles Zyanblau
		$_smoColorIndexDarkGreen, _ 				 ; DunkelgrÜn
		$_smoColorIndexDarkMagenta, _ 				 ; Dunkles Magenta
		$_smoColorIndexDarkRed, _ 					 ; Dunkelrot
		$_smoColorIndexBrown, _ 					 ; Braun
		$_smoColorIndexDarkGray, _ 					 ; Dunkelgrau
		$_smoColorIndexLightGray ; Hellgrau

;Underline
Global Enum $_pmUnderlineNone = 0, _ 		 		 ; aus
		$_pmUnderlineSingle, _ 						 ; einfach durchgehend
		$_pmUnderlineDouble, _ 						 ; doppelt durchgehend
		$_pmUnderlineWords, _ 						 ; einfach wortweise
		$_pmUnderlineWordsDouble ; doppelt wortweise

;	Borders
Global Enum Step -1 $_pmBorderTop = -1, _ 			 ;Linie oberhalb der Zellen
		$_pmBorderLeft, _ 							 ; Linie links der Zellen
		$_pmBorderBottom, _ 						 ; Linie unterhalb der Zellen
		$_pmBorderRight, _ 							 ; Linie rechts der Zellen
		$_pmBorderHorizontal, _ 					 ; Horizontale Gitternetzlinien
		$_pmBorderVertical ;Vertikale Gitternetzlinien

;	Type
Global Enum $_pmLineStyleNone = 0, _ 				 ; Keine Linie
		$_pmLineStyleSingle, _ 						 ; Einfache Linie
		$_pmLineStyleDouble ; Doppelte Linie

;	Texture
Global Enum $_smoPatternNone = 0, _ 				 ;(Kein Muster)
		$_smoPatternHalftone, _
		$_smoPatternRightDiagCoarse, _
		$_smoPatternLeftDiagCoarse, _
		$_smoPatternHashDiagCoarse, _
		$_smoPatternVertCoarse, _
		$_smoPatternHorzCoarse, _
		$_smoPatternHashCoarse, _
		$_smoPatternRightDiagFine, _
		$_smoPatternLeftDiagFine, _
		$_smoPatternHashDiagFine, _
		$_smoPatternVertFine, _
		$_smoPatternHorzFine, _
		$_smoPatternHashFine

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_BookAttach
; Description ...:
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_BookAttach($sFileName)
; Parameter(s): .: $sFileName   - Filename or path with filename
; Return Value ..: Success      - Object of the workbook
;                  Failure      - 0
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Jun 19 00:32:28 CEST 2017 @980 /Internet-Zeit/
; Version .......: 0.3
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_BookAttach($sFileName)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not FileExists($sFileName) Then Return SetError(1, 1, 0)

	Local $oPM = ObjGet("", "PlanMaker.Application")
	If @error Or Not IsObj($oPM) Then Return SetError(2, @error, 0)

	Local $iCount = $oPM.Application.Workbooks.Count
	For $i = 1 To $iCount
		If $sFileName = $oPM.Application.Workbooks.Item($i).FullName Or _
				$sFileName = $oPM.Application.Workbooks.Item($i).Name Then Return SetError(0, 0, $oPM.Application.ActiveWorkbook)
	Next

	Return SetError(3, 1, 0)
EndFunc   ;==>_PlanMaker_BookAttach

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_BookClose
; Description ...: Closes the current workbook
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_BookClose(ByRef $oPM[, $Option = $smoSaveChanges])
; Parameter(s): .: $oPM         -
;                  $Option      - Optional: (Default = $smoSaveChanges) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....:
; Date ..........: Thu Mar 23 16:11:04 CET 2017
; Version .......: 0.1
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_BookClose(ByRef $oWB, $Option = $_smoSaveChanges)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oWB) Then Return SetError(1, 1, 0)

	$oWB.Close($Option)

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_BookClose

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_BookNew
; Description ...: Adds a new workbook
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_BookNew([$bVisible = False[, $bActivate = False]])
; Parameter(s): .: $bVisible    - Optional: (Default = False) :
;                  $bActivate   - Optional: (Default = False) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 27 23:25:57 CEST 2017
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_BookNew($bVisible = False, $bActivate = False)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	Local $oPM = ObjCreate("PlanMaker.Application")
	If @error Or Not IsObj($oPM) Then Return SetError(1, @error, 0)
	Local $oWB = $oPM.Workbooks.Add
	If @error Or Not IsObj($oWB) Then Return SetError(1, @error, 0)

	$oPM.Visible = $bVisible
	$oWB.Activate = $bActivate

	Return SetError(0, 0, $oWB)
EndFunc   ;==>_PlanMaker_BookNew

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_BookOpen
; Description ...: Opens a existing workbook
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_BookOpen($sFileName[, $bVisible = False[, $bActivate = False]])
; Parameter(s): .: $sFileName   -
;                  $bVisible    - Optional: (Default = False) :
;                  $bActivate   - Optional: (Default = False) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 27 10:41:39 CEST 2017
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_BookOpen($sFileName, $bVisible = False, $bActivate = False)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not FileExists($sFileName) Then Return SetError(1, 1, 0)

	Local $oPM = ObjCreate("PlanMaker.Application")
	If @error Or Not IsObj($oPM) Then Return SetError(2, @error, 0)

	Local $iCount = $oPM.Application.Workbooks.Count
	For $i = 1 To $iCount
		If $sFileName = $oPM.Application.Workbooks.Item($i).FullName Then Return SetError(3, 1, $oPM.Application.ActiveWorkbook)
	Next

	Local $oWB = $oPM.Workbooks.Open($sFileName)

	$oPM.Visible = $bVisible
	$oWB.Activate = $bActivate

	Return SetError(0, 0, $oWB)
EndFunc   ;==>_PlanMaker_BookOpen

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_BookSave
; Description ...: Saves the current workbook
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_BookSave(ByRef $oPM)
; Parameter(s): .: $oPM         -
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Thu Mar 23 16:09:59 CET 2017
; Version .......: 0.1
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_BookSave(ByRef $oWB)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oWB) Then Return SetError(1, 1, 0)

	$oWB.Save

	Return SetError(0, 0, $oWB.Saved)
EndFunc   ;==>_PlanMaker_BookSave

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_BookSaveAs
; Description ...: Saves the current workbook as ...
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_BookSaveAs(ByRef $oPM, $sFileName[, $iFileFormat = $_pmFormatDocument[, $sDelimiter = @TAB[, $sTextmarker = $_pmImportTextMarkerNone]]])
; Parameter(s): .: $oPM         -
;                  $sFileName   -
;                  $iFileFormat - Optional: (Default = $_pmFormatDocument) :
;                  $sDelimiter  - Optional: (Default = @TAB) :
;                  $sTextmarker - Optional: (Default = $_pmImportTextMarkerNone) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Mar 26 23:53:42 CEST 2017
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_BookSaveAs(ByRef $oWB, $sFileName, $iFileFormat = $_pmFormatDocument, $sDelimiter = @TAB, $sTextmarker = $_pmImportTextMarkerNone)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oWB) Then Return SetError(1, 1, 0)

	If $iFileFormat > 10 And $iFileFormat < 21 Then
		$oWB.SaveAs($sFileName, $iFileFormat)
	Else
		$oWB.SaveAs($sFileName, $iFileFormat, $sDelimiter, $sTextmarker)
	EndIf

	Return SetError(0, 0, $oWB.Saved)
EndFunc   ;==>_PlanMaker_BookSaveAs

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_CellRangeByPosition
; Description ...: Returns a range-string by the position
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_CellRangeByPosition($iStartColumn, $iStartRow, $iEndColumn, $iEndRow)
; Parameter(s): .: $iStartColumn -
;                  $iStartRow   -
;                  $iEndColumn  -
;                  $iEndRow     -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Mar 28 20:55:05 CEST 2017
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_CellRangeByPosition($iStartColumn, $iStartRow, $iEndColumn, $iEndRow)
	If Not ($iStartRow > 0 And _
			$iStartColumn > 0 And _
			$iEndRow > 0 And _
			$iEndColumn > 0) Then Return SetError(1, 1, 0)
	If Not (IsInt($iStartColumn) And _
			IsInt($iStartRow) And _
			IsInt($iEndColumn) And _
			IsInt($iEndRow)) Then Return SetError(1, 2, 0)

	Return SetError(0, 0, _
			(__ConvertToLetter($iStartColumn) & $iStartRow & ":" & __ConvertToLetter($iEndColumn) & $iEndRow) _
			)
EndFunc   ;==>_PlanMaker_CellRangeByPosition

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_CellRead
; Description ...: Reads the value of a cell
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_CellRead(ByRef $oObj[, $iRow = 1[, $iCol = 1]])
; Parameter(s): .: $oPM         -
;                  $iRow        - Optional: (Default = 1) :
;                  $iCol        - Optional: (Default = 1) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Fri Jun 16 18:38:13 CEST 2017 @734 /Internet-Zeit/
; Version .......: 0.3
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_CellRead(ByRef $oObj, $iRow = 1, $iCol = 1)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, "")

	Local $vValue = $oObj.ActiveSheet.Cells.Item($iRow, $iCol).Value

	Return SetError(0, 0, $vValue)
EndFunc   ;==>_PlanMaker_CellRead

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_CellWrite
; Description ...: Writes the value of a cell
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_CellWrite(ByRef $oObj, $vValue[, $iRow = 1[, $iCol = 1]])
; Parameter(s): .: $oPM         -
;                  $vValue      -
;                  $iRow        - Optional: (Default = 1) :
;                  $iCol        - Optional: (Default = 1) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 27 23:24:39 CEST 2017
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_CellWrite(ByRef $oObj, $vValue, $iRow = 1, $iCol = 1)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)

	$oObj.ActiveSheet.Cells.Item($iRow, $iCol).Value = $vValue

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_CellWrite

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_Color2SmoColor
; Description ...: Converts RGB or HEX color to SoftMaker-BGR color
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_RGB2SmoColor($vR, $iG, $iB)
; Parameter(s): .: $vR          - red value (or HEX-value e.g. "ff00aa" OR RGB-value e.g. "255,0,128")
;                  $vG          - green value (hex OR int)
;                  $vB          - blue value (hex OR int)
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Jun 11 14:57:56 CEST 2017
; Version .......: 1.0
; Example .......: Accepts RGB-values in the following formats
; _PlanMaker_Color2SmoColor("255,0,0")
; _PlanMaker_Color2SmoColor(255,0,0)
; _PlanMaker_Color2SmoColor("FF0000")
; _PlanMaker_Color2SmoColor("FF","00","00")
; ==============================================================================
Func _PlanMaker_Color2SmoColor($vR, $vG = "", $vB = "")
	Local $sRet, $a

	If $vG = "" And $vB = "" Then
		If StringInStr($vR, ",") Then
			$a = StringSplit($vR, ",", 2)
			If Not @error Then
				$vR = $a[0]
				$vG = $a[1]
				$vB = $a[2]
			EndIf
		ElseIf StringRegExp($vR, '^[[:xdigit:]]+$') Then
			$a = StringRegExp($vR, "^(\w{2})(\w{2})(\w{2})$", 1)
			If Not @error Then
				$vR = $a[0]
				$vG = $a[1]
				$vB = $a[2]
			EndIf
		EndIf
	EndIf

	If StringRegExp($vR, '^[[:xdigit:]]{2}$') Then
		$sRet = $vB & $vG & $vR
	ElseIf StringRegExp($vR, '^[[:digit:]]{1,3}$') Then
		$sRet = Hex($vB, 2) & Hex($vG, 2) & Hex($vR, 2)
	EndIf

	Return "&h" & $sRet
EndFunc   ;==>_PlanMaker_Color2SmoColor

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_DocumentPropertyGet
; Description ...: Returns a property of the workbook
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_DocumentPropertyGet(ByRef $oObj, $iProperty)
; Parameter(s): .: $oPM         -
;                  $iProperty   -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 27 23:24:23 CEST 2017
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_DocumentPropertyGet(ByRef $oObj, $iProperty)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)

	If $oObj.ActiveWorkbook.BuiltInDocumentProperties($iProperty).Valid = True Then
		Return SetError(0, 0, $oObj.ActiveWorkbook.BuiltInDocumentProperties($iProperty).Value)
	Else
		Return SetError(1, 2, 0)
	EndIf

EndFunc   ;==>_PlanMaker_DocumentPropertyGet

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_DocumentPropertyGetAll
; Description ...: Returns an array of all document properties
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_DocumentPropertyGetAll(ByRef $oPM)
; Parameter(s): .: $oPM         -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Mar 28 14:55:42 CEST 2017
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_DocumentPropertyGetAll(ByRef $oPM)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oPM) Then Return SetError(1, 1, 0)

	Local $aProperties[23] = [22]

	For $i = 1 To 22
		If $oPM.BuiltInDocumentProperties($i).Valid = True Then $aProperties[$i] = $oPM.BuiltInDocumentProperties($i).Value
	Next

	Return SetError(0, 0, $aProperties)
EndFunc   ;==>_PlanMaker_DocumentPropertyGetAll

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_DocumentPropertySet
; Description ...: Sets an document property
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_DocumentPropertySet(ByRef $oPM, $iProperty, $sValue)
; Parameter(s): .: $oPM         -
;                  $iProperty   -
;                  $sValue      -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Mar 28 14:57:49 CEST 2017
; Version .......: 0.3
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_DocumentPropertySet(ByRef $oPM, $iProperty, $sValue)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oPM) Then Return SetError(1, 1, 0)
	If $iProperty < 1 Or $iProperty > 5 Then Return SetError(1, 2, 0)
	If $oPM.ActiveWorkbook.BuiltInDocumentProperties($iProperty).Valid = False Then Return SetError(1, 3, 0)

	$oPM.ActiveWorkbook.BuiltInDocumentProperties($iProperty).Value = $sValue

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_DocumentPropertySet


Func _PlanMaker_RangeMethode(ByRef $oObj, $vRange, $sMethode)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)
	If Not __IsRange($vRange) Then Return SetError(1, 2, 0)

	Local $oRng = (IsObj($vRange) ? $vRange : $oObj.ActiveSheet.Range($vRange))

	Local $sKey = StringLower($sMethode)

	With $oRng
		Switch $sKey
			Case "autofit"
				.AutoFit
			Case "applyformatting"
				.ApplyFormatting
			Case "select"
				.Select
			Case "copy"
				.Copy
			Case "cut"
				.Cut
			Case "paste"
				.Paste
			Case "clear"
				.Clear
			Case "clearcontents"
				.ClearContents
			Case "clearformats"
				.ClearFormats
			Case "clearconditionalformatting"
				.ClearConditionalFormatting
			Case "clearcomments"
				.ClearComments
			Case "clearinputvalidation"
				.ClearInputValidation
			Case Else
				Return SetError(1, 3, 0)
		EndSwitch
	EndWith

EndFunc

Func _PlanMaker_RangeFormat(ByRef $oObj, $vRange, $aFormat = Default)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)
	If Not __IsRange($vRange) Then Return SetError(1, 2, 0)

	Local $oRng, $iEnd, $sKey, $vValue, $iSize = 0

	$oRng = (IsObj($vRange) ? $vRange : $oObj.ActiveSheet.Range($vRange))

	If Not IsKeyword($aFormat) Then
		If Not IsArray($aFormat) Then Return SetError(1, 4, 0)

		$iEnd = UBound($aFormat) - 1

			With $oRng
				For $i = 0 To $iEnd
					$sKey = StringLower($aFormat[$i][0])
					$vValue = $aFormat[$i][1]

					Switch $sKey
						Case "name"
							.Name = $vValue
						Case "horizontalalignment"
							.HorizontalAlignment = $vValue
						Case "verticalalignment"
							.VerticalAlignment = $vValue
						Case "WrapText"
							.wraptext = $vValue
						Case "leftpadding"
							.LeftPadding = $vValue
						Case "rightpadding"
							.RightPadding = $vValue
						Case "toppadding"
							.TopPadding = $vValue
						Case "bottompadding"
							.BottomPadding = $vValue
						Case "mergecells"
							.MergeCells = $vValue
						Case "orientation"
							.Orientation = $vValue
						Case "verticaltext"
							.VerticalText = $vValue
						Case "pagebreakcol"
							.PageBreakCol = $vValue
						Case "pagebreakrow"
							.PageBreakRow = $vValue
						Case "comment"
							.Comment = $vValue
						Case "locked"
							.Locked = $vValue
						Case "formulahidden"
							.FormulaHidden = $vValue
						Case "cellhidden"
							.CellHidden = $vValue
						Case "nonprintable"
							.Nonprintable = $vValue
						Case "hidden"
							.Hidden = $vValue
						Case "rowheight"
							.RowHeight = $vValue
						Case "columnwidth"
							.ColumnWidth = $vValue
					EndSwitch

				Next
			EndWith
	EndIf

EndFunc

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_FormatBorder
; Description ...: Border-format of the range
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_FormatBorder(ByRef $oObj, $sRange, $iBorder[, $aFormat = Default])
; Parameter(s): .: $oObj        -
;                  $sRange      -
;                  $iBorder     -
;                  $aFormat     - Optional: (Default = Default) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Jun 18 16:37:11 CEST 2017
; Version .......: 1.0
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_FormatBorder(ByRef $oObj, $vRange, $iBorder, $aFormat = Default)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)
	If Not __IsRange($vRange) Then Return SetError(1, 2, 0)
	If ($iBorder > -1) Or ($iBorder < -6) Then Return SetError(1, 3, 0)

	Local $oRng, $iEnd, $sKey, $vValue, $iSize = 0

	$oRng = (IsObj($vRange) ? $vRange : $oObj.ActiveSheet.Range($vRange))

	If Not IsKeyword($aFormat) Then
		If Not IsArray($aFormat) Then Return SetError(1, 4, 0)

		$iEnd = UBound($aFormat) - 1

		With $oRng.Borders($iBorder)
			For $i = 0 To $iEnd
				$sKey = StringLower($aFormat[$i][0])
				$vValue = $aFormat[$i][1]

				If $iSize > 12 Then Return SetError(1, 5, 0) ; Thick1 + Thick2 + Seperator = Max 12

				Switch $sKey
					Case "type"
						.Type = $vValue
					Case "thick1"
						.Thick1 = $vValue
						$iSize += $vValue
					Case "thick2"
						.Thick2 = $vValue
						$iSize += $vValue
					Case "seperator"
						.Seperator = $vValue
						$iSize += $vValue
					Case "color"
						.Color = $vValue
					Case "colorindex"
						.ColorIndex = $vValue
				EndSwitch
			Next
		EndWith
	EndIf

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_FormatBorder
; ==============================================================================
; Only simple wrappers
; ------------------------------------------------------------------------------
Func _PlanMaker_FormatBorders(ByRef $oObj, $vRange, $aTop = Default, $aLeft = Default, $aBottom = Default, $aRight = Default, $aHorizontal = Default, $aVertical = Default)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	Local $bScreen = $oObj.Application.ActiveWindow.Workbook.ScreenUpdate
	$oObj.Application.ActiveWindow.Workbook.ScreenUpdate = False

	_PlanMaker_FormatBorder($oObj, $vRange, $_pmBorderTop, $aTop)
	If @error Then Return SetError(1, 10, 0)
	_PlanMaker_FormatBorder($oObj, $vRange, $_pmBorderLeft, $aLeft)
	If @error Then Return SetError(1, 11, 0)
	_PlanMaker_FormatBorder($oObj, $vRange, $_pmBorderBottom, $aBottom)
	If @error Then Return SetError(1, 12, 0)
	_PlanMaker_FormatBorder($oObj, $vRange, $_pmBorderRight, $aRight)
	If @error Then Return SetError(1, 13, 0)
	_PlanMaker_FormatBorder($oObj, $vRange, $_pmBorderHorizontal, $aHorizontal)
	If @error Then Return SetError(1, 14, 0)
	_PlanMaker_FormatBorder($oObj, $vRange, $_pmBorderVertical, $aVertical)
	If @error Then Return SetError(1, 15, 0)

	$oObj.Application.ActiveWindow.Workbook.ScreenUpdate = $bScreen

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_FormatBorders
; ------------------------------------------------------------------------------
Func _PlanMaker_FormatBorders_All(ByRef $oObj, $vRange, $aFormat)
	Return _PlanMaker_FormatBorders($oObj, $vRange, $aFormat, $aFormat, $aFormat, $aFormat, $aFormat, $aFormat)
EndFunc   ;==>_PlanMaker_FormatBorders_All
; ------------------------------------------------------------------------------
Func _PlanMaker_FormatBorders_Frame(ByRef $oObj, $vRange, $aFormat)
	Return _PlanMaker_FormatBorders($oObj, $vRange, $aFormat, $aFormat, $aFormat, $aFormat)
EndFunc   ;==>_PlanMaker_FormatBorders_Frame
; ------------------------------------------------------------------------------
Func _PlanMaker_FormatBorders_Inner(ByRef $oObj, $vRange, $aFormat)
	Return _PlanMaker_FormatBorders($oObj, $vRange, Default, Default, Default, Default, $aFormat, $aFormat)
EndFunc   ;==>_PlanMaker_FormatBorders_Inner

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_FormatFont
; Description ...: Formats the font in a range
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_FormatFont(ByRef $oObj, $sRange[, $aFormat = Default])
; Parameter(s): .: $oObj        -
;                  $vRange      -
;                  $aFormat     - Optional: (Default = Default) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Jun 18 01:40:12 CEST 2017 @27 /Internet-Zeit/
; Version .......: 0.5
; Remark(s) .....:
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_FormatFont(ByRef $oObj, $vRange, $aFormat = Default)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)
	If Not __IsRange($vRange) Then Return SetError(1, 2, 0)

	Local $oRng, $iEnd, $sKey, $vValue

	$oRng = (IsObj($vRange) ? $vRange : $oObj.ActiveSheet.Range($vRange))

	Local $bMan = $oObj.ManualApply
	$oObj.ManualApply = True

	If Not IsKeyword($aFormat) Then
		If Not IsArray($aFormat) Then Return SetError(1, 3, 0)

		$iEnd = UBound($aFormat) - 1

		With $oRng.Font
			For $i = 0 To $iEnd
				$sKey = StringLower($aFormat[$i][0])
				$vValue = $aFormat[$i][1]
				Switch $sKey
					Case "name"
						.Name = $vValue
					Case "size"
						.Size = $vValue
					Case "bold"
						.Bold = $vValue
					Case "italic"
						.Italic = $vValue
					Case "underline"
						.Underline = $vValue
					Case "superscript"
						.Superscript = $vValue
					Case "subscript"
						.Subscript = $vValue
					Case "allcaps"
						.AllCaps = $vValue
					Case "smallcaps"
						.SmallCaps = $vValue
					Case "preferredsmallcaps"
						.PreferredSmallCaps = $vValue
					Case "blink"
						.Blink = $vValue
					Case "color"
						.Color = $vValue
					Case "colorindex"
						.ColorIndex = $vValue
					Case "bcolor"
						.BColor = $vValue
					Case "bcolorindex"
						.BColorIndex = $vValue
					Case "spacing"
						.Spacing = $vValue
					Case "pitch"
						.Pitch = $vValue
				EndSwitch
			Next
		EndWith
	Else
		;$oRng.ClearFormats !!! alle anderen Formatierung werden auch geloescht
	EndIf

	$oRng.ApplyFormatting

	$oObj.ManualApply = $bMan

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_FormatFont

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_FormatShading
; Description ...: Sets the shading of a range / cell
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_FormatShading(ByRef $oObj, $sRange[, $aFormat = Default])
; Parameter(s): .: $oObj        -
;                  $vRange      - Range-object or String
;                  $aFormat     - Optional: (Default = Default) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Jun 18 20:10:37 CEST 2017
; Version .......: 0.3
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_FormatShading(ByRef $oObj, $vRange, $aFormat = Default)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)
	If Not __IsRange($vRange) Then Return SetError(1, 2, 0)

	Local $oRng, $iEnd, $sKey, $vValue

	$oRng = (IsObj($vRange) ? $vRange : $oObj.ActiveSheet.Range($vRange))

	If Not IsKeyword($aFormat) Then
		If Not IsArray($aFormat) Then Return SetError(1, 3, 0)

		$iEnd = UBound($aFormat) - 1

		With $oRng.Shading
			For $i = 0 To $iEnd
				$sKey = StringLower($aFormat[$i][0])
				$vValue = $aFormat[$i][1]
				Switch $sKey
					Case "texture"
						If $vValue < 0 Or $vValue > 13 Then Return SetError(1, 10, 0)
						.Texture = $vValue
					Case "intensity"
						If .Texture = $_smoPatternHalftone Then
							If $vValue < 0 Or $vValue > 100 Then Return SetError(1, 11, 0)
							.Intensity = $vValue ; 0 - 100%
						Else
							Return SetError(1, 12, 0)
						EndIf
					Case "foregroundpatterncolor"
						.ForegroundPatternColor = $vValue
					Case "foregroundpatterncolorindex"
						.ForegroundPatternColorIndex = $vValue
					Case "backgroundpatterncolor"
						.BackgroundPatternColor = $vValue
					Case "backgroundpatterncolorindex"
						.BackgroundPatternColorIndex = $vValue
				EndSwitch
			Next
		EndWith
	Else
		$oRng.Shading.Texture = $_smoPatternNone
	EndIf

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_FormatShading


Func _PlanMaker_FormatNumber(ByRef $oObj, $vRange, $aFormat = Default)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)
	If Not __IsRange($vRange) Then Return SetError(1, 2, 0)

	Local $oRng, $iEnd, $sKey, $vValue

	$oRng = (IsObj($vRange) ? $vRange : $oObj.ActiveSheet.Range($vRange))

	Local $bMan = $oObj.ManualApply
	$oObj.ManualApply = True

	If Not IsKeyword($aFormat) Then
		If Not IsArray($aFormat) Then Return SetError(1, 3, 0)

		$iEnd = UBound($aFormat) - 1

		With $oRng.NumberFormatting
			For $i = 0 To $iEnd
				$sKey = StringLower($aFormat[$i][0])
				$vValue = $aFormat[$i][1]
				Switch $sKey
					Case "type"
						.type = $vValue
					Case "dateformat"
						.DateFormat = $vValue ; abh�ngig von der Sprache der Benutzeroberfl�che!
					Case "customformat"
						.CustomFormat = $vValue ; siehe PlanMaker Handbuch "Aufbau eines benutzerdefinierten Zahlenformats"
					Case "currency"
						.Currency = $vValue ; http://en.wikipedia.org/wiki/ISO_4217
					Case "accounting"
						.Accounting = $vValue ; http://en.wikipedia.org/wiki/ISO_4217
					Case "digits"
						.Digits = $vValue
					Case "negativered"
						.NegativeRed = $vValue
					Case "suppressminus"
						.SuppressMinus = $vValue
					Case "suppresszeros"
						.SuppressZeros = $vValue
					Case "thousandsseparator"
						.ThousandsSeparator = $vValue
				EndSwitch
			Next
		EndWith
	Else
		;$oRng.ClearFormats !!! alle anderen Formatierung werden auch geloescht
	EndIf

	$oRng.ApplyFormatting

	$oObj.ManualApply = $bMan

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_FormatNumber

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_FormulaRead
; Description ...: Reads the formula of a cell
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_FormulaRead(ByRef $oObj[, $iRow = 1[, $iCol = 1]])
; Parameter(s): .: $oPM         -
;                  $iRow        - Optional: (Default = 1) :
;                  $iCol        - Optional: (Default = 1) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Mar 28 14:57:15 CEST 2017
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_FormulaRead(ByRef $oObj, $iRow = 1, $iCol = 1)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, "")

	Local $sValue = $oObj.ActiveSheet.Cells.Item($iRow, $iCol).Formula

	Return SetError(0, 0, $sValue)
EndFunc   ;==>_PlanMaker_FormulaRead

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_FormulaWrite
; Description ...: Writes the formula of a cell
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_FormulaWrite(ByRef $oObj, $vValue[, $iRow = 1[, $iCol = 1]])
; Parameter(s): .: $oPM         -
;                  $vValue      -
;                  $iRow        - Optional: (Default = 1) :
;                  $iCol        - Optional: (Default = 1) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 27 21:03:24 CEST 2017
; Version .......: 0.1
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_FormulaWrite(ByRef $oObj, $vValue, $iRow = 1, $iCol = 1)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)

	$oObj.ActiveSheet.Cells.Item($iRow, $iCol).Formula = $vValue

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_FormulaWrite


Func _PlanMaker_PageSetup(ByRef $oObj, $aFormat = Default, $bCentimeters = True)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	Local $iEnd, $sKey, $vValue

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)

	If Not IsKeyword($aFormat) Then
		If Not IsArray($aFormat) Then Return SetError(1, 3, 0)

		$iEnd = UBound($aFormat) - 1

		With $oObj.ActiveSheet.PageSetUp
			For $i = 0 To $iEnd
				$sKey = StringLower($aFormat[$i][0])
				$vValue = $aFormat[$i][1]
				Switch $sKey
					Case "leftmargin"
						.LeftMargin = ($bCentimeters ? $oObj.CentimetersToPoints($vValue) : $vValue)
					Case "rightmargin"
						.RightMargin = ($bCentimeters ? $oObj.CentimetersToPoints($vValue) : $vValue)
					Case "topmargin"
						.TopMargin = ($bCentimeters ? $oObj.CentimetersToPoints($vValue) : $vValue)
					Case "bottommargin"
						.BottomMargin = ($bCentimeters ? $oObj.CentimetersToPoints($vValue) : $vValue)
					Case "headermargin"
						.HeaderMargin = ($bCentimeters ? $oObj.CentimetersToPoints($vValue) : $vValue)
					Case "footermargin"
						.FooterMargin = ($bCentimeters ? $oObj.CentimetersToPoints($vValue) : $vValue)
					Case "pageheight"
						.PageHeight = ($bCentimeters ? $oObj.CentimetersToPoints($vValue) : $vValue)
					Case "pagewidth"
						.PageWidth = ($bCentimeters ? $oObj.CentimetersToPoints($vValue) : $vValue)
					Case "orientation"
						.Orientation = $vValue
					Case "papersize"
						.PaperSize = $vValue
					Case "printcomments"
						.PrintComments = $vValue
					Case "centerhorizontally"
						.CenterHorizontally = $vValue
					Case "centerverticaly"
						.CenterVerticaly = $vValue
					Case "zoom"
						.Zoom = $vValue
					Case "firstpagenumber"
						.FirstPageNumber = $vValue
					Case "printgridlines"
						.PrintGridLines = $vValue
					Case "printheadings"
						.PrintHeadings = $vValue
					Case "order"
						.Order = $vValue
					Case "printarea"
						.PrintArea = $vValue
					Case "printtitlerows"
						.PrintTitleRows = $vValue
					Case "printtitlecolumns"
						.PrintTitleColumns = $vValue
				EndSwitch
			Next
		EndWith
	Else
		;$oRng.ClearFormats !!! alle anderen Formatierung werden auch geloescht
	EndIf

EndFunc   ;==>_PlanMaker_PageSetup

; Funktioniert nicht? Nur graues "Blatt"
Func _PlanMaker_Print(ByRef $oObj, $iFrom = Default, $iTo = Default)
	If IsKeyWord($iFrom) And IsKeyWord($iTo) Then
		$oObj.PrintOut
	ElseIf IsInt($iFrom) And IsInt($iTo) Then
		$oObj.PrintOut( $iFrom, $iTo )
	ElseIf IsInt($iFrom) And IsKeyWord($iTo) Then
		$oObj.PrintOut( $iFrom )
	ElseIf IsKeyWord($iFrom) And IsInt($iTo) Then
		$oObj.PrintOut( 1, $iTo )
	EndIf
EndFunc

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_Quit
; Description ...: Quits PlanMaker
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_Quit(ByRef $oPM)
; Parameter(s): .: $oPM         -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Thu Mar 23 16:11:30 CET 2017
; Version .......: 0.1
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_Quit(ByRef $oPM)

	If Not IsObj($oPM) Then Return SetError(1, 1, 0)

	$oPM.Application.Quit

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_Quit

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_ScreenUpdate
; Description ...: Toggles or sets the ScreenUpdate
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_ScreenUpdate(ByRef $oPM[, $bUpdate = Default])
; Parameter(s): .: $oPM         -
;                  $bUpdate     - Optional: (Default = Default) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Jun 19 00:45:28 CEST 2017 @989 /Internet-Zeit/
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_ScreenUpdate(ByRef $oPM, $bUpdate = Default)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oPM) Then Return SetError(1, 1, 0)

	With $oPM.Application.ActiveWindow.Workbook
		If IsKeyword($bUpdate) = 1 Then
			If .ScreenUpdate = True Then
				.ScreenUpdate = False
			Else
				.ScreenUpdate = True
			EndIf
		Else
			.ScreenUpdate = $bUpdate
		EndIf
	EndWith

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_ScreenUpdate

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_SheetActivate
; Description ...: Actibates a sheet
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_SheetActivate(ByRef $oPM[, $vSheet = 1])
; Parameter(s): .: $oPM         -
;                  $vSheet      - Optional: (Default = 1) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Thu Mar 23 16:15:25 CET 2017
; Version .......: 0.1
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_SheetActivate(ByRef $oObj, $vSheet = 1)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)

	$oObj.Application.ActiveWorkbook.Sheets($vSheet).Activate

	Return SetError(0, 0, $oObj.Application.ActiveSheet)
EndFunc   ;==>_PlanMaker_SheetActivate

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_SheetAddNew
; Description ...: Adds a new sheet to the current workbook
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_SheetAddNew($oObj[, $sSheet = ""])
; Parameter(s): .: $oObj        -
;                  $sSheet      - Optional: (Default = "") :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Mar 28 14:59:41 CEST 2017
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_SheetAddNew($oObj, $sSheet = "")
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)

	Local $oSheet = $oObj.Application.ActiveWorkbook.Sheets.Add
	If $sSheet <> "" Then $oSheet.Name = $sSheet

	Return SetError(0, 0, $oSheet)
EndFunc   ;==>_PlanMaker_SheetAddNew

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_SheetDelete
; Description ...: Delets a sheet from the current workbook
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_SheetDelete(ByRef $oPM[, $vSheet = 1])
; Parameter(s): .: $oPM         -
;                  $vSheet      - Optional: (Default = 1) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 27 01:10:49 CEST 2017
; Version .......: 0.1
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_SheetDelete(ByRef $oObj, $vSheet = 1)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)

	$oObj.Application.ActiveWorkbook.Sheets($vSheet).Delete

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_SheetDelete


; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_SheetFromArray
; Description ...: Writes an array to a range
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_SheetFromArray(ByRef $oObj, ByRef $aArray, $vRange)
; Parameter(s): .: $oObj        -
;                  $aArray      -
;                  $vRange      -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Jun 18 21:42:51 CEST 2017 @863 /Internet-Zeit/
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_SheetFromArray(ByRef $oObj, ByRef $aArray, $vRange)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)
	If Not __IsRange($vRange) Then Return SetError(1, 2, 0)

	Local $oRng = (IsObj($vRange) ? $vRange : $oObj.ActiveSheet.Range($vRange))

	Local $iRows = UBound($aArray, 1) - 1
	Local $iCols = UBound($aArray, 2) - 1

	Local $bUpdate = $oObj.Application.ActiveWindow.Workbook.ScreenUpdate
	$oObj.Application.ActiveWindow.Workbook.ScreenUpdate = False

	For $i = 0 To $iRows
		For $j = 0 To $iCols
			$oRng.Cells.Item($i + 1, $j + 1).Value = $aArray[$i][$j]
		Next ; col
	Next ; row

	$oObj.Application.ActiveWindow.Workbook.ScreenUpdate = $bUpdate

	Return SetError(0, 0, 1)
EndFunc   ;==>_PlanMaker_SheetFromArray

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_SheetList
; Description ...: Returns an array of all sheets of the current workbook
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_SheetList(ByRef $oObj)
; Parameter(s): .: $oPM         -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Mar 28 11:39:24 CEST 2017
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_SheetList(ByRef $oObj)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)

	Local $aList[1]
	Local $iCount = $oObj.ActiveWorkbook.Sheets.Count
	$aList[0] = $iCount
	ReDim $aList[$iCount + 1]

	For $i = 1 To $iCount
		$aList[$i] = $oObj.Application.ActiveWorkbook.Sheets.Item($i).Name
	Next

	Return SetError(0, 0, $aList)
EndFunc   ;==>_PlanMaker_SheetList

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_SheetToArray
; Description ...: Returns an array of all values in a range
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_SheetToArray(ByRef $oObj[, $vRangeColumnStart = 1[, $iRowStart = 1[, $iColumnEnd = 1[, $iRowEnd = 1]]]])
; Parameter(s): .: $oObj        -
;                  $vRangeColumnStart - Optional: (Default = -1) :
;                  $iRowStart   - Optional: (Default = -1) :
;                  $iColumnEnd  - Optional: (Default = -1) :
;                  $iRowEnd     - Optional: (Default = -1) :
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Thu Apr 20 16:10:15 CEST 2017
; Version .......: 0.3
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_SheetToArray(ByRef $oObj, $vRangeColumnStart = 1, $iRowStart = 1, $iColumnEnd = 1, $iRowEnd = 1)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oObj) Then Return SetError(1, 1, 0)

	Local $aArray[1][1]

	If IsInt($vRangeColumnStart) And IsInt($iRowStart) And IsInt($iColumnEnd) And IsInt($iRowEnd) Then
		$vRangeColumnStart = _PlanMaker_CellRangeByPosition($vRangeColumnStart, $iRowStart, $iColumnEnd, $iRowEnd)
	EndIf

	;ConsoleWrite($vRangeColumnStart & @CRLF)

	Local $oRng = $oObj.ActiveSheet.Range($vRangeColumnStart)
	Local $iRows = $oRng.Rows.Count
	Local $iCols = $oRng.Columns.Count

	ReDim $aArray[$iRows + 1][$iCols + 1]
	$aArray[0][0] = $iRows

	For $i = 1 To $iRows
		For $j = 0 To $iCols
			$aArray[$i][$j] = $oRng.Cells.Item($i, $j + 1).Value
		Next ; col
	Next ; row

	Return SetError(0, 0, $aArray)
EndFunc   ;==>_PlanMaker_SheetToArray

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_UserPropertyGet
; Description ...: Returns an userproperty
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_UserPropertyGet(ByRef $oPM, $iProperty)
; Parameter(s): .: $oPM         -
;                  $iProperty   -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 27 00:12:47 CEST 2017
; Version .......: 0.1
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_UserPropertyGet(ByRef $oPM, $iProperty)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oPM) Then Return SetError(1, 1, 0)
	If $iProperty < 1 Or $iProperty > 24 Then Return SetError(1, 2, 0)

	Return SetError(0, 0, $oPM.Application.UserProperties.Item($iProperty).Value)
EndFunc   ;==>_PlanMaker_UserPropertyGet

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_UserPropertyGetAll
; Description ...: Returns an array of all userproperties
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_UserPropertyGetAll(ByRef $oPM)
; Parameter(s): .: $oPM         -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 27 00:21:18 CEST 2017
; Version .......: 0.1
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_UserPropertyGetAll(ByRef $oPM)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oPM) Then Return SetError(1, 1, 0)

	Local $aUserProperty[24] = [23]

	For $i = 1 To 23
		$aUserProperty[$i] = $oPM.Application.UserProperties.Item($i).Value
	Next

	Return SetError(0, 0, $aUserProperty)
EndFunc   ;==>_PlanMaker_UserPropertyGetAll

; #FUNCTION# ===================================================================
; Name ..........: _PlanMaker_UserPropertySet
; Description ...: Sets an userproperty
; AutoIt Version : V3.3.14.2
; Syntax ........: _PlanMaker_UserPropertySet(ByRef $oPM, $iProperty, $sValue)
; Parameter(s): .: $oPM         -
;                  $iProperty   -
;                  $sValue      -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Mar 27 00:12:26 CEST 2017
; Version .......: 0.1
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func _PlanMaker_UserPropertySet(ByRef $oPM, $iProperty, $sValue)
	Local $oError = ObjEvent("AutoIt.Error", "__PlanMaker_COMErrFunc")
	#forceref $oError

	If Not IsObj($oPM) Then Return SetError(1, 1, 0)
	If $iProperty < 0 Or $iProperty > 24 Then Return SetError(1, 2, 0)
	If $oPM.Application.UserProperties.Item($iProperty).Valid = False Then Return SetError(1, 3, 0)

	$oPM.Application.UserProperties.Item($iProperty).Value = $sValue

	Return SetError(0, 0, $oPM.Application.UserProperties.Item($iProperty).Value)
EndFunc   ;==>_PlanMaker_UserPropertySet

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __ConvertToLetter
; Description ...: Converts a col-number into letter-format
; AutoIt Version : V3.3.14.2
; Syntax ........: __ConvertToLetter($iCol)
; Parameter(s): .: $iCol        -
; Return Value ..: Success      -
;                  Failure      -
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Tue Mar 28 09:11:09 CEST 2017
; Version .......: 0.1
; Link ..........: https://support.microsoft.com/de-de/help/833402/how-to-convert-excel-column-numbers-into-alphabetical-characters
; Related .......:
; Example .......: Yes
; ==============================================================================
Func __ConvertToLetter($iCol)
	Local $iAlpha = Int($iCol / 26)
	Local $iRemainder = $iCol - $iAlpha * 26
	If $iAlpha > 0 Then
		Return (Chr($iAlpha + 64) & Chr(64 + $iRemainder))
	Else
		Return (Chr($iRemainder + 64))
	EndIf
EndFunc   ;==>__ConvertToLetter

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __IsRange
; Description ...: Checks if a range is valid
; AutoIt Version : V3.3.14.2
; Syntax ........: __IsRange($vRange)
; Parameter(s): .: $vRange      - Object or string
; Return Value ..: Success      - True
;                  Failure      - False
;                  @ERROR       -
;                  @EXTENDED    -
; Author(s) .....: Thorsten Willert
; Date ..........: Mon Jun 19 00:40:14 CEST 2017 @986 /Internet-Zeit/
; Version .......: 0.2
; Link ..........:
; Related .......:
; Example .......: Yes
; ==============================================================================
Func __IsRange($vRange)
#cs
Object-name "IDocRange"
Range("A1:B20") ' Zellen A1 bis B20
Range("A1") ' Nur Zelle A1
Range("A:A") ' Gesamte Spalte A
Range("3:3") ' Gesamte Zeile 3
Range("Sommer") ' Benannter Bereich "Sommer"
#ce
	If ObjName($vRange) = "IDocRange" Or _
			StringRegExp($vRange, '^[A-Z,a-z]+[0-9]+:[A-Z,a-z]+[0-9]+$') Or _
			StringRegExp($vRange, '^[A-Z,a-z]+[0-9]+$') Or _
			StringRegExp($vRange, '^[A-Z,a-z]+:[A-Z,a-z]+$') Or _
			StringRegExp($vRange, '^[0-9]+:[0-9]+$') Or _
			StringRegExp($vRange, '^\w+$') Then Return SetError(0, 0, True)

	Return SetError(0, 0, False)
EndFunc   ;==>__IsRange

; #INTERNAL_USE_ONLY# ==========================================================
; Name ..........: __PlanMaker_COMErrFunc
; Description ...:
; AutoIt Version : V3.3.14.2
; Syntax ........: __PlanMaker_COMErrFunc()
; Parameter(s): .:              -
; Return Value ..: Success      -
;                  Failure      -
; Author(s) .....: Thorsten Willert
; Date ..........: Sun Jan 21 14:00:26 CET 2018 @583 /Internet-Zeit/
; Version .......: 1.0
; ==============================================================================
Func __PlanMaker_COMErrFunc($oError)
	ConsoleWrite( "PlanMaker_UDF COM-error" & @crlf)

		Local Const $aErrDisc[10] = [ _
			"err.number is", _
			"err.windescription", _
			"err.description", _
			"err.windescription", _
			"err.source", _
			"err.helpfile", _
			"err.helpcontext", _
			"err.lastdllerror", _
			"err.scriptline", _
			"err.retcode"]
	Local $aErr[10]
	 $aErr[0] = $oError.number
	 $aErr[1] = $oError.windescription
	 $aErr[2] = $oError.description
	 $aErr[3] = $oError.source
	 $aErr[4] = $oError.helpfile
	 $aErr[5] = $oError.helpcontext
	 $aErr[6] = $oError.lastdllerror
	 $aErr[7] = $oError.scriptline
	 $aErr[8] = $oError.retcode

	For $i = 0 To 9
		If StringStripWS($aErrDisc[$i], 7) <> "" Then ConsoleWrite($aErrDisc[$i] & ": " & $aErr[$i] & @CRLF)
	Next
	; Do nothing special, just check @error after suspect functions.
EndFunc   ;==>__PlanMaker_COMErrFunc
