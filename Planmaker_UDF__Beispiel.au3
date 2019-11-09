#include "PlanMaker.au3"
#include "Array.au3"
#include "File.au3"

Local $sFile = _TempFile()

; Neue Arbeitsmappe
Local $oTest = _PlanMaker_BookNew(True, True)

; Aktualisierung ausschalten
_PlanMaker_ScreenUpdate($oTest, False)

; Zellen fÜllen
For $i = 1 To 10
	_PlanMaker_CellWrite($oTest, $i, $i, 1)
Next

; Formel einfÜgen
_PlanMaker_FormulaWrite($oTest, "=Summe(A1:A10)", 12, 1)

; Schrift-Formatierung
Local $aFont[][] = [["italic", True]]
_PlanMaker_FormatFont($oTest, "A1:A10", $aFont)

;Local $aFont[][] = [["bold", True], ["underline", $_pmUnderlineSingle], ["colorindex", $_smoColorIndexRed]]
; _PlanMaker_Color2SmoColor("255,0,0")
; _PlanMaker_Color2SmoColor(255,0,0)
; _PlanMaker_Color2SmoColor("FF0000")
; _PlanMaker_Color2SmoColor("FF","00","00")

Local $aFont[][] = [ _
	["bold", True], _
	["underline", $_pmUnderlineSingle], _
	["color", _PlanMaker_Color2SmoColor("ff0000")]]
_PlanMaker_FormatFont($oTest, "A12", $aFont)

Local $aNumber[][] = [ _
	["Type", $_pmNumberDecimal], _
	["Digits", 2], _
	["Currency", "€"] ]
_PlanMaker_FormatNumber($oTest, "A1:A12", $aNumber)

Local $aPage[][] = [ _
	[ "PaperSize", $_smoPaperA5 ], _
	[ "CenterHorizontally", True ], _
	[ "Zoom", 200] ]
_PlanMaker_PageSetup($oTest, $aPage)

;_PlanMaker_Print($oTest, 1, 2) ; !!!Ausdruck auf Standard-Drucker

; Rahmen
Local $aBorder[][] = [["Type", $_pmLineStyleSingle]]
_PlanMaker_FormatBorders_Inner($oTest, "A1:B12", $aBorder)

Local $aBorder[][] = [ _
	["Type", $_pmLineStyleSingle], _
	["Thick1", 1.5], _
	["Color", $_smoColorRed]]
_PlanMaker_FormatBorders_Frame($oTest, "A1:B12", $aBorder)

; Shading
; Range
Local $aShading[][] = [ _
	["Texture", $_smoPatternHalftone], _
	["Intensity", 10]]
_PlanMaker_FormatShading($oTest, "B1", $aShading)

; Range als Object
Local $oRange = $oTest.ActiveSheet.Range("B2")
Local $aShading[][] = [ _
	["Texture", $_smoPatternHashDiagCoarse]]
_PlanMaker_FormatShading($oTest, $oRange, $aShading)

; Range als Positionsangabe
Local $aShading[][] = [ _
	["Texture", $_smoPatternRightDiagFine], _
	["ForegroundPatternColor", _PlanMaker_Color2SmoColor("00AAFF")]]
_PlanMaker_FormatShading($oTest, _PlanMaker_CellRangeByPosition(2, 3, 2, 3), $aShading)

; Aktualisierung einschalten
_PlanMaker_ScreenUpdate($oTest, True)
Sleep(5000)

; Zell-Bereich auslesen Über Positionsangabe
Local $aArray = _PlanMaker_SheetToArray($oTest, 1, 1, 1, 12)
_ArrayDisplay($aArray)

; Zell-Bereich auslesen Über Bereichsangabe
$aArray = _PlanMaker_SheetToArray($oTest, "A1:A12")
_ArrayDisplay($aArray)

; Benutzereigenschaften auslesen
$aArray = _PlanMaker_UserPropertyGetAll($oTest)
_ArrayDisplay($aArray)

; Dokumenteneigenschaften auslesen
$aArray = _PlanMaker_DocumentPropertyGetAll($oTest)
_ArrayDisplay($aArray)

Sleep(5000)

; Neues Tabellenblatt
_PlanMaker_SheetAddNew($oTest, "Tabelle 2")

Local $aArray[][] = [["1,1", "1,2"], ["2,1", "2,2"], ["3,1", "3,2"], ["4,1", "4,2"]]
_PlanMaker_SheetFromArray($oTest, $aArray, "B2")

Sleep(5000)
; Datei speichern und Beenden
If _PlanMaker_BookSaveAs($oTest, $sFile) Then
	_PlanMaker_Quit($oTest)
	FileDelete($sFile)
EndIf
; ==============================================================================

