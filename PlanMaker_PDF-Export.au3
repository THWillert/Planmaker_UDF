; A simple  hack for the PDF-export.

; For other languages then german you have to change the
; parameters $sPDFExportTitle and $sSaveTitle

; Works only with an open Workbook

; !!! Doesn't work with Ribbons, only with menus !!!

__PlanMaker_SafePDF("test.pmdx", "Test.pdf")

Func __PlanMaker_SafePDF($sTitle, $sFile, $sPDFExportTitle = "PDF-Export", $sSaveTitle = "Speichern unter best�tigen")
    BlockInput(True)

    Local $iOpt = Opt("WinTitleMatchMode")
    Opt("WinTitleMatchMode", -1)

    ; PlanMaker
    Local $hHnd = WinActivate("[REGEXPTITLE:.*?" & $sTitle & ".*?; CLASS:pmwMdiFrame]", "")
    If $hHnd <> 0 Then
        Send("!D") ; Men� aufrufen
        Send("{DOWN 15}") ; Speichern als PDF
        Send("{ENTER}")

        $hHnd = WinWait("[TITLE:" & $sPDFExportTitle & "; CLASS:SMDIALOG]", "", 2) ; PDF-Export dialog
        If $hHnd <> 0 Then
            Send("{ENTER}") ; export best�tigen

            $hHnd = WinWait($sPDFExportTitle, "", 2) ; Datei-Dialog
            If $hHnd <> 0 Then
                Send($sFile) ; Dateinamen senden
                ControlClick($sPDFExportTitle, "", 1) ; best�tigen
            EndIf

            $hHnd = WinWait("[TITLE:PlanMaker]", "", 2) ; �berschreiben
            If @error = 0 Then
                Send("{ENTER}") ; und best�tigen
            EndIf
        EndIf
    EndIf

    Opt("WinTitleMatchMode", $iOpt)
    BlockInput(False)
EndFunc   ;==>__PlanMaker_SafePDF
