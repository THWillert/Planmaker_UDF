# Planmaker_UDF

Eine [AutoIt](https://www.autoitscript.com/site/autoit/)-UDF für die Tabellenkalkulation [PlanMaker](https://www.softmaker.de/softmaker-office-planmaker) der Firma [SoftMaker](https://www.softmaker.de/softmaker-office).

## Übersicht der Funktionen

| Funktion | Beschreibung |
|--------|--------|
| _PlanMaker_BookAttach| Script mit geöffnetem PlanMaker verbinden|
| _PlanMaker_BookClose| PlanMaker Workbook schließen|
| _PlanMaker_BookNew|Neues Workbook erstellen|
| _PlanMaker_BookOpen|Workbook öffnen|
| _PlanMaker_BookSave|Workbook speichern|
|_PlanMaker_BookSaveAs|Workbook speichern unter ...|
|_PlanMaker_CellRangeByPosition|Gibt einen Bereich anhand der Positionen zurück|
|_PlanMaker_CellRead |Ließt den Inhalt einer Zelle|
| _PlanMaker_CellWrite |Schreibt den Inhalt einer Zelle|
| _PlanMaker_DocumentPropertyGet |Gibt eine bestimmte Eigenschaft dea aktuellen Dokuments zurück|
| _PlanMaker_DocumentPropertyGetAll |Gibt alle Eigenschaften des aktuellen Dokuments als Array zurück|
| _PlanMaker_DocumentPropertySet |Setzt eine Eigenschaft des aktuellen Dokuments|
| _PlanMaker_FormatBorder |Formatiert einen Rand|
| _PlanMaker_FormatBorders |Formatiert alle Ränder|
| _PlanMaker_FormatBorders_All |Formatiert alle Ränder eines Bereichs|
| _PlanMaker_FormatBorders_Frame |Formatiert den Außenrand eines Bereichs|
| _PlanMaker_FormatBorders_Inner |Formatiert die inneren Ränder eines Bereichs|
| _PlanMaker_FormatFont |Schriftformatierung eines Bereichs|
| _PlanMaker_FormatNumber |Nummer-Formatierung eines Bereichs|
| _PlanMaker_FormatShading |Hintergrund eines Bereichs|
| _PlanMaker_FormulaRead |Ließt die Formel einer Zelle|
| _PlanMaker_FormulaWrite |Schreibt die Formel einer Zelle|
| _PlanMaker_PageSetup |Seiten Formatierung|
| _PlanMaker_Print |Druckt die aktuelle Tabelle|
| _PlanMaker_Quit |Beendet Planmaker|
| _PlanMaker_Color2SmoColor |Wandelt RGB und HEX Farbwerte in SoftMaker-Office BGR Farben um|
| _PlanMaker_ScreenUpdate |Schaltet die Aktualisierung der Anzeige ein/aus|
| _PlanMaker_SheetActivate |Aktiviert ein Arbeitsblatt|
| _PlanMaker_SheetAddNew |Erstellt ein neues Arbeitsblatt|
| _PlanMaker_SheetDelete |Löscht das aktuelle Arbeitsblatt|
| _PlanMaker_SheetList |Gibt ein Array mit einer Liste aller Arbeitsblätter zurück|
| _PlanMaker_SheetFromArray |Schreibt alle Werte eines 2D-Arrays in ein Arbeitsblatt|
| _PlanMaker_SheetToArray |Liest alle Werte eines Arbeitsblattes in ein 2D-Array|
| _PlanMaker_UserPropertyGet |Gibt eine Benutzereigenschaft zurück|
| _PlanMaker_UserPropertyGetAll |Gibt ein Array mit allen Benutzereigenschaften zurück|
| _PlanMaker_UserPropertySet |Setzt eine Benutzereigenschaft|

### Voraussetzungen

Kompatibel mit SoftMaker Office 2018 und FreeOffice 2018.


### Installation

Die UDF in das Include Verzeichnis von AutoIt kopieren.


### Weiterführende Informationen

[Handbuch für BasicMaker](http://www.softmaker.net/down/bm2018manual_de.pdf)

### Diskusion und Vorschläge

[SoftMaker-Forum](https://forum.softmaker.de/viewtopic.php?f=230&t=22648)

[AutoIt.de](https://autoit.de/thread/85113-planmaker-udf-tabellenkalkulation-von-softmaker-office-und-freeoffice/)


## ToDo

- [ ] Konstanten für Datentypen
- [ ] Konstanten für Währung
- [ ] Beschreibung der Rückgabeparameter


## Author
Thorsten Willert

[Homepage](http://www.thorsten-willert.de/)

## Lizenz
Das ganze steht unter der [Apache 2.0](https://github.com/THWillert/HomeMatic_CSS/blob/master/LICENSE) Lizenz
.
