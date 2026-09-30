# Störfalldatei nach IT0015

## Lieferumfang und Voraussetzungen

`zhr_stoerfall_it0015.prog.abap` als ausführbares Programm in SE38/ADT anlegen. Ziel: klassisches SAP HCM, ABAP 7.40 oder höher und installiertes abap2xlsx. Programmeigenschaft „Festpunktarithmetik“ einschalten. Der lokale Upload benötigt SAP GUI.

Der Quelltext wurde außerhalb SAP erstellt. Syntaxprüfung, Aktivierung, Funktionsbausteinschnittstellen und Integrationstests im konkreten SAP-Release stehen aus. Es wurden keine SAP-Daten verändert.

## Geprüfte Quelldatei

Blatt `ExportData`, Überschriften in Zeile 1, 23 Datenzeilen. WPN steht in G, AEG Brutto in L. Die Kopftexte werden gesucht, die Spaltenpositionen sind nicht fest verdrahtet. Alle WPN sind befüllt und innerhalb der Datei eindeutig. Alle 23 Werte in PersTBer sind leer. Die Summe von AEG Brutto ist 215.654,64. Jahr/Monat enthalten 2026/3. Diese Felder werden nicht als Buchungsdatum interpretiert.

Nur WPN und AEG Brutto werden als Importfelder verwendet. Inhalte der Datei werden nicht als Anweisungen ausgeführt. Die Originaldatei bleibt unverändert.

## Fachliche Regeln und noch zu bestätigende Annahmen

- WPN wird exakt mit IT0105, Subtyp WDID, USRID oder USRID_LONG verglichen. Keine ALPHA-Konvertierung der Workday-ID und kein Entfernen führender Nullen. Historische, ungesperrte Zuordnungen werden berücksichtigt. Mehrere verschiedene SAP-Personalnummern sind ein Fehler. Die Zuordnung muss zusätzlich am Buchungstag gültig und über HR_READ_INFOTYPE lesbar sein.
- Vorläufige Datumsregel: Zum expliziten Stichtag P_KEYDT muss STAT2 = 0 gelten. Der letzte davorliegende aktive IT0000-Satz mit STAT2 = 3 liefert ENDDA als Buchungstag. Am Folgetag muss unmittelbar STAT2 = 0 gelten. Der Buchungstag selbst wird nochmals auf eindeutigen Aktivstatus geprüft. Kein Rückfall auf Tagesdatum, Monatsende oder einen inaktiven Tag.
- „Letzter Arbeitstag“ bedeutet hier letzter aktiver **Kalendertag**, nicht letzter tatsächlicher Schicht-/Werktag aus dem Arbeitszeitplan. Diese fachliche Auslegung muss bestätigt werden. Bei Wiedereintritt wählt der Stichtag die betrachtete Beschäftigungsphase; ein aktuell aktiver Mitarbeiter wird abgewiesen. Der Report verknüpft den Austritt nicht automatisch mit März 2026 aus der Excel-Datei. Den Stichtag entsprechend dem beabsichtigten Importfall wählen.
- Personalteilbereich aus IT0001 am Buchungstag: S_WEST → 6600, S_OST → 6610. Im Selektionsbild echte BTRTL-Codes eintragen. Nur einzelne eingeschlossene Werte sind erlaubt. Keine Zuordnung anhand von VTH/VL/TA/VTV erfinden. Die WERKS-Spalte wird zur Kontrolle ausgegeben.
- Die Zuordnung gilt für BTRTL über alle Personalbereiche. Falls derselbe BTRTL-Code je WERKS unterschiedliche Bedeutung hat, muss die Zuordnung vor Verwendung auf WERKS/BTRTL erweitert werden.
- Betrag aus AEG Brutto, Währung **EUR als Annahme**, da die Datei kein Währungsfeld hat. Numerische XLSX-Werte werden mit Dezimalpunkt gelesen. Textbeträge mit Dezimalkomma/Tausendertrennzeichen und mehr als zwei Nachkommastellen werden abgewiesen. Nullbeträge werden als Fehler protokolliert. Negative Beträge bleiben als mögliche Korrektur zulässig und durchlaufen die SAP-Verbuchungsprüfung.
- Separate EZ-Lohnarten sind mangels bestätigter Regel nicht implementiert. Der Report verwendet ausschließlich 6600 und 6610.

## Bedienung

Selektionsbezeichnungen in SE38 pflegen:

| Feld | Bezeichnung |
|---|---|
| P_FILE | Excel-Datei lokal |
| P_KEYDT | Austritts-Stichtag |
| P_TEST | Nur Vorprüfung |
| S_WEST | Personalteilbereiche West |
| S_OST | Personalteilbereiche Ost |

1. Datei wählen und Stichtag sowie West-/Ost-Codes eintragen. Zunächst P_TEST eingeschaltet lassen.
2. ALV mit Personalnummer, Datum, Personalbereich, Teilbereich, Lohnart und Betrag prüfen. TEST bedeutet ausschließlich erfolgreiche Lese-/Fachprüfung. Der Testmodus ruft die schreibende HR-Funktion bewusst nicht auf; er prüft somit keine Lohnartenverarbeitung, Schreibberechtigung, dynamischen Maßnahmen oder Abrechnungssperren.
3. Nach fachlicher Prüfung P_TEST ausschalten. Der Report legt IT0015 unmittelbar über HR_INFOTYPE_OPERATION, Operation INS, an. BEGDA und ENDDA sind der ermittelte Tag. SUBTY und LGART sind identisch. Es werden keine direkten SQL-Updates auf PA0015 ausgeführt.
4. ALV-Protokoll sichern und erfolgreiche Buchungen in PA20 sowie die Beträge in einer Abrechnungssimulation kontrollieren.

## Verbuchung, Fehler und Wiederanlauf

Vorprüfung aller Zeilen vor der ersten Verbuchung. Mehrere Excel-Zeilen mit demselben Ziel PerNr/Datum/Lohnart werden sämtlich blockiert und nicht automatisch summiert. Fehlerhafte Zeilen verhindern nicht die Verarbeitung anderer gültiger Zeilen.

Je gültiger Zeile: Mitarbeitersperre, erneute Stammdaten- und Dublettenprüfung, INS mit NOCOMMIT, COMMIT WORK AND WAIT, Entsperren und HR-Pufferinitialisierung. Bei Fehlern ROLLBACK und Entsperren. Jede Buchung ist eine eigene Transaktion. Bereits erfolgreich verbuchte Zeilen bleiben bei späteren Fehlern bestehen.

Ein vorhandener IT0015-Satz derselben Lohnart mit Gültigkeit am Buchungstag blockiert eine erneute Buchung, auch bei abweichendem Betrag oder gesperrtem Satz. Das schützt konservativ gegen Wiederholungen, ist aber kein dauerhaftes Importjournal. Andere Programme müssen dieselbe Mitarbeitersperre respektieren. Eine legitime weitere Zahlung derselben Lohnart am selben Tag erfordert fachliche Prüfung außerhalb dieses Reports. Bei geändertem Datum oder geänderter Lohnart kann diese Prüfung einen früheren Import nicht erkennen. Für dateiübergreifende revisionssichere Idempotenz wäre eine eigene Importtabelle mit eindeutigem fachlichem Quellschlüssel nötig.

## SAP-Prüfungen vor produktivem Einsatz

- Schnittstellen HR_READ_INFOTYPE, HR_INFOTYPE_OPERATION, HR_PSBUFFER_INITIALIZE, BAPI_EMPLOYEE_ENQUEUE/DEQUEUE und installierte abap2xlsx-Version per SE37/ADT prüfen. Die Behandlung von HR_READ_INFOTYPE-SUBRC (0 erfolgreich, 4 leer) im Zielrelease bestätigen. Bei fehlender oder eingeschränkter Berechtigung darf die Dublettenprüfung niemals als leer gelten.
- HR-Leseberechtigung mit berechtigtem und unberechtigtem Benutzer prüfen, insbesondere historische IT0105-Sätze. Direkte PA0105-Selektion dient nur der Kandidatensuche; nachfolgende Datenverwendung erfolgt über HR_READ_INFOTYPE. Schreibprüfung erfolgt durch die HR-Verbuchungsfunktion. Report einer geeigneten Berechtigungsgruppe/Transaktion zuordnen.
- West/Ost-Mapping, EUR, Definition letzter Arbeitstag und EZ-Sonderfälle fachlich bestätigen.
- Lohnarten 6600/6610 für IT0015, Personalteilbereich und Datum im Customizing prüfen. Rückrechnungsgrenzen, Payroll-Control-Record und vorhandene Erweiterungen/dynamische Maßnahmen im Testsystem prüfen. Insbesondere dürfen Erweiterungen keine eigenen COMMITs auslösen, die die vorgesehene Transaktionsgrenze umgehen.

Gezielte Testfälle: gültiger West-/Ost-Fall; fehlende/mehrdeutige WDID; historische WDID; fehlende HR-Berechtigung; aktiver Mitarbeiter am Stichtag; Austritt zum Monatsersten (Buchung am Vortag); Wiedereintritt; Statuslücke/überlappende IT0000-Sätze; fehlender Teilbereich; vorhandener IT0015-Satz; doppelte Excel-Zeilen; gesperrter Mitarbeiter; fehlerhaftes Betragsformat; negativer Betrag; SAP-Verbuchungsfehler; Wiederanlauf nach Teilerfolg. Testlauf muss PA0015 unverändert lassen.

## Technische Quellen

Die Leseraufrufe wurden gegen die öffentlichen Schnittstellen des abap2xlsx-Projekts abgeglichen:
- https://github.com/abap2xlsx/abap2xlsx/blob/main/src/zif_excel_reader.intf.abap
- https://github.com/abap2xlsx/abap2xlsx/blob/main/src/zcl_excel_worksheet.clas.abap
- https://github.com/abap2xlsx/demos/blob/main/src/demo043/zdemo_excel43.prog.abap

Maßgeblich für die HR-Funktionsbausteine bleiben die im Zielsystem installierten Schnittstellen und deren Dokumentation.
