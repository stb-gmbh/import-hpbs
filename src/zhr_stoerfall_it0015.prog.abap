*&---------------------------------------------------------------------*
*& Report ZHR_STOERFALL_IT0015
*& Import einer Stoerfalldatei nach IT0015 mit abap2xlsx
*&---------------------------------------------------------------------*
REPORT zhr_stoerfall_it0015.

*=======================================================================
* Datendeklarationen
*=======================================================================
TABLES: pa0001.

TYPES: BEGIN OF ty_data,
         zeile   TYPE i,
         wpn     TYPE string,
         betrag  TYPE p0015-betrg,
         pernr   TYPE pernr_d,
         datum   TYPE sy-datum,
         werks   TYPE p0001-werks,
         btrtl   TYPE p0001-btrtl,
         lgart   TYPE p0015-lgart,
         waers   TYPE p0015-waers,
         status  TYPE c LENGTH 10,
         meldung TYPE string,
       END OF ty_data.

DATA: t_data      TYPE STANDARD TABLE OF ty_data,
      wa_data     TYPE ty_data,
      l_fehler    TYPE string,
      lo_fehler   TYPE REF TO cx_root.

*=======================================================================
* Selektionsbild
*=======================================================================
PARAMETERS: p_file  TYPE string LOWER CASE OBLIGATORY,
            p_keydt TYPE sy-datum OBLIGATORY DEFAULT sy-datum,
            p_test  AS CHECKBOX DEFAULT 'X'.

SELECT-OPTIONS: s_west FOR pa0001-btrtl NO INTERVALS,
                s_ost  FOR pa0001-btrtl NO INTERVALS.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
*----------------------------------------------------------------------*
  PERFORM datei_auswaehlen.

*----------------------------------------------------------------------*
AT SELECTION-SCREEN.
*----------------------------------------------------------------------*
  PERFORM selektion_pruefen.

*----------------------------------------------------------------------*
START-OF-SELECTION.
*----------------------------------------------------------------------*
  REFRESH t_data.
  CLEAR l_fehler.

  TRY.
      PERFORM excel_einlesen CHANGING l_fehler.
    CATCH cx_root INTO lo_fehler.
      l_fehler = lo_fehler->get_text( ).
  ENDTRY.

  IF l_fehler IS NOT INITIAL.
    MESSAGE l_fehler TYPE 'S' DISPLAY LIKE 'E'.
  ELSE.
    PERFORM daten_pruefen.
    PERFORM dubletten_pruefen.
    PERFORM daten_verarbeiten.
  ENDIF.

*----------------------------------------------------------------------*
END-OF-SELECTION.
*----------------------------------------------------------------------*
  CHECK l_fehler IS INITIAL.
  CHECK t_data IS NOT INITIAL.
  PERFORM alv_ausgabe.

*=======================================================================
* Unterprogramme (keine Includes)
*=======================================================================

*----------------------------------------------------------------------*
* Form datei_auswaehlen
*----------------------------------------------------------------------*
FORM datei_auswaehlen.
  DATA: lt_files TYPE filetable,
        wa_file  TYPE file_table,
        l_rc     TYPE i.

  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      file_filter = 'Excel (*.xlsx)|*.xlsx|'
    CHANGING
      file_table  = lt_files
      rc          = l_rc
    EXCEPTIONS
      OTHERS      = 1.

  IF sy-subrc EQ 0 AND l_rc GT 0.
    READ TABLE lt_files INDEX 1 INTO wa_file.
    IF sy-subrc EQ 0.
      p_file = wa_file-filename.
    ENDIF.
  ENDIF.
ENDFORM.

*----------------------------------------------------------------------*
* Form selektion_pruefen
*----------------------------------------------------------------------*
FORM selektion_pruefen.
  IF sy-batch EQ abap_true.
    MESSAGE 'Dieser Report benoetigt SAP GUI fuer den lokalen XLSX-Upload.' TYPE 'E'.
  ENDIF.
  IF s_west[] IS INITIAL AND s_ost[] IS INITIAL.
    MESSAGE 'Mindestens eine West-/Ost-Zuordnung eingeben.' TYPE 'E'.
  ENDIF.
  LOOP AT s_west.
    IF s_west-sign NE 'I' OR s_west-option NE 'EQ'.
      MESSAGE 'West: nur einzelne eingeschlossene Teilbereiche erlaubt.' TYPE 'E'.
    ENDIF.
    IF s_ost[] IS NOT INITIAL AND s_west-low IN s_ost.
      MESSAGE 'West und Ost duerfen sich nicht ueberschneiden.' TYPE 'E'.
    ENDIF.
  ENDLOOP.
  LOOP AT s_ost.
    IF s_ost-sign NE 'I' OR s_ost-option NE 'EQ'.
      MESSAGE 'Ost: nur einzelne eingeschlossene Teilbereiche erlaubt.' TYPE 'E'.
    ENDIF.
  ENDLOOP.
ENDFORM.

*----------------------------------------------------------------------*
* Form zelle_lesen
*----------------------------------------------------------------------*
FORM zelle_lesen USING    po_blatt TYPE REF TO zcl_excel_worksheet
                         p_spalte TYPE i
                         p_zeile TYPE i
                CHANGING p_wert TYPE string
                         p_fehler TYPE string.
  DATA: l_wert   TYPE zexcel_cell_value,
        l_formel TYPE zexcel_cell_formula,
        lo_error TYPE REF TO cx_root.

  CLEAR: p_wert, p_fehler.
  TRY.
      CALL METHOD po_blatt->get_cell
        EXPORTING
          ip_column = p_spalte
          ip_row    = p_zeile
        IMPORTING
          ep_value   = l_wert
          ep_formula = l_formel.
      IF l_formel IS NOT INITIAL.
        p_fehler = 'Formeln nicht erlaubt; Excel-Werte exportieren.'.
        RETURN.
      ENDIF.
      p_wert = l_wert.
      CONDENSE p_wert.
    CATCH cx_root INTO lo_error.
      p_fehler = lo_error->get_text( ).
  ENDTRY.
ENDFORM.

*----------------------------------------------------------------------*
* Form excel_einlesen
*----------------------------------------------------------------------*
FORM excel_einlesen CHANGING p_fehler TYPE string.
  DATA: lo_reader TYPE REF TO zif_excel_reader,
        lo_excel  TYPE REF TO zcl_excel,
        lo_blatt  TYPE REF TO zcl_excel_worksheet,
        l_titel   TYPE zexcel_sheet_title,
        l_spalten TYPE i,
        l_zeilen  TYPE i,
        l_spalte  TYPE i,
        l_zeile   TYPE i,
        l_wpn_col TYPE i,
        l_bet_col TYPE i,
        l_kopf    TYPE string,
        l_betrag  TYPE string,
        l_meldung TYPE string,
        wa_import TYPE ty_data.

  CLEAR p_fehler.
  CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
  lo_excel = lo_reader->load_file(
    i_filename        = p_file
    i_from_applserver = abap_false ).
  lo_blatt = lo_excel->get_worksheet_by_index( iv_index = 1 ).
  l_titel = lo_blatt->get_title( ).
  IF l_titel NE 'ExportData'.
    p_fehler = 'Erstes Tabellenblatt muss ExportData heissen.'.
    RETURN.
  ENDIF.

  l_spalten = lo_blatt->get_highest_column( ).
  DO l_spalten TIMES.
    l_spalte = sy-index.
    PERFORM zelle_lesen USING lo_blatt l_spalte 1
                       CHANGING l_kopf p_fehler.
    IF p_fehler IS NOT INITIAL.
      RETURN.
    ENDIF.
    CASE l_kopf.
      WHEN 'WPN'.
        IF l_wpn_col NE 0.
          p_fehler = 'WPN-Kopf doppelt.'.
          RETURN.
        ENDIF.
        l_wpn_col = l_spalte.
      WHEN 'AEG Brutto'.
        IF l_bet_col NE 0.
          p_fehler = 'Betragskopf doppelt.'.
          RETURN.
        ENDIF.
        l_bet_col = l_spalte.
    ENDCASE.
  ENDDO.
  IF l_wpn_col EQ 0 OR l_bet_col EQ 0.
    p_fehler = 'Spalten WPN/AEG Brutto fehlen in Zeile 1.'.
    RETURN.
  ENDIF.

  l_zeilen = lo_blatt->get_highest_row( ).
  l_zeilen = l_zeilen - 1.
  DO l_zeilen TIMES.
    l_zeile = sy-index + 1.
    CLEAR: wa_import, l_betrag, l_meldung.
    wa_import-zeile = l_zeile.
    PERFORM zelle_lesen USING lo_blatt l_wpn_col l_zeile
                       CHANGING wa_import-wpn l_meldung.
    IF l_meldung IS INITIAL.
      PERFORM zelle_lesen USING lo_blatt l_bet_col l_zeile
                         CHANGING l_betrag l_meldung.
    ENDIF.
    IF l_meldung IS INITIAL AND
       wa_import-wpn IS INITIAL AND l_betrag IS INITIAL.
      CONTINUE.
    ENDIF.
    IF l_meldung IS INITIAL.
      PERFORM importwerte_pruefen USING l_betrag
                                 CHANGING wa_import l_meldung.
    ENDIF.
    IF l_meldung IS NOT INITIAL.
      wa_import-status = 'FEHLER'.
      wa_import-meldung = l_meldung.
    ENDIF.
    APPEND wa_import TO t_data.
  ENDDO.
ENDFORM.

*----------------------------------------------------------------------*
* Form importwerte_pruefen
*----------------------------------------------------------------------*
FORM importwerte_pruefen USING p_betrag TYPE string
                            CHANGING ps_data TYPE ty_data
                                     p_fehler TYPE string.
  DATA: l_betrag TYPE decfloat34,
        lo_error TYPE REF TO cx_root.

  CLEAR p_fehler.
  FIND REGEX '^[0-9]+$' IN ps_data-wpn.
  IF sy-subrc NE 0.
    p_fehler = 'WPN muss eine Ziffernfolge sein (keine Exponentialzahl).'.
    RETURN.
  ENDIF.
* XLSX liefert numerische Werte mit Dezimalpunkt.
  FIND REGEX '^[+-]?[0-9]+([.][0-9]{1,2})?$' IN p_betrag.
  IF sy-subrc NE 0.
    p_fehler = 'AEG Brutto: Dezimalpunkt und max. 2 Nachkommastellen.'.
    RETURN.
  ENDIF.
  TRY.
      l_betrag = p_betrag.
      ps_data-betrag = l_betrag.
    CATCH cx_root INTO lo_error.
      p_fehler = lo_error->get_text( ).
      RETURN.
  ENDTRY.
  IF ps_data-betrag EQ 0.
    p_fehler = 'Nullbetrag nicht gebucht.'.
  ENDIF.
ENDFORM.

*----------------------------------------------------------------------*
* Form infotyp_lesen
*----------------------------------------------------------------------*
FORM infotyp_lesen USING VALUE(p_pernr) TYPE pernr_d
                           VALUE(p_infty) TYPE infty
                           VALUE(p_begda) TYPE sy-datum
                           VALUE(p_endda) TYPE sy-datum
                  CHANGING pt_daten TYPE STANDARD TABLE
                           p_fehler TYPE string.
  DATA l_rc TYPE sy-subrc.

  REFRESH pt_daten.
  CLEAR p_fehler.
  CALL FUNCTION 'HR_READ_INFOTYPE'
    EXPORTING
      pernr         = p_pernr
      infty         = p_infty
      begda         = p_begda
      endda         = p_endda
      bypass_buffer = 'X'
    IMPORTING
      subrc         = l_rc
    TABLES
      infty_tab     = pt_daten
    EXCEPTIONS
      infty_not_found = 1
      OTHERS          = 2.
  IF sy-subrc NE 0 OR l_rc NE 0.
    CONCATENATE 'IT' p_infty
                ': Lesen fehlgeschlagen/keine Berechtigung.'
           INTO p_fehler.
  ENDIF.
ENDFORM.

*----------------------------------------------------------------------*
* Form stammdaten_pruefen
*----------------------------------------------------------------------*
FORM stammdaten_pruefen CHANGING ps_data TYPE ty_data
                                  p_fehler TYPE string.
  DATA: lt_pernr TYPE SORTED TABLE OF pernr_d WITH UNIQUE KEY table_line,
        l_wdid TYPE pa0105-usrid,
        l_wdid_long TYPE pa0105-usrid_long,
        lt_p0105 TYPE STANDARD TABLE OF p0105,
        lt_p0000 TYPE STANDARD TABLE OF p0000,
        lt_p0001 TYPE STANDARD TABLE OF p0001,
        lt_p0015 TYPE STANDARD TABLE OF p0015,
        l_folgetag TYPE sy-datum,
        l_anzahl TYPE i,
        l_treffer TYPE i.
  DATA: wa_p0000 TYPE p0000,
        wa_p0001 TYPE p0001,
        l_rc TYPE sy-subrc.

  CLEAR p_fehler.
  CLEAR: ps_data-pernr, ps_data-datum, ps_data-werks, ps_data-btrtl,
         ps_data-lgart, ps_data-waers.
  IF ps_data-wpn IS INITIAL OR strlen( ps_data-wpn ) GT 30.
    p_fehler = 'WPN fehlt/ist zu lang.'.
    RETURN.
  ENDIF.
  l_wdid = ps_data-wpn.
  l_wdid_long = ps_data-wpn.
* SQL nur Kandidatensuche. Keine Stammdatenverwendung ohne HR-Lesepruefung.
* Historische WDID-Saetze einschliessen, da ausgeschiedene Mitarbeiter.
  SELECT DISTINCT pernr FROM pa0105 INTO TABLE lt_pernr
    WHERE subty EQ 'WDID' AND sprps EQ space
      AND ( usrid EQ l_wdid OR usrid_long EQ l_wdid_long ).
  IF lines( lt_pernr ) NE 1.
    p_fehler = 'WDID fehlt oder ist mehreren Personalnummern zugeordnet.'.
    RETURN.
  ENDIF.
  READ TABLE lt_pernr INDEX 1 INTO ps_data-pernr.
  PERFORM infotyp_lesen USING ps_data-pernr '0105' '18000101' '99991231'
                         CHANGING lt_p0105 p_fehler.
  IF p_fehler IS NOT INITIAL.
    RETURN.
  ENDIF.
  LOOP AT lt_p0105 TRANSPORTING NO FIELDS
    WHERE subty EQ 'WDID' AND sprps EQ space
      AND ( usrid EQ l_wdid OR usrid_long EQ l_wdid_long ).
    l_treffer = l_treffer + 1.
  ENDLOOP.
  IF l_treffer EQ 0.
    p_fehler = 'WDID nicht berechtigt lesbar.'.
    RETURN.
  ENDIF.
  PERFORM infotyp_lesen USING ps_data-pernr '0000' '18000101' p_keydt
                         CHANGING lt_p0000 p_fehler.
  IF p_fehler IS NOT INITIAL.
    RETURN.
  ENDIF.
  DELETE lt_p0000 WHERE sprps NE space.
* Zum Stichtag muss der Mitarbeiter ausgetreten sein (STAT2 = 0).
  CLEAR l_anzahl.
  LOOP AT lt_p0000 INTO wa_p0000
    WHERE begda LE p_keydt AND endda GE p_keydt.
    l_anzahl = l_anzahl + 1.
    IF wa_p0000-stat2 NE '0'.
      p_fehler = 'Am Stichtag nicht ausgetreten (STAT2 <> 0).'.
      RETURN.
    ENDIF.
  ENDLOOP.
  IF l_anzahl NE 1.
    p_fehler = 'IT0000 am Stichtag fehlt/ist mehrdeutig.'.
    RETURN.
  ENDIF.
* Letzter aktiver Kalendertag. Kein Arbeitstags-/Feiertagskalender.
  LOOP AT lt_p0000 INTO wa_p0000 WHERE stat2 EQ '3' AND endda LT p_keydt.
    IF wa_p0000-endda GT ps_data-datum.
      ps_data-datum = wa_p0000-endda.
    ENDIF.
  ENDLOOP.
  IF ps_data-datum IS INITIAL.
    p_fehler = 'Kein letzter aktiver Tag vor Stichtag gefunden.'.
    RETURN.
  ENDIF.
  l_folgetag = ps_data-datum + 1.
  CLEAR l_anzahl.
  LOOP AT lt_p0000 INTO wa_p0000
    WHERE begda LE l_folgetag AND endda GE l_folgetag.
    l_anzahl = l_anzahl + 1.
    IF wa_p0000-stat2 NE '0'.
      p_fehler = 'Auf letzten aktiven Tag folgt kein Austritt.'.
      RETURN.
    ENDIF.
  ENDLOOP.
  IF l_anzahl NE 1.
    p_fehler = 'Austrittsbeginn fehlt/ist mehrdeutig.'.
    RETURN.
  ENDIF.
  CLEAR l_anzahl.
  LOOP AT lt_p0000 INTO wa_p0000
    WHERE begda LE ps_data-datum AND endda GE ps_data-datum.
    l_anzahl = l_anzahl + 1.
    IF wa_p0000-stat2 NE '3'.
      p_fehler = 'Buchungstag ist nicht aktiv.'.
      RETURN.
    ENDIF.
  ENDLOOP.
  IF l_anzahl NE 1.
    p_fehler = 'Aktivstatus am Buchungstag nicht eindeutig.'.
    RETURN.
  ENDIF.
  CLEAR l_anzahl.
  LOOP AT lt_p0105 TRANSPORTING NO FIELDS
    WHERE subty EQ 'WDID' AND sprps EQ space
      AND begda LE ps_data-datum AND endda GE ps_data-datum
      AND ( usrid EQ l_wdid OR usrid_long EQ l_wdid_long ).
    l_anzahl = l_anzahl + 1.
  ENDLOOP.
  IF l_anzahl NE 1.
    p_fehler = 'WDID am Buchungstag fehlt/ist mehrdeutig.'.
    RETURN.
  ENDIF.
  PERFORM infotyp_lesen USING ps_data-pernr '0001' ps_data-datum ps_data-datum
                         CHANGING lt_p0001 p_fehler.
  IF p_fehler IS NOT INITIAL.
    RETURN.
  ENDIF.
  DELETE lt_p0001 WHERE sprps NE space.
  IF lines( lt_p0001 ) NE 1.
    p_fehler = 'IT0001 am Buchungstag fehlt/ist mehrdeutig.'.
    RETURN.
  ENDIF.
  READ TABLE lt_p0001 INDEX 1 INTO wa_p0001.
  ps_data-werks = wa_p0001-werks.
  ps_data-btrtl = wa_p0001-btrtl.
  IF s_west[] IS NOT INITIAL AND wa_p0001-btrtl IN s_west.
    ps_data-lgart = '6600'.
  ENDIF.
  IF s_ost[] IS NOT INITIAL AND wa_p0001-btrtl IN s_ost.
    IF ps_data-lgart IS NOT INITIAL.
      p_fehler = 'Personalteilbereich gleichzeitig West und Ost.'.
      RETURN.
    ENDIF.
    ps_data-lgart = '6610'.
  ENDIF.
  IF ps_data-lgart IS INITIAL.
    CONCATENATE 'Keine West/Ost-Zuordnung fuer' wa_p0001-werks wa_p0001-btrtl
        INTO p_fehler SEPARATED BY space.
    RETURN.
  ENDIF.
* Quelle enthaelt keine Waehrung; fachliche Annahme EUR.
  ps_data-waers = 'EUR'.
* Vorhandene Saetze auch mit anderem Betrag blockieren, niemals addieren.
* Bei leerem IT0015 ist RC 4 regulaer. Fehler/Teilberechtigung nicht ignorieren.
  CALL FUNCTION 'HR_READ_INFOTYPE'
    EXPORTING pernr = ps_data-pernr infty = '0015'
              begda = ps_data-datum endda = ps_data-datum bypass_buffer = 'X'
    IMPORTING subrc = l_rc
    TABLES infty_tab = lt_p0015
    EXCEPTIONS infty_not_found = 1 OTHERS = 2.
  IF sy-subrc NE 0 OR ( l_rc NE 0 AND l_rc NE 4 ).
    p_fehler = 'IT0015 nicht vollstaendig berechtigt lesbar.'.
    RETURN.
  ENDIF.
  LOOP AT lt_p0015 TRANSPORTING NO FIELDS WHERE lgart EQ ps_data-lgart.
    p_fehler = 'IT0015 fuer Datum/Lohnart bereits vorhanden; manuell pruefen.'.
    RETURN.
  ENDLOOP.
ENDFORM.

*----------------------------------------------------------------------*
* Form daten_pruefen
*----------------------------------------------------------------------*
FORM daten_pruefen.
  DATA: l_index   TYPE sy-tabix,
        l_meldung TYPE string,
        lo_error  TYPE REF TO cx_root.

  LOOP AT t_data INTO wa_data WHERE status IS INITIAL.
    l_index = sy-tabix.
    CLEAR l_meldung.
    TRY.
        PERFORM stammdaten_pruefen CHANGING wa_data l_meldung.
      CATCH cx_root INTO lo_error.
        l_meldung = lo_error->get_text( ).
    ENDTRY.
    IF l_meldung IS INITIAL.
      wa_data-status = 'BEREIT'.
      wa_data-meldung = 'Fachliche Vorpruefung erfolgreich.'.
    ELSE.
      wa_data-status = 'FEHLER'.
      wa_data-meldung = l_meldung.
    ENDIF.
    MODIFY t_data FROM wa_data INDEX l_index.
  ENDLOOP.
ENDFORM.

*----------------------------------------------------------------------*
* Form dubletten_pruefen
*----------------------------------------------------------------------*
FORM dubletten_pruefen.
  DATA: l_index  TYPE sy-tabix,
        l_anzahl TYPE i.

* Alle mehrfach vorkommenden Zielschluessel blockieren.
  LOOP AT t_data INTO wa_data WHERE status EQ 'BEREIT'.
    l_index = sy-tabix.
    CLEAR l_anzahl.
    LOOP AT t_data TRANSPORTING NO FIELDS
      WHERE pernr EQ wa_data-pernr
        AND datum EQ wa_data-datum
        AND lgart EQ wa_data-lgart.
      l_anzahl = l_anzahl + 1.
    ENDLOOP.
    IF l_anzahl GT 1.
      wa_data-status = 'FEHLER'.
      wa_data-meldung =
        'Mehrere Excel-Zeilen fuer dieselbe PerNr/Datum/Lohnart.'.
      MODIFY t_data FROM wa_data INDEX l_index.
    ENDIF.
  ENDLOOP.
ENDFORM.

*----------------------------------------------------------------------*
* Form daten_verarbeiten
*----------------------------------------------------------------------*
FORM daten_verarbeiten.
  DATA l_index TYPE sy-tabix.

  LOOP AT t_data INTO wa_data WHERE status EQ 'BEREIT'.
    l_index = sy-tabix.
    IF p_test IS NOT INITIAL.
      wa_data-status = 'TEST'.
      wa_data-meldung =
        'Vorpruefung OK; keine Buchung, keine SAP-Schreibpruefung.'.
    ELSE.
      PERFORM personalnummer_buchen CHANGING wa_data.
    ENDIF.
    MODIFY t_data FROM wa_data INDEX l_index.
  ENDLOOP.
ENDFORM.

*----------------------------------------------------------------------*
* Form personalnummer_buchen
*----------------------------------------------------------------------*
FORM personalnummer_buchen CHANGING ps_data TYPE ty_data.
  DATA: wa_return   TYPE bapireturn1,
        wa_original TYPE ty_data,
        l_gesperrt  TYPE c LENGTH 1,
        l_meldung   TYPE string,
        lo_error    TYPE REF TO cx_root.

  wa_original = ps_data.
  TRY.
      CALL FUNCTION 'BAPI_EMPLOYEE_ENQUEUE'
        EXPORTING
          number = ps_data-pernr
        IMPORTING
          return = wa_return.
      IF wa_return-type CA 'AEX'.
        l_meldung = wa_return-message.
        IF l_meldung IS INITIAL.
          l_meldung = 'Personalnummer konnte nicht gesperrt werden.'.
        ENDIF.
      ELSE.
        l_gesperrt = 'X'.
* Stammdaten und Dubletten innerhalb der Mitarbeitersperre neu pruefen.
        PERFORM stammdaten_pruefen CHANGING ps_data l_meldung.
        IF l_meldung IS INITIAL.
          IF ps_data-pernr NE wa_original-pernr OR
             ps_data-datum NE wa_original-datum OR
             ps_data-lgart NE wa_original-lgart OR
             ps_data-btrtl NE wa_original-btrtl OR
             ps_data-werks NE wa_original-werks.
            l_meldung = 'Stammdaten seit Vorpruefung geaendert; neu starten.'.
          ELSE.
            PERFORM infotyp_verbuchen CHANGING ps_data l_meldung.
          ENDIF.
        ENDIF.
      ENDIF.
    CATCH cx_root INTO lo_error.
      l_meldung = lo_error->get_text( ).
  ENDTRY.

  IF l_meldung IS NOT INITIAL.
    ROLLBACK WORK.
    ps_data-status = 'FEHLER'.
    ps_data-meldung = l_meldung.
  ENDIF.
  IF l_gesperrt IS NOT INITIAL.
    CALL FUNCTION 'BAPI_EMPLOYEE_DEQUEUE'
      EXPORTING
        number = wa_original-pernr.
  ENDIF.
  CALL FUNCTION 'HR_PSBUFFER_INITIALIZE'.
ENDFORM.

*----------------------------------------------------------------------*
* Form infotyp_verbuchen
*----------------------------------------------------------------------*
FORM infotyp_verbuchen CHANGING ps_data TYPE ty_data
                                 p_fehler TYPE string.
  DATA: wa_p0015  TYPE p0015,
        wa_return TYPE bapireturn1.

  CLEAR: p_fehler, wa_p0015, wa_return.
  wa_p0015-pernr = ps_data-pernr.
  wa_p0015-infty = '0015'.
  wa_p0015-subty = ps_data-lgart.
  wa_p0015-lgart = ps_data-lgart.
  wa_p0015-begda = ps_data-datum.
  wa_p0015-endda = ps_data-datum.
  wa_p0015-betrg = ps_data-betrag.
  wa_p0015-waers = ps_data-waers.

  CALL FUNCTION 'HR_INFOTYPE_OPERATION'
    EXPORTING
      infty         = '0015'
      number        = wa_p0015-pernr
      subtype       = wa_p0015-subty
      validitybegin = wa_p0015-begda
      validityend   = wa_p0015-endda
      record        = wa_p0015
      operation     = 'INS'
      tclas         = 'A'
      dialog_mode   = '0'
      nocommit      = 'X'
    IMPORTING
      return        = wa_return
    EXCEPTIONS
      OTHERS        = 1.
  IF sy-subrc NE 0 OR wa_return-type CA 'AEX'.
    CONCATENATE 'IT0015 fehlgeschlagen:' wa_return-message
           INTO p_fehler SEPARATED BY space.
    RETURN.
  ENDIF.

  COMMIT WORK AND WAIT.
  IF sy-subrc NE 0.
    p_fehler = 'Update fehlgeschlagen; SM13 und PA20 pruefen.'.
    RETURN.
  ENDIF.
  ps_data-status = 'GEBUCHT'.
  CONCATENATE 'IT0015 angelegt.' wa_return-message
         INTO ps_data-meldung SEPARATED BY space.
ENDFORM.

*----------------------------------------------------------------------*
* Form alv_ausgabe
*----------------------------------------------------------------------*
FORM alv_ausgabe.
  DATA: lo_alv        TYPE REF TO cl_salv_table,
        lo_spalten    TYPE REF TO cl_salv_columns_table,
        lo_funktionen TYPE REF TO cl_salv_functions_list,
        lo_error      TYPE REF TO cx_root,
        l_meldung     TYPE string.

  TRY.
      CALL METHOD cl_salv_table=>factory
        IMPORTING
          r_salv_table = lo_alv
        CHANGING
          t_table      = t_data.
      lo_spalten = lo_alv->get_columns( ).

      PERFORM alv_spalte_beschriften USING lo_spalten
        'ZEILE' 'Excelzeile'
        'Excel-Zeile' 'Zeilennummer in der Excel-Datei'.
      PERFORM alv_spalte_beschriften USING lo_spalten
        'WPN' 'Workday-ID'
        'Workday-ID (WPN)' 'Workday-ID (WPN)'.
      PERFORM alv_spalte_beschriften USING lo_spalten
        'BETRAG' 'AEG Brutto'
        'AEG Brutto' 'Zahlungsbetrag AEG Brutto'.
      PERFORM alv_spalte_beschriften USING lo_spalten
        'PERNR' 'SAP-PersNr'
        'SAP-Personalnummer' 'SAP-Personalnummer'.
      PERFORM alv_spalte_beschriften USING lo_spalten
        'DATUM' 'Buch.-Tag'
        'Buchungsdatum' 'Buchungsdatum (letzter aktiver Tag)'.
      PERFORM alv_spalte_beschriften USING lo_spalten
        'WERKS' 'PersBer'
        'Personalbereich' 'SAP-Personalbereich'.
      PERFORM alv_spalte_beschriften USING lo_spalten
        'BTRTL' 'PersTBer'
        'Personalteilbereich' 'SAP-Personalteilbereich'.
      PERFORM alv_spalte_beschriften USING lo_spalten
        'LGART' 'Lohnart'
        'Lohnart IT0015' 'Lohnart fuer IT0015'.
      PERFORM alv_spalte_beschriften USING lo_spalten
        'WAERS' 'Waehrung'
        'Waehrung' 'Waehrung des Zahlungsbetrags'.
      PERFORM alv_spalte_beschriften USING lo_spalten
        'STATUS' 'Status'
        'Buchungsstatus' 'Pruef- und Buchungsstatus'.
      PERFORM alv_spalte_beschriften USING lo_spalten
        'MELDUNG' 'Meldung'
        'Verarbeitungsmeldung' 'Pruefergebnis / Verarbeitungsmeldung'.

      lo_spalten->set_optimize( abap_true ).
      lo_funktionen = lo_alv->get_functions( ).
      lo_funktionen->set_all( abap_true ).
      lo_alv->display( ).
    CATCH cx_root INTO lo_error.
      l_meldung = lo_error->get_text( ).
      MESSAGE l_meldung TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.
ENDFORM.

*----------------------------------------------------------------------*
* Form alv_spalte_beschriften
*----------------------------------------------------------------------*
FORM alv_spalte_beschriften
  USING po_spalten TYPE REF TO cl_salv_columns_table
        VALUE(p_name) TYPE salv_de_column
        VALUE(p_kurz) TYPE scrtext_s
        VALUE(p_mittel) TYPE scrtext_m
        VALUE(p_lang) TYPE scrtext_l
  RAISING cx_salv_not_found.

  DATA lo_spalte TYPE REF TO cl_salv_column.

  lo_spalte = po_spalten->get_column( p_name ).
  lo_spalte->set_short_text( p_kurz ).
  lo_spalte->set_medium_text( p_mittel ).
  lo_spalte->set_long_text( p_lang ).
ENDFORM.
