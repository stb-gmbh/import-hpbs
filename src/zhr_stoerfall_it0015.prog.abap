REPORT zhr_stoerfall_it0015.
* ABAP >= 7.40; abap2xlsx; SAP GUI; classic SAP HCM.
* Keine direkten Updates auf PA0015. Siehe begleitende Anleitung.
TABLES: pa0001.
PARAMETERS: p_file TYPE string LOWER CASE OBLIGATORY,
            p_keydt TYPE sy-datum OBLIGATORY DEFAULT sy-datum,
            p_test AS CHECKBOX DEFAULT 'X'.
SELECT-OPTIONS: s_west FOR pa0001-btrtl NO INTERVALS,
                s_ost  FOR pa0001-btrtl NO INTERVALS.

TYPES: BEGIN OF ty_row,
         zeile TYPE i,
         wpn TYPE string,
         betrag TYPE p0015-betrg,
         pernr TYPE pernr_d,
         datum TYPE sy-datum,
         werks TYPE p0001-werks,
         btrtl TYPE p0001-btrtl,
         lgart TYPE p0015-lgart,
         waers TYPE p0015-waers,
         status TYPE c LENGTH 10,
         meldung TYPE string,
       END OF ty_row.
DATA gt_rows TYPE STANDARD TABLE OF ty_row WITH DEFAULT KEY.

CLASS lcx_error DEFINITION INHERITING FROM cx_static_check.
  PUBLIC SECTION.
    DATA detail TYPE string READ-ONLY.
    METHODS constructor IMPORTING text TYPE string.
ENDCLASS.
CLASS lcx_error IMPLEMENTATION.
  METHOD constructor.
    super->constructor( ).
    detail = text.
  ENDMETHOD.
ENDCLASS.

CLASS lcl_app DEFINITION FINAL.
  PUBLIC SECTION.
    CLASS-METHODS run RAISING cx_static_check.
  PRIVATE SECTION.
    CLASS-METHODS cell IMPORTING sheet TYPE REF TO zcl_excel_worksheet
                                col TYPE i row TYPE i
                      RETURNING VALUE(value) TYPE string
                      RAISING zcx_excel lcx_error.
    CLASS-METHODS read_it IMPORTING pernr TYPE pernr_d infty TYPE infty
                                   begda TYPE sy-datum endda TYPE sy-datum
                         CHANGING records TYPE STANDARD TABLE
                         RAISING lcx_error.
    CLASS-METHODS prepare CHANGING item TYPE ty_row RAISING lcx_error.
    CLASS-METHODS process CHANGING item TYPE ty_row.
ENDCLASS.

CLASS lcl_app IMPLEMENTATION.
  METHOD cell.
    DATA raw TYPE zexcel_cell_value.
    DATA formula TYPE zexcel_cell_formula.
    sheet->get_cell( EXPORTING ip_column = col ip_row = row
                     IMPORTING ep_value = raw ep_formula = formula ).
    IF formula IS NOT INITIAL.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = |Zeile { row }: Formeln nicht erlaubt; Werte exportieren.|.
    ENDIF.
    value = raw.
    CONDENSE value.
  ENDMETHOD.

  METHOD read_it.
    DATA rc TYPE sy-subrc.
    CLEAR records.
    CALL FUNCTION 'HR_READ_INFOTYPE'
      EXPORTING pernr = pernr infty = infty begda = begda endda = endda
                bypass_buffer = 'X'
      IMPORTING subrc = rc
      TABLES infty_tab = records
      EXCEPTIONS infty_not_found = 1 OTHERS = 2.
    IF sy-subrc <> 0 OR rc <> 0.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = |IT{ infty }: Lesen fehlgeschlagen/keine Berechtigung (RC { rc }).|.
    ENDIF.
  ENDMETHOD.

  METHOD prepare.
    DATA: ids TYPE SORTED TABLE OF pernr_d WITH UNIQUE KEY table_line,
          short_id TYPE pa0105-usrid,
          long_id TYPE pa0105-usrid_long,
          it105 TYPE STANDARD TABLE OF p0105,
          it000 TYPE STANDARD TABLE OF p0000,
          it001 TYPE STANDARD TABLE OF p0001,
          it015 TYPE STANDARD TABLE OF p0015,
          next_day TYPE sy-datum,
          count TYPE i,
          matches TYPE i.
    CLEAR: item-pernr, item-datum, item-werks, item-btrtl,
           item-lgart, item-waers.
    IF item-wpn IS INITIAL OR strlen( item-wpn ) > 30.
      RAISE EXCEPTION TYPE lcx_error EXPORTING text = 'WPN fehlt/ist zu lang.'.
    ENDIF.
    short_id = item-wpn.
    long_id = item-wpn.
* SQL nur Kandidatensuche. Keine Stammdatenverwendung ohne HR-Lesepruefung.
* Historische WDID-Saetze einschliessen, da ausgeschiedene Mitarbeiter.
    SELECT DISTINCT pernr FROM pa0105 INTO TABLE ids
      WHERE subty = 'WDID' AND sprps = space
        AND ( usrid = short_id OR usrid_long = long_id ).
    IF lines( ids ) <> 1.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'WDID fehlt oder ist mehreren Personalnummern zugeordnet.'.
    ENDIF.
    READ TABLE ids INDEX 1 INTO item-pernr.
    read_it( EXPORTING pernr = item-pernr infty = '0105'
                       begda = '18000101' endda = '99991231'
             CHANGING records = it105 ).
    LOOP AT it105 TRANSPORTING NO FIELDS
      WHERE subty = 'WDID' AND sprps = space
        AND ( usrid = short_id OR usrid_long = long_id ).
      matches = matches + 1.
    ENDLOOP.
    IF matches = 0.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'WDID nicht berechtigt lesbar.'.
    ENDIF.
    read_it( EXPORTING pernr = item-pernr infty = '0000'
                       begda = '18000101' endda = p_keydt
             CHANGING records = it000 ).
    DELETE it000 WHERE sprps <> space.
* Zum Stichtag muss der Mitarbeiter ausgetreten sein (STAT2 = 0).
    CLEAR count.
    LOOP AT it000 INTO DATA(action)
      WHERE begda <= p_keydt AND endda >= p_keydt.
      count = count + 1.
      IF action-stat2 <> '0'.
        RAISE EXCEPTION TYPE lcx_error
          EXPORTING text = 'Am Stichtag nicht ausgetreten (STAT2 <> 0).'.
      ENDIF.
    ENDLOOP.
    IF count <> 1.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'IT0000 am Stichtag fehlt/ist mehrdeutig.'.
    ENDIF.
* Letzter aktiver Kalendertag. Kein Arbeitstags-/Feiertagskalender.
    LOOP AT it000 INTO action WHERE stat2 = '3' AND endda < p_keydt.
      IF action-endda > item-datum.
        item-datum = action-endda.
      ENDIF.
    ENDLOOP.
    IF item-datum IS INITIAL.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'Kein letzter aktiver Tag vor Stichtag gefunden.'.
    ENDIF.
    next_day = item-datum + 1.
    CLEAR count.
    LOOP AT it000 INTO action
      WHERE begda <= next_day AND endda >= next_day.
      count = count + 1.
      IF action-stat2 <> '0'.
        RAISE EXCEPTION TYPE lcx_error
          EXPORTING text = 'Auf letzten aktiven Tag folgt kein Austritt.'.
      ENDIF.
    ENDLOOP.
    IF count <> 1.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'Austrittsbeginn fehlt/ist mehrdeutig.'.
    ENDIF.
    CLEAR count.
    LOOP AT it000 INTO action
      WHERE begda <= item-datum AND endda >= item-datum.
      count = count + 1.
      IF action-stat2 <> '3'.
        RAISE EXCEPTION TYPE lcx_error
          EXPORTING text = 'Buchungstag ist nicht aktiv.'.
      ENDIF.
    ENDLOOP.
    IF count <> 1.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'Aktivstatus am Buchungstag nicht eindeutig.'.
    ENDIF.
    CLEAR count.
    LOOP AT it105 TRANSPORTING NO FIELDS
      WHERE subty = 'WDID' AND sprps = space
        AND begda <= item-datum AND endda >= item-datum
        AND ( usrid = short_id OR usrid_long = long_id ).
      count = count + 1.
    ENDLOOP.
    IF count <> 1.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'WDID am Buchungstag fehlt/ist mehrdeutig.'.
    ENDIF.
    read_it( EXPORTING pernr = item-pernr infty = '0001'
                       begda = item-datum endda = item-datum
             CHANGING records = it001 ).
    DELETE it001 WHERE sprps <> space.
    IF lines( it001 ) <> 1.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'IT0001 am Buchungstag fehlt/ist mehrdeutig.'.
    ENDIF.
    READ TABLE it001 INDEX 1 INTO DATA(org).
    item-werks = org-werks.
    item-btrtl = org-btrtl.
    IF s_west[] IS NOT INITIAL AND org-btrtl IN s_west.
      item-lgart = '6600'.
    ENDIF.
    IF s_ost[] IS NOT INITIAL AND org-btrtl IN s_ost.
      IF item-lgart IS NOT INITIAL.
        RAISE EXCEPTION TYPE lcx_error
          EXPORTING text = 'Personalteilbereich gleichzeitig West und Ost.'.
      ENDIF.
      item-lgart = '6610'.
    ENDIF.
    IF item-lgart IS INITIAL.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = |Keine West/Ost-Zuordnung fuer { org-werks }/{ org-btrtl }.|.
    ENDIF.
* Quelle enthaelt keine Waehrung; fachliche Annahme EUR.
    item-waers = 'EUR'.
* Vorhandene Saetze auch mit anderem Betrag blockieren, niemals addieren.
* Bei leerem IT0015 ist RC 4 regulaer. Fehler/Teilberechtigung nicht ignorieren.
    DATA rc TYPE sy-subrc.
    CALL FUNCTION 'HR_READ_INFOTYPE'
      EXPORTING pernr = item-pernr infty = '0015'
                begda = item-datum endda = item-datum bypass_buffer = 'X'
      IMPORTING subrc = rc
      TABLES infty_tab = it015
      EXCEPTIONS infty_not_found = 1 OTHERS = 2.
    IF sy-subrc <> 0 OR ( rc <> 0 AND rc <> 4 ).
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'IT0015 nicht vollstaendig berechtigt lesbar.'.
    ENDIF.
    LOOP AT it015 TRANSPORTING NO FIELDS WHERE lgart = item-lgart.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'IT0015 fuer Datum/Lohnart bereits vorhanden; manuell pruefen.'.
    ENDLOOP.
  ENDMETHOD.

  METHOD process.
    DATA: ret TYPE bapireturn1, record TYPE p0015,
          locked TYPE abap_bool, original TYPE ty_row.
    original = item.
    TRY.
        CALL FUNCTION 'BAPI_EMPLOYEE_ENQUEUE'
          EXPORTING number = item-pernr IMPORTING return = ret.
        IF ret-type CA 'AEX'.
          RAISE EXCEPTION TYPE lcx_error EXPORTING text = CONV string( ret-message ).
        ENDIF.
        locked = abap_true.
* Erneute Ableitung/Dublettenpruefung innerhalb der Mitarbeitersperre.
        prepare( CHANGING item = item ).
        IF item-pernr <> original-pernr OR item-datum <> original-datum
           OR item-lgart <> original-lgart OR item-btrtl <> original-btrtl
           OR item-werks <> original-werks.
          RAISE EXCEPTION TYPE lcx_error
            EXPORTING text = 'Stammdaten seit Vorpruefung geaendert; neu starten.'.
        ENDIF.
        record-pernr = item-pernr.
        record-infty = '0015'.
        record-subty = item-lgart.
        record-lgart = item-lgart.
        record-begda = item-datum.
        record-endda = item-datum.
        record-betrg = item-betrag.
        record-waers = item-waers.
        CLEAR ret.
        CALL FUNCTION 'HR_INFOTYPE_OPERATION'
          EXPORTING infty = '0015' number = record-pernr
                    subtype = record-subty validitybegin = record-begda
                    validityend = record-endda record = record
                    operation = 'INS' tclas = 'A' dialog_mode = '0'
                    nocommit = 'X'
          IMPORTING return = ret
          EXCEPTIONS OTHERS = 1.
        IF sy-subrc <> 0 OR ret-type CA 'AEX'.
          RAISE EXCEPTION TYPE lcx_error
            EXPORTING text = |IT0015 fehlgeschlagen: { ret-message }|.
        ENDIF.
        COMMIT WORK AND WAIT.
        IF sy-subrc <> 0.
          RAISE EXCEPTION TYPE lcx_error
            EXPORTING text = 'Update fehlgeschlagen; SM13 und PA20 pruefen.'.
        ENDIF.
        item-status = 'GEBUCHT'.
        item-meldung = |IT0015 angelegt. { ret-message }|.
      CATCH lcx_error INTO DATA(err).
        ROLLBACK WORK.
        item-status = 'FEHLER'.
        item-meldung = err->detail.
      CATCH cx_root INTO DATA(unexpected).
        ROLLBACK WORK.
        item-status = 'FEHLER'.
        item-meldung = unexpected->get_text( ).
    ENDTRY.
    IF locked = abap_true.
      CALL FUNCTION 'BAPI_EMPLOYEE_DEQUEUE'
        EXPORTING number = original-pernr.
    ENDIF.
    CALL FUNCTION 'HR_PSBUFFER_INITIALIZE'.
  ENDMETHOD.

  METHOD run.
    DATA reader TYPE REF TO zif_excel_reader.
    CREATE OBJECT reader TYPE zcl_excel_reader_2007.
    DATA(book) = reader->load_file( i_filename = p_file
                                    i_from_applserver = abap_false ).
    DATA(sheet) = book->get_worksheet_by_index( iv_index = 1 ).
    IF sheet->get_title( ) <> 'ExportData'.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'Erstes Tabellenblatt muss ExportData heissen.'.
    ENDIF.
    DATA: col_wpn TYPE i, col_amt TYPE i.
    DO sheet->get_highest_column( ) TIMES.
      DATA(col) = sy-index.
      DATA(header) = cell( sheet = sheet col = col row = 1 ).
      CASE header.
        WHEN 'WPN'.
          IF col_wpn <> 0.
            RAISE EXCEPTION TYPE lcx_error EXPORTING text = 'WPN-Kopf doppelt.'.
          ENDIF.
          col_wpn = col.
        WHEN 'AEG Brutto'.
          IF col_amt <> 0.
            RAISE EXCEPTION TYPE lcx_error EXPORTING text = 'Betragskopf doppelt.'.
          ENDIF.
          col_amt = col.
      ENDCASE.
    ENDDO.
    IF col_wpn = 0 OR col_amt = 0.
      RAISE EXCEPTION TYPE lcx_error
        EXPORTING text = 'Spalten WPN/AEG Brutto fehlen in Zeile 1.'.
    ENDIF.
    DATA total_rows TYPE i.
    total_rows = sheet->get_highest_row( ) - 1.
    DO total_rows TIMES.
      DATA(item) = VALUE ty_row( zeile = sy-index + 1 ).
      TRY.
          item-wpn = cell( sheet = sheet col = col_wpn row = item-zeile ).
          DATA(amount) = cell( sheet = sheet col = col_amt row = item-zeile ).
          IF item-wpn IS INITIAL AND amount IS INITIAL.
            CONTINUE.
          ENDIF.
          FIND REGEX '^[0-9]+$' IN item-wpn.
          IF sy-subrc <> 0.
            RAISE EXCEPTION TYPE lcx_error
              EXPORTING text = 'WPN muss eine Ziffernfolge sein (keine Exponentialzahl).'.
          ENDIF.
* Numerische XLSX-Zellen liefern Dezimalpunkt, unabhaengig von SAP-Benutzerformat.
* Keine stillschweigende Rundung oder Tausenderseparator-Interpretation.
          FIND REGEX '^[+-]?[0-9]+([.][0-9]{1,2})?$' IN amount.
          IF sy-subrc <> 0.
            RAISE EXCEPTION TYPE lcx_error
              EXPORTING text = 'AEG Brutto: Zahl mit Dezimalpunkt und max. 2 Nachkommastellen erforderlich.'.
          ENDIF.
          item-betrag = CONV decfloat34( amount ).
          IF item-betrag = 0.
            RAISE EXCEPTION TYPE lcx_error EXPORTING text = 'Nullbetrag nicht gebucht.'.
          ENDIF.
          prepare( CHANGING item = item ).
          item-status = 'BEREIT'.
          item-meldung = 'Fachliche Vorpruefung erfolgreich.'.
        CATCH lcx_error INTO DATA(err).
          item-status = 'FEHLER'. item-meldung = err->detail.
        CATCH cx_sy_conversion_error INTO DATA(conv_err).
          item-status = 'FEHLER'. item-meldung = conv_err->get_text( ).
      ENDTRY.
      APPEND item TO gt_rows.
    ENDDO.
* Alle gleichen Zielschluessel blockieren, auch bei verschiedenen Betraegen.
    LOOP AT gt_rows ASSIGNING FIELD-SYMBOL(<row>) WHERE status = 'BEREIT'.
      DATA(hits) = 0.
      LOOP AT gt_rows TRANSPORTING NO FIELDS
        WHERE pernr = <row>-pernr AND datum = <row>-datum
          AND lgart = <row>-lgart.
        hits = hits + 1.
      ENDLOOP.
      IF hits > 1.
        <row>-status = 'FEHLER'.
        <row>-meldung = 'Mehrere Excel-Zeilen fuer dieselbe PerNr/Datum/Lohnart.'.
      ENDIF.
    ENDLOOP.
    LOOP AT gt_rows ASSIGNING <row> WHERE status = 'BEREIT'.
      IF p_test = abap_true.
        <row>-status = 'TEST'.
        <row>-meldung = 'Vorpruefung OK; keine Buchung, keine SAP-Schreibpruefung.'.
      ELSE.
        process( CHANGING item = <row> ).
      ENDIF.
    ENDLOOP.
    DATA alv TYPE REF TO cl_salv_table.
    cl_salv_table=>factory( IMPORTING r_salv_table = alv
                            CHANGING t_table = gt_rows ).
    alv->get_columns( )->set_optimize( abap_true ).
    alv->get_functions( )->set_all( abap_true ).
    alv->display( ).
  ENDMETHOD.
ENDCLASS.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  DATA: files TYPE filetable, rc TYPE i.
  cl_gui_frontend_services=>file_open_dialog(
    EXPORTING file_filter = 'Excel (*.xlsx)|*.xlsx|'
    CHANGING file_table = files rc = rc
    EXCEPTIONS OTHERS = 1 ).
  IF sy-subrc = 0 AND rc > 0.
    READ TABLE files INDEX 1 INTO DATA(file).
    p_file = file-filename.
  ENDIF.

AT SELECTION-SCREEN.
  IF sy-batch = abap_true.
    MESSAGE 'Dieser Report benoetigt SAP GUI fuer den lokalen XLSX-Upload.' TYPE 'E'.
  ENDIF.
  IF s_west[] IS INITIAL AND s_ost[] IS INITIAL.
    MESSAGE 'Mindestens eine West-/Ost-Zuordnung eingeben.' TYPE 'E'.
  ENDIF.
  LOOP AT s_west.
    IF s_west-sign <> 'I' OR s_west-option <> 'EQ'.
      MESSAGE 'West: nur einzelne eingeschlossene Teilbereiche erlaubt.' TYPE 'E'.
    ENDIF.
    IF s_ost[] IS NOT INITIAL AND s_west-low IN s_ost.
      MESSAGE 'West und Ost duerfen sich nicht ueberschneiden.' TYPE 'E'.
    ENDIF.
  ENDLOOP.
  LOOP AT s_ost.
    IF s_ost-sign <> 'I' OR s_ost-option <> 'EQ'.
      MESSAGE 'Ost: nur einzelne eingeschlossene Teilbereiche erlaubt.' TYPE 'E'.
    ENDIF.
  ENDLOOP.

START-OF-SELECTION.
  TRY.
      lcl_app=>run( ).
    CATCH lcx_error INTO DATA(error).
      MESSAGE error->detail TYPE 'S' DISPLAY LIKE 'E'.
    CATCH cx_root INTO DATA(unexpected).
      MESSAGE unexpected->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.

