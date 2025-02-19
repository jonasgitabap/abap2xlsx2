*"* use this source file for your ABAP unit test classes

CLASS ltc_column_formula DEFINITION DEFERRED.
CLASS ltc_is_formula_shareable DEFINITION DEFERRED.
CLASS ltc_escape_string_value DEFINITION DEFERRED.
CLASS /nag/cl_excel_writer_2007 DEFINITION LOCAL FRIENDS
    ltc_column_formula
    ltc_is_formula_shareable
    ltc_escape_string_value.

CLASS ltc_column_formula DEFINITION FOR TESTING
    RISK LEVEL HARMLESS
    DURATION SHORT.

  PRIVATE SECTION.

    METHODS one_column_formula FOR TESTING RAISING cx_static_check.
    METHODS two_column_formulas FOR TESTING RAISING cx_static_check.

    METHODS setup.
    METHODS set_cell
      IMPORTING
        is_cell_data TYPE /nag/excel_s_cell_data
      RAISING
        /nag/cx_excel.
    METHODS get_xml
      RETURNING
        VALUE(rv_xml) TYPE string.

    DATA lo_document TYPE REF TO if_ixml_document.
    DATA lo_excel TYPE REF TO /nag/cl_excel_writer_2007.
    DATA lo_ixml TYPE REF TO if_ixml.
    DATA lo_root TYPE REF TO if_ixml_element.
    DATA lt_column_formulas TYPE /nag/cl_excel_worksheet=>mty_th_column_formula.
    DATA lt_column_formulas_used TYPE /nag/cl_excel_writer_2007=>mty_column_formulas_used.
    DATA lv_si TYPE i.

ENDCLASS.


CLASS ltc_is_formula_shareable DEFINITION FOR TESTING
    RISK LEVEL HARMLESS
    DURATION SHORT.

  PRIVATE SECTION.

    METHODS is_formula_shareable FOR TESTING RAISING cx_static_check.

ENDCLASS.


CLASS ltc_escape_string_value DEFINITION
      FOR TESTING
      DURATION SHORT
      RISK LEVEL HARMLESS.

  PRIVATE SECTION.

    METHODS escaped_character_inside_text FOR TESTING.
    METHODS no_escaping FOR TESTING.
    METHODS one_escaped_character FOR TESTING.
    METHODS two_escaped_characters FOR TESTING.

    METHODS run_cut
      IMPORTING
        input TYPE string
        exp   TYPE string.

ENDCLASS.


CLASS ltc_column_formula IMPLEMENTATION.

  METHOD one_column_formula.

    DATA ls_cell_data TYPE /nag/excel_s_cell_data.
    DATA lv_xml TYPE string.

    ls_cell_data-cell_row = 2.
    ls_cell_data-cell_column = 1.
    ls_cell_data-column_formula_id = 1.
    set_cell( ls_cell_data ).

    lv_xml = get_xml( ).

    cl_abap_unit_assert=>assert_equals( act = lv_xml exp = '<c r="R2C1"><f ref="A2:A4" si="0" t="shared">A2</f></c>' ).

  ENDMETHOD.

  METHOD two_column_formulas.

    DATA ls_cell_data TYPE /nag/excel_s_cell_data.
    DATA lv_xml TYPE string.

    ls_cell_data-cell_row = 2.
    ls_cell_data-cell_column = 1.
    ls_cell_data-column_formula_id = 1.
    set_cell( ls_cell_data ).

    ls_cell_data-cell_row = 3.
    ls_cell_data-cell_column = 1.
    ls_cell_data-column_formula_id = 1.
    set_cell( ls_cell_data ).

    lv_xml = get_xml( ).

    cl_abap_unit_assert=>assert_equals( act = lv_xml exp = '<c r="R2C1"><f ref="A2:A4" si="0" t="shared">A2</f></c>'
                                                        && '<c r="R3C1"><f si="0" t="shared"/></c>' ).
  ENDMETHOD.


  METHOD set_cell.

    DATA lo_cell TYPE REF TO if_ixml_element.
    DATA lo_formula TYPE REF TO if_ixml_element.

    lo_cell = lo_document->create_element( 'c' ).
    lo_cell->set_attribute( name = 'r' value = |R{ is_cell_data-cell_row }C{ is_cell_data-cell_column }| ).
    lo_root->append_child( lo_cell ).

    lo_excel->create_xl_sheet_column_formula(
              EXPORTING
                io_document             = lo_document
                it_column_formulas      = lt_column_formulas
                is_sheet_content        = is_cell_data
              IMPORTING
                eo_element              = lo_formula
              CHANGING
                ct_column_formulas_used = lt_column_formulas_used
                cv_si                   = lv_si ).
    lo_cell->append_child( lo_formula ).

  ENDMETHOD.


  METHOD get_xml.

    DATA lo_stream_factory TYPE REF TO if_ixml_stream_factory.
    DATA lo_ostream TYPE REF TO if_ixml_ostream.
    DATA lo_renderer TYPE REF TO if_ixml_renderer.

    lo_stream_factory = lo_ixml->create_stream_factory( ).
    lo_ostream = lo_stream_factory->create_ostream_cstring( string = rv_xml ).
    lo_renderer = lo_ixml->create_renderer( ostream = lo_ostream document = lo_document ).
    lo_renderer->render( ).
    CALL TRANSFORMATION id SOURCE XML rv_xml RESULT XML rv_xml OPTIONS xml_header = 'no'.
    IF strlen( rv_xml ) >= 2.
      rv_xml = substring( val = rv_xml off = 1 ). " Remove Byte Order Mark (one character in Unicode systems)
    ENDIF.
    FIND REGEX '^<dummy>(.*)</dummy>$' IN rv_xml SUBMATCHES rv_xml.

  ENDMETHOD.


  METHOD setup.

    DATA ls_column_formula TYPE /nag/cl_excel_worksheet=>mty_s_column_formula.

    lo_ixml = cl_ixml=>create( ).
    lo_document = lo_ixml->create_document( ).
    lo_root = lo_document->create_element( 'dummy' ).
    lo_document->append_child( lo_root ).

    CREATE OBJECT lo_excel.

    ls_column_formula-id = 1.
    ls_column_formula-formula = 'A2'.
    ls_column_formula-table_top_left_row = 1.
    ls_column_formula-table_bottom_right_row = 3.
    INSERT ls_column_formula INTO TABLE lt_column_formulas.

  ENDMETHOD.

ENDCLASS.


CLASS ltc_is_formula_shareable IMPLEMENTATION.

  METHOD is_formula_shareable.

    DATA lo_excel TYPE REF TO /nag/cl_excel_writer_2007.

    CREATE OBJECT lo_excel.

    cl_abap_unit_assert=>assert_equals( act = lo_excel->is_formula_shareable( 'TblFlights[[#This Row],[Airfare]]+222' ) exp = abap_false ). " [@Airfare]+222
    cl_abap_unit_assert=>assert_equals( act = lo_excel->is_formula_shareable( 'OtherSheet!A2' ) exp = abap_false ).
    cl_abap_unit_assert=>assert_equals( act = lo_excel->is_formula_shareable( 'D2+100' ) exp = abap_true ).
    cl_abap_unit_assert=>assert_equals( act = lo_excel->is_formula_shareable( 'A1&";"&_xlfn.IFS(TRUE,NamedRange)' ) exp = abap_true ).

  ENDMETHOD.

ENDCLASS.


CLASS ltc_escape_string_value IMPLEMENTATION.

  METHOD escaped_character_inside_text.
    run_cut( input = 'start U+0000 end' exp = 'start _x0000_ end' ).
  ENDMETHOD.

  METHOD no_escaping.
    run_cut( input = 'no escaping' exp = 'no escaping' ).
  ENDMETHOD.

  METHOD one_escaped_character.
    run_cut( input = 'U+0000' exp = '_x0000_' ).
  ENDMETHOD.

  METHOD run_cut.

    DATA: lo_excel    TYPE REF TO /nag/cl_excel_writer_2007,
          lv_input    TYPE string,
          lt_matches  TYPE match_result_tab,
          ls_match    TYPE match_result,
          ls_submatch TYPE submatch_result,
          lv_hexa     TYPE c LENGTH 4,
          lv_uccp     TYPE c LENGTH 1.

    CREATE OBJECT lo_excel.

    lv_input = input.
    FIND ALL OCCURRENCES OF REGEX 'U\+(....)' IN lv_input RESULTS lt_matches.
    SORT lt_matches BY offset DESCENDING.
    LOOP AT lt_matches INTO ls_match.
      READ TABLE ls_match-submatches INDEX 1 INTO ls_submatch.
      ASSERT sy-subrc = 0.
      lv_hexa = substring( val = lv_input off = ls_submatch-offset len = ls_submatch-length ).
      lv_uccp = cl_abap_conv_in_ce=>uccp( lv_hexa ).
      REPLACE SECTION OFFSET ls_match-offset LENGTH ls_match-length OF lv_input WITH lv_uccp.
    ENDLOOP.

    cl_abap_unit_assert=>assert_equals( act = lo_excel->escape_string_value( lv_input ) exp = exp msg = |input: { input }| ).

  ENDMETHOD.

  METHOD two_escaped_characters.
    run_cut( input = 'U+0000 and U+FFFF' exp = '_x0000_ and _xFFFF_' ).
  ENDMETHOD.

ENDCLASS.
