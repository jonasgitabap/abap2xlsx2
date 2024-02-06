CLASS /nag/cl_excel_style_number_format DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF t_num_format,
        id     TYPE string,
        format TYPE REF TO /nag/cl_excel_style_number_format,
      END OF t_num_format .
    TYPES:
      t_num_formats TYPE HASHED TABLE OF t_num_format WITH UNIQUE KEY id .

*"* public components of class /NAG/CL_EXCEL_STYLE_NUMBER_FORMAT
*"* do not include other source files here!!!
    CONSTANTS c_format_numc_std TYPE /nag/excel_number_format VALUE 'STD_NDEC'. "#EC NOTEXT
    CONSTANTS c_format_date_std TYPE /nag/excel_number_format VALUE 'STD_DATE'. "#EC NOTEXT
    CONSTANTS c_format_currency_eur_simple TYPE /nag/excel_number_format VALUE '[$EUR ]#,##0.00_-'. "#EC NOTEXT
    CONSTANTS c_format_currency_usd TYPE /nag/excel_number_format VALUE '$#,##0_-'. "#EC NOTEXT
    CONSTANTS c_format_currency_usd_simple TYPE /nag/excel_number_format VALUE '"$"#,##0.00_-'. "#EC NOTEXT
    CONSTANTS c_format_currency_simple TYPE /nag/excel_number_format VALUE '$#,##0_);($#,##0)'. "#EC NOTEXT
    CONSTANTS c_format_currency_simple_red TYPE /nag/excel_number_format VALUE '$#,##0_);[Red]($#,##0)'. "#EC NOTEXT
    CONSTANTS c_format_currency_simple2 TYPE /nag/excel_number_format VALUE '$#,##0.00_);($#,##0.00)'. "#EC NOTEXT
    CONSTANTS c_format_currency_simple_red2 TYPE /nag/excel_number_format VALUE '$#,##0.00_);[Red]($#,##0.00)'. "#EC NOTEXT
    CONSTANTS c_format_date_datetime TYPE /nag/excel_number_format VALUE 'd/m/y h:mm'. "#EC NOTEXT
    CONSTANTS c_format_date_ddmmyyyy TYPE /nag/excel_number_format VALUE 'dd/mm/yy'. "#EC NOTEXT
    CONSTANTS c_format_date_ddmmyyyydot TYPE /nag/excel_number_format VALUE 'dd\.mm\.yyyy'. "#EC NOTEXT
    CONSTANTS c_format_date_dmminus TYPE /nag/excel_number_format VALUE 'd-m'. "#EC NOTEXT
    CONSTANTS c_format_date_dmyminus TYPE /nag/excel_number_format VALUE 'd-m-y'. "#EC NOTEXT
    CONSTANTS c_format_date_dmyslash TYPE /nag/excel_number_format VALUE 'd/m/y'. "#EC NOTEXT
    CONSTANTS c_format_date_myminus TYPE /nag/excel_number_format VALUE 'm-y'. "#EC NOTEXT
    CONSTANTS c_format_date_time1 TYPE /nag/excel_number_format VALUE 'h:mm AM/PM'. "#EC NOTEXT
    CONSTANTS c_format_date_time2 TYPE /nag/excel_number_format VALUE 'h:mm:ss AM/PM'. "#EC NOTEXT
    CONSTANTS c_format_date_time3 TYPE /nag/excel_number_format VALUE 'h:mm'. "#EC NOTEXT
    CONSTANTS c_format_date_time4 TYPE /nag/excel_number_format VALUE 'h:mm:ss'. "#EC NOTEXT
    CONSTANTS c_format_date_time5 TYPE /nag/excel_number_format VALUE 'mm:ss'. "#EC NOTEXT
    CONSTANTS c_format_date_time6 TYPE /nag/excel_number_format VALUE 'h:mm:ss'. "#EC NOTEXT
    CONSTANTS c_format_date_time7 TYPE /nag/excel_number_format VALUE 'i:s.S'. "#EC NOTEXT
    CONSTANTS c_format_date_time8 TYPE /nag/excel_number_format VALUE 'h:mm:ss@'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx14 TYPE /nag/excel_number_format VALUE 'mm-dd-yy'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx15 TYPE /nag/excel_number_format VALUE 'd-mmm-yy'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx16 TYPE /nag/excel_number_format VALUE 'd-mmm'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx17 TYPE /nag/excel_number_format VALUE 'mmm-yy'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx22 TYPE /nag/excel_number_format VALUE 'm/d/yy h:mm'. "#EC NOTEXT
    CONSTANTS c_format_date_yymmdd TYPE /nag/excel_number_format VALUE 'yymmdd'. "#EC NOTEXT
    CONSTANTS c_format_date_yymmddminus TYPE /nag/excel_number_format VALUE 'yy-mm-dd'. "#EC NOTEXT
    CONSTANTS c_format_date_yymmddslash TYPE /nag/excel_number_format VALUE 'yy/mm/dd'. "#EC NOTEXT
    CONSTANTS c_format_date_yyyymmdd TYPE /nag/excel_number_format VALUE 'yyyymmdd'. "#EC NOTEXT
    CONSTANTS c_format_date_yyyymmddminus TYPE /nag/excel_number_format VALUE 'yyyy-mm-dd'. "#EC NOTEXT
    CONSTANTS c_format_date_yyyymmddslash TYPE /nag/excel_number_format VALUE 'yyyy/mm/dd'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx45 TYPE /nag/excel_number_format VALUE 'mm:ss'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx46 TYPE /nag/excel_number_format VALUE '[h]:mm:ss'. "#EC NOTEXT
    CONSTANTS c_format_date_xlsx47 TYPE /nag/excel_number_format VALUE 'mm:ss.0'. "#EC NOTEXT
    CONSTANTS c_format_general TYPE /nag/excel_number_format VALUE ''. "#EC NOTEXT
    CONSTANTS c_format_number TYPE /nag/excel_number_format VALUE '0'. "#EC NOTEXT
    CONSTANTS c_format_number_00 TYPE /nag/excel_number_format VALUE '0.00'. "#EC NOTEXT
    CONSTANTS c_format_number_comma_sep0 TYPE /nag/excel_number_format VALUE '#,##0'. "#EC NOTEXT
    CONSTANTS c_format_number_comma_sep1 TYPE /nag/excel_number_format VALUE '#,##0.00'. "#EC NOTEXT
    CONSTANTS c_format_number_comma_sep2 TYPE /nag/excel_number_format VALUE '#,##0.00_-'. "#EC NOTEXT
    CONSTANTS c_format_percentage TYPE /nag/excel_number_format VALUE '0%'. "#EC NOTEXT
    CONSTANTS c_format_percentage_00 TYPE /nag/excel_number_format VALUE '0.00%'. "#EC NOTEXT
    CONSTANTS c_format_text TYPE /nag/excel_number_format VALUE '@'. "#EC NOTEXT
    CONSTANTS c_format_fraction_1 TYPE /nag/excel_number_format VALUE '# ?/?'. "#EC NOTEXT
    CONSTANTS c_format_fraction_2 TYPE /nag/excel_number_format VALUE '# ??/??'. "#EC NOTEXT
    CONSTANTS c_format_scientific TYPE /nag/excel_number_format VALUE '0.00E+00'. "#EC NOTEXT
    CONSTANTS c_format_special_01 TYPE /nag/excel_number_format VALUE '##0.0E+0'. "#EC NOTEXT
    DATA format_code TYPE /nag/excel_number_format .
    CLASS-DATA mt_built_in_num_formats TYPE t_num_formats READ-ONLY .
    CONSTANTS c_format_xlsx37 TYPE /nag/excel_number_format VALUE '#,##0_);(#,##0)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx38 TYPE /nag/excel_number_format VALUE '#,##0_);[Red](#,##0)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx39 TYPE /nag/excel_number_format VALUE '#,##0.00_);(#,##0.00)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx40 TYPE /nag/excel_number_format VALUE '#,##0.00_);[Red](#,##0.00)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx41 TYPE /nag/excel_number_format VALUE '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx42 TYPE /nag/excel_number_format VALUE '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx43 TYPE /nag/excel_number_format VALUE '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'. "#EC NOTEXT
    CONSTANTS c_format_xlsx44 TYPE /nag/excel_number_format VALUE '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'. "#EC NOTEXT
    CONSTANTS c_format_currency_gbp_simple TYPE /nag/excel_number_format VALUE '[$£-809]#,##0.00'. "#EC NOTEXT
    CONSTANTS c_format_currency_pln_simple TYPE /nag/excel_number_format VALUE '#,##0.00\ "zł"'. "#EC NOTEXT

    CLASS-METHODS class_constructor .
    METHODS constructor .
    METHODS get_structure
      RETURNING
        VALUE(ep_number_format) TYPE /nag/excel_s_style_numfmt .
*"* protected components of class /NAG/ABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class /NAG/ABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
  PRIVATE SECTION.
    CLASS-METHODS add_format
      IMPORTING
        id   TYPE string
        code TYPE /nag/excel_number_format.
*"* private components of class /NAG/CL_EXCEL_STYLE_NUMBER_FORMAT
*"* do not include other source files here!!!
ENDCLASS.



CLASS /nag/cl_excel_style_number_format IMPLEMENTATION.

  METHOD add_format.
    DATA ls_num_format LIKE LINE OF mt_built_in_num_formats.
    ls_num_format-id                  = id.
    CREATE OBJECT ls_num_format-format.
    ls_num_format-format->format_code = code.
    INSERT ls_num_format INTO TABLE mt_built_in_num_formats.
  ENDMETHOD.

  METHOD class_constructor.

    CLEAR mt_built_in_num_formats.

    add_format( id = '1' code = /nag/cl_excel_style_number_format=>c_format_number ).               " '0'.
    add_format( id = '2' code = /nag/cl_excel_style_number_format=>c_format_number_00 ).            " '0.00'.
    add_format( id = '3' code = /nag/cl_excel_style_number_format=>c_format_number_comma_sep0 ).    " '#,##0'.
    add_format( id = '4' code = /nag/cl_excel_style_number_format=>c_format_number_comma_sep1 ).    " '#,##0.00'.
    add_format( id = '5' code = /nag/cl_excel_style_number_format=>c_format_currency_simple ).      " '$#,##0_);($#,##0)'.
    add_format( id = '6' code = /nag/cl_excel_style_number_format=>c_format_currency_simple_red ).  " '$#,##0_);[Red]($#,##0)'.
    add_format( id = '7' code = /nag/cl_excel_style_number_format=>c_format_currency_simple2 ).     " '$#,##0.00_);($#,##0.00)'.
    add_format( id = '8' code = /nag/cl_excel_style_number_format=>c_format_currency_simple_red2 ). " '$#,##0.00_);[Red]($#,##0.00)'.
    add_format( id = '9' code = /nag/cl_excel_style_number_format=>c_format_percentage ).           " '0%'.
    add_format( id = '10' code = /nag/cl_excel_style_number_format=>c_format_percentage_00 ).        " '0.00%'.
    add_format( id = '11' code = /nag/cl_excel_style_number_format=>c_format_scientific ).           " '0.00E+00'.
    add_format( id = '12' code = /nag/cl_excel_style_number_format=>c_format_fraction_1 ).           " '# ?/?'.
    add_format( id = '13' code = /nag/cl_excel_style_number_format=>c_format_fraction_2 ).           " '# ??/??'.
    add_format( id = '14' code = /nag/cl_excel_style_number_format=>c_format_date_xlsx14 ).          "'m/d/yyyy'.  <--  should have been 'mm-dd-yy' like constant in /nag/cl_excel_style_number_format
    add_format( id = '15' code = /nag/cl_excel_style_number_format=>c_format_date_xlsx15 ).          "'d-mmm-yy'.
    add_format( id = '16' code = /nag/cl_excel_style_number_format=>c_format_date_xlsx16 ).          "'d-mmm'.
    add_format( id = '17' code = /nag/cl_excel_style_number_format=>c_format_date_xlsx17 ).          "'mmm-yy'.
    add_format( id = '18' code = /nag/cl_excel_style_number_format=>c_format_date_time1 ).           " 'h:mm AM/PM'.
    add_format( id = '19' code = /nag/cl_excel_style_number_format=>c_format_date_time2 ).           " 'h:mm:ss AM/PM'.
    add_format( id = '20' code = /nag/cl_excel_style_number_format=>c_format_date_time3 ).           " 'h:mm'.
    add_format( id = '21' code = /nag/cl_excel_style_number_format=>c_format_date_time4 ).           " 'h:mm:ss'.
    add_format( id = '22' code = /nag/cl_excel_style_number_format=>c_format_date_xlsx22 ).          " 'm/d/yyyy h:mm'.


    add_format( id = '37' code = /nag/cl_excel_style_number_format=>c_format_xlsx37 ).               " '#,##0_);(#,##0)'.
    add_format( id = '38' code = /nag/cl_excel_style_number_format=>c_format_xlsx38 ).               " '#,##0_);[Red](#,##0)'.
    add_format( id = '39' code = /nag/cl_excel_style_number_format=>c_format_xlsx39 ).               " '#,##0.00_);(#,##0.00)'.
    add_format( id = '40' code = /nag/cl_excel_style_number_format=>c_format_xlsx40 ).               " '#,##0.00_);[Red](#,##0.00)'.
    add_format( id = '41' code = /nag/cl_excel_style_number_format=>c_format_xlsx41 ).               " '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'.
    add_format( id = '42' code = /nag/cl_excel_style_number_format=>c_format_xlsx42 ).               " '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'.
    add_format( id = '43' code = /nag/cl_excel_style_number_format=>c_format_xlsx43 ).               " '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'.
    add_format( id = '44' code = /nag/cl_excel_style_number_format=>c_format_xlsx44 ).               " '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'.
    add_format( id = '45' code = /nag/cl_excel_style_number_format=>c_format_date_xlsx45 ).          " 'mm:ss'.
    add_format( id = '46' code = /nag/cl_excel_style_number_format=>c_format_date_xlsx46 ).          " '[h]:mm:ss'.
    add_format( id = '47' code = /nag/cl_excel_style_number_format=>c_format_date_xlsx47 ).          "  'mm:ss.0'.
    add_format( id = '48' code = /nag/cl_excel_style_number_format=>c_format_special_01 ).           " '##0.0E+0'.
    add_format( id = '49' code = /nag/cl_excel_style_number_format=>c_format_text ).                 " '@'.

  ENDMETHOD.


  METHOD constructor.
    format_code = me->c_format_general.
  ENDMETHOD.


  METHOD get_structure.
    ep_number_format-numfmt = me->format_code.
  ENDMETHOD.
ENDCLASS.
