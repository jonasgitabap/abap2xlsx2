CLASS /nag/cl_excel_worksheets DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

*"* public components of class /NAG/CL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PUBLIC SECTION.

    DATA active_worksheet TYPE /nag/excel_active_worksheet VALUE 1. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .
    DATA name TYPE /nag/excel_worksheets_name VALUE 'Worksheets'. "#EC NOTEXT .  .  .  .  .  .  .  .  . " .

    METHODS add
      IMPORTING
        !ip_worksheet TYPE REF TO /nag/cl_excel_worksheet .
    METHODS clear .
    METHODS constructor .
    METHODS get
      IMPORTING
        !ip_index           TYPE /nag/excel_active_worksheet
      RETURNING
        VALUE(eo_worksheet) TYPE REF TO /nag/cl_excel_worksheet .
    METHODS get_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO /nag/cl_excel_collection_iterator .
    METHODS is_empty
      RETURNING
        VALUE(is_empty) TYPE flag .
    METHODS remove
      IMPORTING
        !ip_worksheet TYPE REF TO /nag/cl_excel_worksheet .
    METHODS size
      RETURNING
        VALUE(ep_size) TYPE i .
*"* protected components of class /NAG/CL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class /NAG/CL_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  PRIVATE SECTION.

    DATA worksheets TYPE REF TO /nag/cl_excel_collection .
ENDCLASS.



CLASS /nag/cl_excel_worksheets IMPLEMENTATION.


  METHOD add.

    worksheets->add( ip_worksheet ).

  ENDMETHOD.


  METHOD clear.

    worksheets->clear( ).

  ENDMETHOD.


  METHOD constructor.

    CREATE OBJECT worksheets.

  ENDMETHOD.


  METHOD get.

    DATA lv_index TYPE i.
    lv_index = ip_index.
    eo_worksheet ?= worksheets->get( lv_index ).

  ENDMETHOD.


  METHOD get_iterator.

    eo_iterator ?= worksheets->get_iterator( ).

  ENDMETHOD.


  METHOD is_empty.

    is_empty = worksheets->is_empty( ).

  ENDMETHOD.


  METHOD remove.

    worksheets->remove( ip_worksheet ).

  ENDMETHOD.


  METHOD size.

    ep_size = worksheets->size( ).

  ENDMETHOD.
ENDCLASS.
