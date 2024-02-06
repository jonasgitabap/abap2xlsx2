INTERFACE /nag/if_excel_reader
  PUBLIC .


  METHODS load_file
    IMPORTING
      !i_filename             TYPE csequence
      !i_use_alternate_zip    TYPE seoclsname DEFAULT space
      !i_from_applserver      TYPE abap_bool DEFAULT sy-batch
      !iv_/nag/cl_excel_classname TYPE clike OPTIONAL
    RETURNING
      VALUE(r_excel)          TYPE REF TO /nag/cl_excel
    RAISING
      /nag/cx_excel .
  METHODS load
    IMPORTING
      !i_excel2007            TYPE xstring
      !i_use_alternate_zip    TYPE seoclsname DEFAULT space
      !iv_/nag/cl_excel_classname TYPE clike OPTIONAL
    RETURNING
      VALUE(r_excel)          TYPE REF TO /nag/cl_excel
    RAISING
      /nag/cx_excel .
ENDINTERFACE.
