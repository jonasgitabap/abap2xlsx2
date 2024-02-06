INTERFACE /nag/if_excel_writer
  PUBLIC .


  METHODS write_file
    IMPORTING
      !io_excel      TYPE REF TO /nag/cl_excel
    RETURNING
      VALUE(ep_file) TYPE xstring
    RAISING
      /nag/cx_excel.
ENDINTERFACE.
