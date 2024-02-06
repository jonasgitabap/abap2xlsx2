INTERFACE /nag/if_excel_converter
  PUBLIC .


  METHODS can_convert_object
    IMPORTING
      !io_object TYPE REF TO object
    RAISING
      /nag/cx_excel .
  METHODS create_fieldcatalog
    IMPORTING
      !is_option       TYPE /nag/excel_s_converter_option
      !io_object       TYPE REF TO object
      !it_table        TYPE STANDARD TABLE
    EXPORTING
      !es_layout       TYPE /nag/excel_s_converter_layo
      !et_fieldcatalog TYPE /nag/excel_t_converter_fcat
      !eo_table        TYPE REF TO data
      !et_colors       TYPE /nag/excel_t_converter_col
      !et_filter       TYPE /nag/excel_t_converter_fil
    RAISING
      /nag/cx_excel .

   METHODS get_supported_class
     RETURNING VALUE(rv_supported_class) TYPE seoclsname.
ENDINTERFACE.
