CLASS /nag/cl_excel_converter_alv_grid DEFINITION
  PUBLIC
  INHERITING FROM /nag/cl_excel_converter_alv
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS /nag/if_excel_converter~can_convert_object
        REDEFINITION .
*"* public components of class /NAG/CL_EXCEL_CONVERTER_ALV_GRID
*"* do not include other source files here!!!
    METHODS /nag/if_excel_converter~create_fieldcatalog
        REDEFINITION .
    METHODS /nag/if_excel_converter~get_supported_class
        REDEFINITION .
  PROTECTED SECTION.
*"* protected components of class /NAG/CL_EXCEL_CONVERTER_ALV_GRID
*"* do not include other source files here!!!
*"* private components of class /NAG/CL_EXCEL_CONVERTER_ALV_GRID
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS /nag/cl_excel_converter_alv_grid IMPLEMENTATION.


  METHOD /nag/if_excel_converter~get_supported_class.
    rv_supported_class = 'CL_GUI_ALV_GRID'.
  ENDMETHOD.

  METHOD /nag/if_excel_converter~can_convert_object.
    DATA: lo_alv TYPE REF TO cl_gui_alv_grid.

    TRY.
        lo_alv ?= io_object.
      CATCH cx_sy_move_cast_error .
        RAISE EXCEPTION TYPE /nag/cx_excel.
    ENDTRY.

  ENDMETHOD.


  METHOD /nag/if_excel_converter~create_fieldcatalog.
    DATA: lo_alv TYPE REF TO cl_gui_alv_grid.

    /nag/if_excel_converter~can_convert_object( io_object = io_object ).

    ws_option = is_option.

    lo_alv ?= io_object.

    CLEAR: es_layout,
           et_fieldcatalog.

    IF lo_alv IS BOUND.
      lo_alv->get_frontend_fieldcatalog( IMPORTING et_fieldcatalog = wt_fcat ).
      lo_alv->get_frontend_layout( IMPORTING es_layout = ws_layo ).
      lo_alv->get_sort_criteria( IMPORTING et_sort = wt_sort ) .
      lo_alv->get_filter_criteria( IMPORTING et_filter = wt_filt ) .

      apply_sort( EXPORTING it_table = it_table
                  IMPORTING eo_table = eo_table ) .

      get_color( EXPORTING io_table    = eo_table
                 IMPORTING et_colors   = et_colors ) .

      get_filter( IMPORTING et_filter  = et_filter
                  CHANGING  xo_table   = eo_table  ) .

      update_catalog( CHANGING  cs_layout       = es_layout
                                ct_fieldcatalog = et_fieldcatalog ).
    ENDIF.
  ENDMETHOD.
ENDCLASS.
