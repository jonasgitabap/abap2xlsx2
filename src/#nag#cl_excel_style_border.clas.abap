CLASS /nag/cl_excel_style_border DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

*"* public components of class /NAG/CL_EXCEL_STYLE_BORDER
*"* do not include other source files here!!!
    DATA border_style TYPE /nag/excel_border .
    DATA border_color TYPE /nag/excel_s_style_color .
    CONSTANTS c_border_none TYPE /nag/excel_border VALUE 'none'. "#EC NOTEXT
    CONSTANTS c_border_dashdot TYPE /nag/excel_border VALUE 'dashDot'. "#EC NOTEXT
    CONSTANTS c_border_dashdotdot TYPE /nag/excel_border VALUE 'dashDotDot'. "#EC NOTEXT
    CONSTANTS c_border_dashed TYPE /nag/excel_border VALUE 'dashed'. "#EC NOTEXT
    CONSTANTS c_border_dotted TYPE /nag/excel_border VALUE 'dotted'. "#EC NOTEXT
    CONSTANTS c_border_double TYPE /nag/excel_border VALUE 'double'. "#EC NOTEXT
    CONSTANTS c_border_hair TYPE /nag/excel_border VALUE 'hair'. "#EC NOTEXT
    CONSTANTS c_border_medium TYPE /nag/excel_border VALUE 'medium'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashdot TYPE /nag/excel_border VALUE 'mediumDashDot'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashdotdot TYPE /nag/excel_border VALUE 'mediumDashDotDot'. "#EC NOTEXT
    CONSTANTS c_border_mediumdashed TYPE /nag/excel_border VALUE 'mediumDashed'. "#EC NOTEXT
    CONSTANTS c_border_slantdashdot TYPE /nag/excel_border VALUE 'slantDashDot'. "#EC NOTEXT
    CONSTANTS c_border_thick TYPE /nag/excel_border VALUE 'thick'. "#EC NOTEXT
    CONSTANTS c_border_thin TYPE /nag/excel_border VALUE 'thin'. "#EC NOTEXT

    METHODS constructor .
*"* protected components of class /NAG/ABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
*"* protected components of class /NAG/ABAP_EXCEL_STYLE_FONT
*"* do not include other source files here!!!
  PROTECTED SECTION.
*"* private components of class /NAG/CL_EXCEL_STYLE_BORDER
*"* do not include other source files here!!!
  PRIVATE SECTION.
ENDCLASS.



CLASS /nag/cl_excel_style_border IMPLEMENTATION.


  METHOD constructor.
    border_style = /nag/cl_excel_style_border=>c_border_none.
    border_color-theme     = /nag/cl_excel_style_color=>c_theme_not_set.
    border_color-indexed   = /nag/cl_excel_style_color=>c_indexed_not_set.
  ENDMETHOD.
ENDCLASS.
