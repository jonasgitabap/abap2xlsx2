INTERFACE /nag/if_excel_book_properties
  PUBLIC .

  TYPES tv_excel_appversion TYPE c LENGTH 7.

  DATA creator TYPE /nag/excel_creator .
  DATA lastmodifiedby TYPE /nag/excel_creator .
  DATA created TYPE timestampl .
  DATA modified TYPE timestampl .
  DATA title TYPE /nag/excel_title .
  DATA subject TYPE /nag/excel_subject .
  DATA description TYPE /nag/excel_description .
  DATA keywords TYPE /nag/excel_keywords .
  DATA category TYPE /nag/excel_category .
  DATA company TYPE /nag/excel_company .
  DATA application TYPE /nag/excel_application .
  DATA docsecurity TYPE /nag/excel_docsecurity .
  DATA scalecrop TYPE /nag/excel_scalecrop .
  DATA linksuptodate TYPE flag .
  DATA shareddoc TYPE flag .
  DATA hyperlinkschanged TYPE flag .
  DATA appversion TYPE tv_excel_appversion .

  METHODS initialize .
ENDINTERFACE.
