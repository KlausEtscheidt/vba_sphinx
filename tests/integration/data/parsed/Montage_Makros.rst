Montage_Makros
==============

.. vba:module:: Montage_Makros
   :scope: 
   :withevents:
   :static:


   .. vba:vbmodule:: Globals

      .. vba:vbconst:: stuecklisten_dir = "X:\KOM_ST"
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: xls_auftragsbestand_dir = "V:\Fertigung"
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: xls_auftragsbestand_file = "Neupumpen_Montageverfolgung.xlsm"
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: xls_auftragsbestand_main_sheet = "Uebersicht"   
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: xls_auftragsbestand_erl_sheet = "erledigt"   
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: xls_auftragsbestand_entf_sheet = "entfallen"   
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: xls_auftragsbestand_KA_import_sheet = "KA_UNIPPS"   
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: max_col_autofilled As Long = 13 
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: max_col_filled As Long = 20 
         :scope: Public
         :withevents:
         :static:


      .. vba:vbvar:: data_wb As Workbook
         :scope: Public
         :withevents:

      .. vba:vbvar:: main_sheet As Worksheet
         :scope: Public
         :withevents:

      .. vba:vbvar:: main_sheet_bck As Worksheet
         :scope: Public
         :withevents:

      .. vba:vbvar:: erl_sheet As Worksheet
         :scope: Public
         :withevents:

      .. vba:vbvar:: entf_sheet As Worksheet
         :scope: Public
         :withevents:

      .. vba:vbvar:: KA_imp_sheet As Worksheet
         :scope: Public
         :withevents:

      .. vba:vbvar:: had_filter As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: UNIPPS_dbr As DB_Reader
         :scope: Public
         :withevents:

      .. vba:vbvar:: ================================================================================
         :scope: Public
         :withevents:

      .. vba:vbvar:: vbmodule:
         :scope: Public
         :withevents:

      .. vba:vbvar:: Auftragsbestand
         :scope: Public
         :withevents:

      .. vba:vbvar:: ================================================================================
         :scope: Public
         :withevents:

      .. vba:vbvar:: Option
         :scope: Public
         :withevents:

      .. vba:vbvar:: Explicit
         :scope: Public
         :withevents:

      .. vba:vbsub:: men_move_Status5()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: main_Update_Auftragsbestand()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: open_ka_rs_from_unipps()
         :scope: Private
         :withevents:
         :static:




      .. vba:vbsub:: fuege_neue_FA_an()
         :scope: Private
         :withevents:
         :static:




      .. vba:vbsub:: aktualisiere_Datenbestand()
         :scope: Private
         :withevents:
         :static:




      .. vba:vbsub:: suche_stueckliste()
         :scope: Private
         :withevents:
         :static:




      .. vba:vbsub:: aktualisiere_einen_Datensatz(row As Long, record)
         :scope: Private
         :withevents:
         :static:


         :arg Long row:
         :arg  record:


      .. vba:vbsub:: fuege_einen_neue_FA_an(record)
         :scope: Private
         :withevents:
         :static:


         :arg  record:


      .. vba:vbsub:: finish()
         :scope: Private
         :withevents:
         :static:




      .. vba:vbsub:: xx()
         :scope: 
         :withevents:
         :static:




      .. vba:vbsub:: set_globals()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: check_workbook()
         :scope: Private
         :withevents:
         :static:




      .. vba:vbsub:: prepare_workbook()
         :scope: Private
         :withevents:
         :static:




   .. vba:vbmodule:: Sort_u_Format

      .. vba:vbvar:: filterArray
         :scope: Dim
         :withevents:

      .. vba:vbvar:: currentFiltRange As String
         :scope: Dim
         :withevents:

      .. vba:vbsub:: MerkeFilter()
         :scope: 
         :withevents:
         :static:




      .. vba:vbsub:: Filter_Restore()
         :scope: 
         :withevents:
         :static:




      .. vba:vbsub:: add_filter(sort_type$)
         :scope: 
         :withevents:
         :static:


         :arg $ sort_type:


      .. vba:vbsub:: sort_sheet(sort_type$)
         :scope: 
         :withevents:
         :static:


         :arg $ sort_type:


      .. vba:vbsub:: xx_sort_sheet(sort_type$)
         :scope: 
         :withevents:
         :static:


         :arg $ sort_type:


      .. vba:vbsub:: markiere_fertige()
         :scope: 
         :withevents:
         :static:




   .. vba:vbmodule:: Menues

      .. vba:vbsub:: Workbook_Open_handler()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: define_menues()
         :scope: Private
         :withevents:
         :static:




      .. vba:vbsub:: Double_click_handler(ByVal Target As Range)
         :scope: Public
         :withevents:
         :static:


         :arg Range Target:


      .. vba:vbsub:: men_reload()
         :scope: 
         :withevents:
         :static:




   .. vba:vbmodule:: Spielwiese

      .. vba:vbsub:: import2()
         :scope: Private
         :withevents:
         :static:



