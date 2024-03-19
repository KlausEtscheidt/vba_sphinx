Montage_Makros
==============

.. vba:module:: Montage_Makros


   .. vba:vbmodule:: Globals


      .. vba:vbconst:: stuecklisten_dir = "X:\KOM_ST"
         :scope: Public


      .. vba:vbconst:: xls_auftragsbestand_dir = "V:\Fertigung"
         :scope: Public


      .. vba:vbconst:: xls_auftragsbestand_file = "Neupumpen_Montageverfolgung.xlsm"
         :scope: Public


      .. vba:vbconst:: xls_auftragsbestand_main_sheet = "Uebersicht"   
         :scope: Public


      .. vba:vbconst:: xls_auftragsbestand_erl_sheet = "erledigt"   
         :scope: Public


      .. vba:vbconst:: xls_auftragsbestand_entf_sheet = "entfallen"   
         :scope: Public


      .. vba:vbconst:: xls_auftragsbestand_KA_import_sheet = "KA_UNIPPS"   
         :scope: Public


      .. vba:vbconst:: max_col_autofilled As Long = 13 
         :scope: Public


      .. vba:vbconst:: max_col_filled As Long = 20 
         :scope: Public


      .. vba:vbvar:: data_wb As Workbook
         :scope: Public


      .. vba:vbvar:: main_sheet As Worksheet
         :scope: Public


      .. vba:vbvar:: main_sheet_bck As Worksheet
         :scope: Public


      .. vba:vbvar:: erl_sheet As Worksheet
         :scope: Public


      .. vba:vbvar:: entf_sheet As Worksheet
         :scope: Public


      .. vba:vbvar:: KA_imp_sheet As Worksheet
         :scope: Public


      .. vba:vbvar:: had_filter As Boolean
         :scope: Public


      .. vba:vbvar:: UNIPPS_dbr As DB_Reader
         :scope: Public


   .. vba:vbmodule:: Auftragsbestand


      .. vba:vbsub:: men_move_Status5()
         :scope: Public




      .. vba:vbsub:: main_Update_Auftragsbestand()
         :scope: Public




      .. vba:vbsub:: open_ka_rs_from_unipps()
         :scope: Private




      .. vba:vbsub:: fuege_neue_FA_an()
         :scope: Private




      .. vba:vbsub:: aktualisiere_Datenbestand()
         :scope: Private




      .. vba:vbsub:: suche_stueckliste()
         :scope: Private




      .. vba:vbsub:: aktualisiere_einen_Datensatz(row As Long, record)
         :scope: Private


         :arg Long row:
         :arg  record:


      .. vba:vbsub:: fuege_einen_neue_FA_an(record)
         :scope: Private


         :arg  record:


      .. vba:vbsub:: finish()
         :scope: Private




      .. vba:vbsub:: xx()




      .. vba:vbsub:: set_globals()
         :scope: Public




      .. vba:vbsub:: check_workbook()
         :scope: Private




      .. vba:vbsub:: prepare_workbook()
         :scope: Private




   .. vba:vbmodule:: Sort_u_Format


      .. vba:vbvar:: filterArray
         :scope: Dim


      .. vba:vbvar:: currentFiltRange As String
         :scope: Dim


      .. vba:vbsub:: MerkeFilter()




      .. vba:vbsub:: Filter_Restore()




      .. vba:vbsub:: add_filter(sort_type$)


         :arg $ sort_type:


      .. vba:vbsub:: sort_sheet(sort_type$)


         :arg $ sort_type:


      .. vba:vbsub:: xx_sort_sheet(sort_type$)


         :arg $ sort_type:


      .. vba:vbsub:: markiere_fertige()




   .. vba:vbmodule:: Menues


      .. vba:vbsub:: Workbook_Open_handler()
         :scope: Public




      .. vba:vbsub:: define_menues()
         :scope: Private




      .. vba:vbsub:: Double_click_handler(ByVal Target As Range)
         :scope: Public


         :arg Range Target:


      .. vba:vbsub:: men_reload()




   .. vba:vbmodule:: Spielwiese


      .. vba:vbsub:: import2()
         :scope: Private



