Ersatz_Etiketten Win10
======================

.. vba:module:: Ersatz_Etiketten Win10


   .. vba:vbmodule:: Globals


      .. vba:vbconst:: testmode = False
         :scope: Public


      .. vba:vbconst:: test_ab = "132372"
         :scope: Public


      .. vba:vbconst:: label_type = "A" 
         :scope: Public


      .. vba:vbconst:: Seitenrand_oben = 2 
         :scope: Public


      .. vba:vbconst:: Seitenrand_unten = 0 
         :scope: Public


      .. vba:vbconst:: Anz_Etik_vert = 4 
         :scope: Public


      .. vba:vbconst:: ER_v_oben_vor = 3 
         :scope: Public


      .. vba:vbconst:: ER_v_oben_nach = 8 
         :scope: Public


      .. vba:vbconst:: ER_v_unten_vor = 9 
         :scope: Public


      .. vba:vbconst:: ER_v_unten_nach = 3 
         :scope: Public


      .. vba:vbconst:: ER_v_mitte_vor = 11 
         :scope: Public


      .. vba:vbconst:: ER_v_mitte_nach = 11.5 
         :scope: Public


      .. vba:vbconst:: h_abnr = 11 
         :scope: Public


      .. vba:vbconst:: h_ueb_pos = 14 
         :scope: Public


      .. vba:vbconst:: h_ueb_bez = 14 
         :scope: Public


      .. vba:vbconst:: h_standard = 12 
         :scope: Public


      .. vba:vbconst:: max_lines = 16               
         :scope: Public


      .. vba:vbconst:: max_print_seiten = 10   
         :scope: Public


      .. vba:vbconst:: xls_auftragsbestand_dir = "V:\Fertigung\Excel Makros"
         :scope: Public


      .. vba:vbconst:: xls_auftragsbestand_file = "Etiketten Win10.xls"
         :scope: Public


      .. vba:vbconst:: xls_auftragsbestand_main_sheet = "Import"   
         :scope: Public


      .. vba:vbconst:: xls_auftragsbestand_KA_sheet = "KA"   
         :scope: Public


      .. vba:vbconst:: xls_print_sheet = "Print"   
         :scope: Public


      .. vba:vbconst:: xls_status_sheet = "Status"   
         :scope: Public


      .. vba:vbconst:: xls_ui_sheet = "Start"   
         :scope: Public


      .. vba:vbvar:: prog_status As Status_typ
         :scope: Public


      .. vba:vbvar:: pump_mode
         :scope: Public


      .. vba:vbvar:: data_wb As Workbook
         :scope: Public


      .. vba:vbvar:: main_sheet As Worksheet
         :scope: Public


      .. vba:vbvar:: KA_sheet As Worksheet
         :scope: Public


      .. vba:vbvar:: print_sheet As Worksheet
         :scope: Public


      .. vba:vbvar:: status_sheet As Worksheet
         :scope: Public


      .. vba:vbvar:: UI_sheet As Worksheet
         :scope: Public


      .. vba:vbvar:: lines_per_page
         :scope: Public


      .. vba:vbvar:: KA_Id_max
         :scope: Public


      .. vba:vbvar:: KA_Id_min
         :scope: Public


      .. vba:vbvar:: KA_Id_liste As Long
         :scope: Public


      .. vba:vbvar:: UNIPPS_dbr As DB_Reader
         :scope: Public


      .. vba:vbsub:: set_globals()
         :scope: Public




   .. vba:vbmodule:: Menues


      .. vba:vbsub:: Workbook_Open_handler()
         :scope: Public




      .. vba:vbsub:: main_Print_multi()
         :scope: Public




      .. vba:vbsub:: main_Print_single()
         :scope: Public




      .. vba:vbsub:: main_Update_Auftragsbestand()
         :scope: Public




      .. vba:vbsub:: main_Pumpenauftrag_lesen_und_drucken()
         :scope: Public




      .. vba:vbsub:: main_Update_format()
         :scope: Public




   .. vba:vbmodule:: Spielwiese


      .. vba:vbsub:: test()




   .. vba:vbmodule:: Auftragsbestand


      .. vba:vbfunc:: get_min_KA_Id()
         :scope: Public




      .. vba:vbfunc:: get_min_KA_date()
         :scope: Public




      .. vba:vbfunc:: get_max_KA_Id()
         :scope: Public




      .. vba:vbfunc:: get_max_KA_date()
         :scope: Public




      .. vba:vbfunc:: translate(text_id, sprache)
         :scope: Public


         :arg  text_id:
         :arg  sprache:


      .. vba:vbfunc:: id_in_excel(id_2_searchfor As Long)
         :scope: Public


         :arg Long id_2_searchfor:


      .. vba:vbsub:: get_list_of_ab_ids(min_id As Long, max_id As Long)
         :scope: Public


         :arg Long min_id:
         :arg Long max_id:


      .. vba:vbsub:: get_ka_ID_only_from_unipps(start_datum As Date)
         :scope: Public


         :arg Date start_datum:


      .. vba:vbsub:: get_ka_ID_only_from_unipps_per_ID(auftragkopf_ident_nr As Long)
         :scope: Public


         :arg Long auftragkopf_ident_nr:


      .. vba:vbsub:: get_ka_with_data_from_unipps(start_datum As Date)
         :scope: Public


         :arg Date start_datum:


      .. vba:vbsub:: get_ka_with_data_from_unipps_per_ID(auftragkopf_ident_nr As Long)
         :scope: Public


         :arg Long auftragkopf_ident_nr:


      .. vba:vbfunc:: teileinfo(tg_nr, sprache, art) As Recordset
         :scope: Private


         :arg  tg_nr:
         :arg  sprache:
         :arg  art:
         :returns:
         :returntype: Recordset


      .. vba:vbsub:: fuege_Teile_Info_an()
         :scope: Public




   .. vba:vbmodule:: Formatieren


      .. vba:vbconst:: pages_2_format = 100


      .. vba:vbconst:: withlines = False


      .. vba:vbsub:: format_print_sheet()
         :scope: Public




      .. vba:vbsub:: format_print_sheet_columns()
         :scope: Public




      .. vba:vbsub:: format_print_sheet_common()
         :scope: Public




      .. vba:vbsub:: format_print_sheet_page_breaks()
         :scope: Public




      .. vba:vbsub:: print_test_page()
         :scope: Public




      .. vba:vbsub:: format_rows_for_one_label(row As Long, ER_vor, ER_nach)
         :scope: Public


         :arg Long row:
         :arg  ER_vor:
         :arg  ER_nach:


      .. vba:vbsub:: format_print_sheet_rows()
         :scope: Public




      .. vba:vbsub:: xxxformat_print_sheet_rows()
         :scope: Public




      .. vba:vbsub:: copy_page_format()
         :scope: Public




   .. vba:vbmodule:: Tools


      .. vba:vbfunc:: add_sheet(name) As Worksheet


         :arg  name:
         :returns:
         :returntype: Worksheet


      .. vba:vbsub:: del_sheet(sheet2del As Worksheet)


         :arg Worksheet sheet2del:


   .. vba:vbform:: Vorauswahl_frm


      .. vba:vbvar:: ok_pressed As Boolean
         :scope: Public


      .. vba:vbsub:: ESC_btn_Click()
         :scope: Private




      .. vba:vbsub:: OK_btn_Click()
         :scope: Private




      .. vba:vbsub:: UserForm_Activate()
         :scope: Private




      .. vba:vbsub:: UserForm_Initialize()
         :scope: Private




      .. vba:vbsub:: Update_Form_Before_Showing()




   .. vba:vbmodule:: Status


      .. vba:vbsub:: Status_lesen()
         :scope: Public




      .. vba:vbsub:: Status_speichern()
         :scope: Public




   .. vba:vbform:: Import_frm


      .. vba:vbvar:: importieren As Boolean
         :scope: Public


      .. vba:vbsub:: ESC_btn_Click()
         :scope: Private




      .. vba:vbsub:: OK_btn_Click()
         :scope: Private




      .. vba:vbsub:: UserForm_Activate()
         :scope: Private




      .. vba:vbsub:: UserForm_Initialize()
         :scope: Private




      .. vba:vbsub:: Update_Form_Before_Showing()




   .. vba:vbmodule:: Transfer_2_print_Sheet


      .. vba:vbvar:: out_row As Long
         :scope: Public


      .. vba:vbsub:: transfer_selected_ABs()
         :scope: Public




      .. vba:vbfunc:: transfer_single_AB(start_row As Long, id As Long) As Long
         :scope: Public


         :arg Long start_row:
         :arg Long id:
         :returns:
         :returntype: Long


      .. vba:vbsub:: print_attribute_with_translation(in_row As Long, in_col%, out_row As Long, out_col%, trans_id, sprache)
         :scope: Public


         :arg Long in_row:
         :arg % in_col:
         :arg Long out_row:
         :arg % out_col:
         :arg  trans_id:
         :arg  sprache:


      .. vba:vbsub:: transfer_single_label(in_row As Long, start_out_row As Long, out_col%)
         :scope: Public


         :arg Long in_row:
         :arg Long start_out_row:
         :arg % out_col:


      .. vba:vbsub:: print_preview()




      .. vba:vbsub:: print_it()




   .. vba:vbform:: multi_Auswahl_frm


      .. vba:vbvar:: ok_pressed As Boolean
         :scope: Public


      .. vba:vbsub:: ESC_btn_Click()
         :scope: Private




      .. vba:vbsub:: OK_btn_Click()
         :scope: Private




      .. vba:vbsub:: print_lb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
         :scope: Private


         :arg MSForms.ReturnBoolean Cancel:


      .. vba:vbsub:: deselect_all_Click()
         :scope: Private




      .. vba:vbsub:: deselect_one_Click()
         :scope: Private




      .. vba:vbsub:: no_print_lb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
         :scope: Private


         :arg MSForms.ReturnBoolean Cancel:


      .. vba:vbsub:: select_all_Click()
         :scope: Private




      .. vba:vbsub:: select_one_Click()
         :scope: Private




      .. vba:vbsub:: move_one_Click(source_lb As MSForms.ListBox, target_lb As MSForms.ListBox, moveall As Boolean)
         :scope: Private


         :arg MSForms.ListBox source_lb:
         :arg MSForms.ListBox target_lb:
         :arg Boolean moveall:


      .. vba:vbsub:: UserForm_Activate()
         :scope: Private




      .. vba:vbsub:: UserForm_Initialize()
         :scope: Private




      .. vba:vbsub:: Update_Form_Before_Showing()




   .. vba:vbform:: Auswahl_frm


      .. vba:vbvar:: ok_pressed As Boolean
         :scope: Public


      .. vba:vbsub:: ESC_btn_Click()
         :scope: Private




      .. vba:vbsub:: OK_btn_Click()
         :scope: Private




      .. vba:vbsub:: UserForm_Activate()
         :scope: Private




      .. vba:vbsub:: UserForm_Initialize()
         :scope: Private




      .. vba:vbsub:: Update_Form_Before_Showing()




   .. vba:vbmodule:: Const_Spalten_Namen


      .. vba:vbconst:: col_ab_nr = 1
         :scope: Public


      .. vba:vbconst:: col_erstanlage = 2
         :scope: Public


      .. vba:vbconst:: col_status = 3
         :scope: Public


      .. vba:vbconst:: col_sprache = 4
         :scope: Public


      .. vba:vbconst:: col_pos_nr = 5
         :scope: Public


      .. vba:vbconst:: col_t_tg_nr = 6
         :scope: Public


      .. vba:vbconst:: col_werkstoff = 7
         :scope: Public


      .. vba:vbconst:: col_spezifikation = 8
         :scope: Public


      .. vba:vbconst:: col_typ = 9
         :scope: Public


      .. vba:vbconst:: col_k_ident = 10
         :scope: Public


      .. vba:vbconst:: col_menge = 11
         :scope: Public


      .. vba:vbconst:: col_k_Typ = 12
         :scope: Public


      .. vba:vbconst:: col_k_Zchn_Nr = 13
         :scope: Public


      .. vba:vbconst:: start_col_teileinfo = 14
         :scope: Public


      .. vba:vbconst:: col_teil_bezeich = 14
         :scope: Public

