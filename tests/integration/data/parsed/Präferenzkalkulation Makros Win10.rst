Präferenzkalkulation Makros Win10
=================================

.. vba:module:: Präferenzkalkulation Makros Win10
   :scope: 
   :withevents:
   :static:


   .. vba:vb_office_obj:: DieseArbeitsmappe

   .. vba:vbmodule:: Globals

      .. vba:vbconst:: mit_csv_export = False
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: xls_codemappe = "Pr
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: import_sheet_name = "import"   
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: preis_sheet_name = "Listenpr"  
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: stu_sheet_name = "Kalkulation"   
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: rs_debug_sheet_name = "rs_debug"   
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: full_header = "Ebene,Typ,zu Teil,FA,id_pos,ueb_s_nr,ds,pos_nr,verurs_art,t_tg_nr,oa,Bezchng,typ,v_besch_art,urspr_land,ausl_u_land,praeferenzkennung," & "menge,sme,faktlme_sme,lme," & "bestell_id,bestell_datum,preis,basis,pme,bme,faktlme_bme,faktbme_pme,id_lief,"                         & "lieferant,pos_menge,preis_eu,preis_n_eu,Summe_Eu,Summe_n_EU,LP je St
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: KA_doku_header = "Ebene,t_tg_nr,Bezeichnung,"                         & "Menge,Lieferant,Preis_eu,Preis_n_eu,Summe_Eu,Summe_n_EU,LP(St
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: KA_doku_header_min_col = 1
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: KA_doku_header_max_col = 12
         :scope: Public
         :withevents:
         :static:


      .. vba:vbconst:: Preis_header = "id_pos,Menge,t_tg_nr,Bezeichnung,VK brutto je St
         :scope: Public
         :withevents:
         :static:


      .. vba:vbvar:: xls_hauptmappe$
         :scope: Public
         :withevents:

      .. vba:vbvar:: data_wb As Workbook
         :scope: Public
         :withevents:

      .. vba:vbvar:: code_wb As Workbook
         :scope: Public
         :withevents:

      .. vba:vbvar:: imp_sheet As Worksheet
         :scope: Public
         :withevents:

      .. vba:vbvar:: stu_sheet As Worksheet
         :scope: Public
         :withevents:

      .. vba:vbvar:: preis_sheet As Worksheet
         :scope: Public
         :withevents:

      .. vba:vbvar:: rs_debug_sheet As Worksheet
         :scope: Public
         :withevents:

      .. vba:vbvar:: UNIPPS_dbr As DB_Reader
         :scope: Public
         :withevents:

      .. vba:vbvar:: SQLite_dbr As DB_Reader
         :scope: Public
         :withevents:

      .. vba:vbvar:: SQL_exec As SQL_Executor
         :scope: Public
         :withevents:

      .. vba:vbvar:: teile_ohne_stu As Collection
         :scope: Public
         :withevents:

      .. vba:vbvar:: Logger As Logger_cls
         :scope: Public
         :withevents:

      .. vba:vbsub:: set_globals()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: set_logger(Optional batchmode As Boolean = False)
         :scope: Public
         :withevents:
         :static:


         :arg Boolean batchmode:


   .. vba:vbmodule:: main

      .. vba:vbsub:: Btn_hole_Preise_fuer_KA_Positionen()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: Btn_KA_Analyse()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: Btn_print_doku()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: Btn_speichere_pdf()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: hole_KA_Positionen_fuer_Preisblatt(ka_id$)
         :scope: Public
         :withevents:
         :static:


         :arg $ ka_id:


      .. vba:vbsub:: start_KA_Analyse(ka_id$)
         :scope: Public
         :withevents:
         :static:


         :arg $ ka_id:


      .. vba:vbsub:: store_eu_non_eu_parts(ka_id$, berechtigte As Boolean)
         :scope: Public
         :withevents:
         :static:


         :arg $ ka_id:
         :arg Boolean berechtigte:


      .. vba:vbsub:: store_pdf(ka_id$, Optional zeigen As Boolean = True)
         :scope: Public
         :withevents:
         :static:


         :arg $ ka_id:
         :arg Boolean zeigen:


   .. vba:vbmodule:: nach_Excel

      .. vba:vbsub:: import_sheet_reset()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: Preis_sheet_reset()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: KA_doku_sheet_reset()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: write_debug_header()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: write_KA_doku_header()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: write_header(target_sheet As Worksheet, row As Long, header_liste)
         :scope: Public
         :withevents:
         :static:


         :arg Worksheet target_sheet:
         :arg Long row:
         :arg  header_liste:


      .. vba:vbsub:: DeColorCells(target_sheet As Worksheet)
         :scope: Public
         :withevents:
         :static:


         :arg Worksheet target_sheet:


      .. vba:vbsub:: DeColorColumn(target_sheet As Worksheet, mycol%)
         :scope: Public
         :withevents:
         :static:


         :arg Worksheet target_sheet:
         :arg % mycol:


      .. vba:vbsub:: ColorCells(target_sheet As Worksheet, row As Long, col_min%, col_max%, farbe$)
         :scope: Public
         :withevents:
         :static:


         :arg Worksheet target_sheet:
         :arg Long row:
         :arg % col_min:
         :arg % col_max:
         :arg $ farbe:


      .. vba:vbfunc:: level_formatiert(level)
         :scope: Public
         :withevents:
         :static:


         :arg  level:


   .. vba:vbclass:: Bestellung

      .. vba:vbvar:: bestell_id
         :scope: Public
         :withevents:

      .. vba:vbvar:: bestell_datum
         :scope: Public
         :withevents:

      .. vba:vbvar:: pme_preis As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: basis
         :scope: Public
         :withevents:

      .. vba:vbvar:: pme
         :scope: Public
         :withevents:

      .. vba:vbvar:: bme
         :scope: Public
         :withevents:

      .. vba:vbvar:: faktlme_bme
         :scope: Public
         :withevents:

      .. vba:vbvar:: faktbme_pme
         :scope: Public
         :withevents:

      .. vba:vbvar:: netto_poswert
         :scope: Public
         :withevents:

      .. vba:vbvar:: menge
         :scope: Public
         :withevents:

      .. vba:vbvar:: we_menge
         :scope: Public
         :withevents:

      .. vba:vbvar:: lieferant
         :scope: Public
         :withevents:

      .. vba:vbvar:: kurzname
         :scope: Public
         :withevents:

      .. vba:vbvar:: existiert As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: Preis_je_LME As Double
         :scope: Private
         :withevents:

      .. vba:vbvar:: last_col%
         :scope: Public
         :withevents:

      .. vba:vbvar:: rs As Recordset
         :scope: Private
         :withevents:

      .. vba:vbsub:: init(t_tg_nr$)
         :scope: Public
         :withevents:
         :static:


         :arg $ t_tg_nr:


      .. vba:vbfunc:: Berechne_Preis_je_LME_rabattiert() As Double
         :scope: Private
         :withevents:
         :static:


         :returns:
         :returntype: Double


      .. vba:vbfunc:: Berechne_Preis_je_LME_unrabattiert() As Double
         :scope: Private
         :withevents:
         :static:


         :returns:
         :returntype: Double


      .. vba:vbfunc:: STU_Pos_Preis(menge As Double, faktlme_sme As Double) As Double
         :scope: Public
         :withevents:
         :static:


         :arg Double menge:
         :arg Double faktlme_sme:
         :returns:
         :returntype: Double


      .. vba:vbsub:: write2Excel_debug(myrow As Long, start_col%)
         :scope: Public
         :withevents:
         :static:


         :arg Long myrow:
         :arg % start_col:


   .. vba:vbclass:: Kundenauftrag

      .. vba:vbvar:: ka_id$
         :scope: Public
         :withevents:

      .. vba:vbvar:: kunden_id$
         :scope: Public
         :withevents:

      .. vba:vbvar:: komm_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: stueli As Collection
         :scope: Public
         :withevents:

      .. vba:vbvar:: zu_Pos As Dictionary
         :scope: Public
         :withevents:

      .. vba:vbsub:: init(id$)
         :scope: Public
         :withevents:
         :static:


         :arg $ id:


      .. vba:vbsub:: sortiere_neu()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: hole_Listenpreise()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: hole_Kinder()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: erzeuge_Baum(Baum As STU_Baum, mit_FA As Boolean)
         :scope: Public
         :withevents:
         :static:


         :arg STU_Baum Baum:
         :arg Boolean mit_FA:


   .. vba:vbclass:: SQL_Executor

      .. vba:vbfunc:: suche_FA_zu_KAPos(id_stu$, id_pos$, rs As Recordset) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg $ id_stu:
         :arg $ id_pos:
         :arg Recordset rs:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: suche_FA_zu_Teil(t_tg_nr$, rs As Recordset) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg $ t_tg_nr:
         :arg Recordset rs:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: suche_Stueli_zu_Teil(t_tg_nr$, rs As Recordset) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg $ t_tg_nr:
         :arg Recordset rs:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: suche_Kundenauftragspositionen(ka_id$, rs As Recordset) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg $ ka_id:
         :arg Recordset rs:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: hole_Rabatt_zum_Kunden(kunden_id$, rs As Recordset) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg $ kunden_id:
         :arg Recordset rs:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: hole_Pos_zu_FA(FA_id$, rs As Recordset) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg $ FA_id:
         :arg Recordset rs:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: suche_Daten_zum_Teil(t_tg_nr$, rs As Recordset) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg $ t_tg_nr:
         :arg Recordset rs:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: suche_letzte_3_Bestellungen(t_tg_nr$, rs As Recordset) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg $ t_tg_nr:
         :arg Recordset rs:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: hole_Teile_Bezeichnung(t_tg_nr$, rs As Recordset) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg $ t_tg_nr:
         :arg Recordset rs:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: hole_recordset(sql$) As Recordset
         :scope: Public
         :withevents:
         :static:


         :arg $ sql:
         :returns:
         :returntype: Recordset


   .. vba:vbclass:: Kundenauftrags_Position

      .. vba:vbvar:: pos_typ$
         :scope: Public
         :withevents:

      .. vba:vbvar:: id_stu$
         :scope: Public
         :withevents:

      .. vba:vbvar:: t_tg_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: pos_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: menge As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: teile_daten As Teiledaten
         :scope: Public
         :withevents:

      .. vba:vbvar:: stueli As Collection
         :scope: Public
         :withevents:

      .. vba:vbvar:: id_pos$
         :scope: Public
         :withevents:

      .. vba:vbvar:: vk_preis As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: vk_rabatt As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: rabatt As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: unipps_typ$
         :scope: Public
         :withevents:

      .. vba:vbvar:: komm_nr$
         :scope: Public
         :withevents:

      .. vba:vbsub:: init(record As Fields, my_rabatt As Double)
         :scope: Public
         :withevents:
         :static:


         :arg Fields record:
         :arg Double my_rabatt:


      .. vba:vbsub:: hole_Kinder_aus_Komm_FA()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: write2Excel_Preisblatt(myrow As Long)
         :scope: Public
         :withevents:
         :static:


         :arg Long myrow:


   .. vba:vbclass:: STUELI_Position

      .. vba:vbvar:: level%
         :scope: Public
         :withevents:

      .. vba:vbvar:: menge_ueb As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: Pos_daten As Variant
         :scope: Public
         :withevents:

      .. vba:vbvar:: pos_typ$
         :scope: Public
         :withevents:

      .. vba:vbvar:: id_stu$
         :scope: Public
         :withevents:

      .. vba:vbvar:: id_pos$
         :scope: Public
         :withevents:

      .. vba:vbvar:: ueb_s_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: ds$
         :scope: Public
         :withevents:

      .. vba:vbvar:: pos_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: verurs_art$
         :scope: Public
         :withevents:

      .. vba:vbvar:: menge As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: vk_preis As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: vk_rabatt As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: rabatt As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: FA_Nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: komm_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: teile_daten As Teiledaten
         :scope: Public
         :withevents:

      .. vba:vbvar:: stueli As Collection
         :scope: Public
         :withevents:

      .. vba:vbvar:: preis_EU As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: preis_Non_EU As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: Summe_EU As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: Summe_Non_EU As Double
         :scope: Public
         :withevents:

      .. vba:vbsub:: init(meine_Pos, act_level%, act_menge_ueb As Double)
         :scope: Public
         :withevents:
         :static:


         :arg  meine_Pos:
         :arg % act_level:
         :arg Double act_menge_ueb:


      .. vba:vbsub:: berechne_Preis_der_Position()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: summiere_Preise()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: writeSTU2Excel_KA_doku(row As Long)
         :scope: Public
         :withevents:
         :static:


         :arg Long row:


      .. vba:vbsub:: writeSTU2Excel_debug(row As Long)
         :scope: Public
         :withevents:
         :static:


         :arg Long row:


      .. vba:vbsub:: writePos2Excel_KA_doku(myrow As Long)
         :scope: Public
         :withevents:
         :static:


         :arg Long myrow:


      .. vba:vbsub:: writePos2Excel_debug(myrow As Long)
         :scope: Public
         :withevents:
         :static:


         :arg Long myrow:


   .. vba:vbclass:: Teiledaten

      .. vba:vbvar:: hat_stueli As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: t_tg_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: oa%
         :scope: Public
         :withevents:

      .. vba:vbvar:: bezeichnung$
         :scope: Public
         :withevents:

      .. vba:vbvar:: unipps_typ$
         :scope: Public
         :withevents:

      .. vba:vbvar:: besch_art%
         :scope: Public
         :withevents:

      .. vba:vbvar:: urspr_land%
         :scope: Public
         :withevents:

      .. vba:vbvar:: ausl_u_land%
         :scope: Public
         :withevents:

      .. vba:vbvar:: praeferenzkennung%
         :scope: Public
         :withevents:

      .. vba:vbvar:: sme%
         :scope: Public
         :withevents:

      .. vba:vbvar:: faktlme_sme As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: lme%
         :scope: Public
         :withevents:

      .. vba:vbvar:: ist_Kaufteil As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: ist_Fremdfertigung As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: ist_Eigenfertigung As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: hat_Preis As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: preis As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: preis_EU As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: preis_Non_EU As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: Bestelldaten As Bestellung
         :scope: Public
         :withevents:

      .. vba:vbsub:: init(record As Fields)
         :scope: Public
         :withevents:
         :static:


         :arg Fields record:


   .. vba:vbclass:: FA

      .. vba:vbvar:: pos_typ$
         :scope: Public
         :withevents:

      .. vba:vbvar:: id_stu$
         :scope: Public
         :withevents:

      .. vba:vbvar:: pos_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: unipps_typ$
         :scope: Public
         :withevents:

      .. vba:vbvar:: menge As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: teile_daten As Teiledaten
         :scope: Public
         :withevents:

      .. vba:vbvar:: stueli As Collection
         :scope: Public
         :withevents:

      .. vba:vbvar:: FA_Nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: verurs_art%
         :scope: Public
         :withevents:

      .. vba:vbvar:: auftragsart%
         :scope: Public
         :withevents:

      .. vba:vbsub:: init(record As Fields)
         :scope: Public
         :withevents:
         :static:


         :arg Fields record:


      .. vba:vbsub:: init_serie(record As Fields)
         :scope: Public
         :withevents:
         :static:


         :arg Fields record:


      .. vba:vbsub:: hole_Kinder()
         :scope: Public
         :withevents:
         :static:




   .. vba:vbclass:: Teil_in_STU

      .. vba:vbvar:: pos_typ$
         :scope: Public
         :withevents:

      .. vba:vbvar:: id_stu$
         :scope: Public
         :withevents:

      .. vba:vbvar:: t_tg_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: pos_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: menge As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: teile_daten As Teiledaten
         :scope: Public
         :withevents:

      .. vba:vbvar:: stueli As Collection
         :scope: Public
         :withevents:

      .. vba:vbsub:: init(record As Fields)
         :scope: Public
         :withevents:
         :static:


         :arg Fields record:


      .. vba:vbsub:: xxxerzeuge_Baum(vater_stueli_pos As STUELI_Position)
         :scope: Public
         :withevents:
         :static:


         :arg STUELI_Position vater_stueli_pos:


   .. vba:vbclass:: FA_Pos

      .. vba:vbvar:: pos_typ$
         :scope: Public
         :withevents:

      .. vba:vbvar:: t_tg_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: pos_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: menge As Double
         :scope: Public
         :withevents:

      .. vba:vbvar:: teile_daten As Teiledaten
         :scope: Public
         :withevents:

      .. vba:vbvar:: stueli As Collection
         :scope: Public
         :withevents:

      .. vba:vbvar:: id_stu$
         :scope: Public
         :withevents:

      .. vba:vbvar:: id_pos$
         :scope: Public
         :withevents:

      .. vba:vbvar:: ueb_s_nr$
         :scope: Public
         :withevents:

      .. vba:vbvar:: ds$
         :scope: Public
         :withevents:

      .. vba:vbvar:: set_block$
         :scope: Public
         :withevents:

      .. vba:vbvar:: unipps_typ$
         :scope: Public
         :withevents:

      .. vba:vbvar:: ist_toplevel As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: hat_Kinder As Boolean
         :scope: Public
         :withevents:

      .. vba:vbsub:: init(rs As Recordset)
         :scope: Public
         :withevents:
         :static:


         :arg Recordset rs:


      .. vba:vbsub:: hole_Kinder(fa_rs As Recordset, vater_stuli_id%)
         :scope: Public
         :withevents:
         :static:


         :arg Recordset fa_rs:
         :arg % vater_stuli_id:


      .. vba:vbsub:: xxxhole_Kinder(fa_rs As Recordset, vater_stuli_id%)
         :scope: Public
         :withevents:
         :static:


         :arg Recordset fa_rs:
         :arg % vater_stuli_id:


   .. vba:vbmodule:: Suche_Kinder

      .. vba:vbfunc:: suche_Kinder_v_Serien_Teil(teil As Variant) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg Variant teil:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: suche_Kinder_in_Teile_Stu(teil As Variant) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg Variant teil:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: suche_Serien_FA(teil As Variant) As Boolean
         :scope: Public
         :withevents:
         :static:


         :arg Variant teil:
         :returns:
         :returntype: Boolean


   .. vba:vbclass:: STU_Baum

      .. vba:vbvar:: stueli As Collection
         :scope: Public
         :withevents:

      .. vba:vbsub:: init()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: summiere_Preise()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: write2Excel_KA_doku()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: write2Excel_debug()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: erzeuge_Baum(typ_spez_pos As Variant, non_type_pos As STUELI_Position, mit_FA As Boolean)
         :scope: Public
         :withevents:
         :static:


         :arg Variant typ_spez_pos:
         :arg STUELI_Position non_type_pos:
         :arg Boolean mit_FA:


   .. vba:vbmodule:: Tests

      .. vba:vbvar:: fehler_sheet As Worksheet
         :scope: Dim
         :withevents:

      .. vba:vbvar:: f_row As Long
         :scope: Dim
         :withevents:

      .. vba:vbsub:: export()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: test_KA_Analyse()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: test_store_pdf()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: test_hole_KA_Positionen_fuer_Preisblatt()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: test_hole_rabatt()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: test_Dauerlauf()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbfunc:: hole_KA_aus_UNIPPS(my_dbr As DB_Reader, rs As Recordset)
         :scope: Public
         :withevents:
         :static:


         :arg DB_Reader my_dbr:
         :arg Recordset rs:


      .. vba:vbsub:: STU_Vergleich()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: Stueli_Vergleich(t_tg_nr$, rs_stu As Recordset, rs_fa As Recordset)
         :scope: Public
         :withevents:
         :static:


         :arg $ t_tg_nr:
         :arg Recordset rs_stu:
         :arg Recordset rs_fa:


      .. vba:vbsub:: hole_FA_Stueli(rs As Recordset, stueli As Collection)
         :scope: Public
         :withevents:
         :static:


         :arg Recordset rs:
         :arg Collection stueli:


      .. vba:vbsub:: hole_Stueli_zu_Teil(rs As Recordset, stueli As Collection)
         :scope: Public
         :withevents:
         :static:


         :arg Recordset rs:
         :arg Collection stueli:


      .. vba:vbfunc:: hole_Teile_aus_UNIPPS(rs As Recordset, teile_art$, besch_art%)
         :scope: Public
         :withevents:
         :static:


         :arg Recordset rs:
         :arg $ teile_art:
         :arg % besch_art:


   .. vba:vbclass:: Logger_cls

      .. vba:vbvar:: batch_modus As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: logfile As TextStream
         :scope: Private
         :withevents:

      .. vba:vbvar:: fso As FileSystemObject
         :scope: Private
         :withevents:

      .. vba:vbsub:: init(batch_mod As Boolean)
         :scope: Public
         :withevents:
         :static:


         :arg Boolean batch_mod:


      .. vba:vbsub:: user_info(msg$, Optional level% = 0)
         :scope: Public
         :withevents:
         :static:


         :arg $ msg:
         :arg % level:


      .. vba:vbsub:: log(msg$, Optional level% = 0)
         :scope: Public
         :withevents:
         :static:


         :arg $ msg:
         :arg % level:


      .. vba:vbfunc:: space(level%) As String
         :scope: Private
         :withevents:
         :static:


         :arg % level:
         :returns:
         :returntype: String


      .. vba:vbsub:: Class_Terminate()
         :scope: Private
         :withevents:
         :static:




   .. vba:vbmodule:: csv_export

      .. vba:vbvar:: SQLiteConnection As ADODB.Connection
         :scope: Public
         :withevents:

      .. vba:vbfunc:: get_csv_file(filename$) As TextStream
         :scope: Public
         :withevents:
         :static:


         :arg $ filename:
         :returns:
         :returntype: TextStream


      .. vba:vbsub:: Open_SQLite_Connection()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: csv_out(rs As Recordset, filename$)
         :scope: Public
         :withevents:
         :static:


         :arg Recordset rs:
         :arg $ filename:


      .. vba:vbsub:: sqlite_out(rs As Recordset, tablename$)
         :scope: Public
         :withevents:
         :static:


         :arg Recordset rs:
         :arg $ tablename:


   .. vba:vbmodule:: xxxweg

      .. vba:vbsub:: xxxstore_non_eu_parts(ka_id$)
         :scope: Public
         :withevents:
         :static:


         :arg $ ka_id:

