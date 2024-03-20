KE_Tools Win10 V2
=================

.. vba:module:: KE_Tools Win10 V2


   .. vba:vbmodule:: Konstruktoren


      .. vba:vbfunc:: New_DB_Reader() As DB_Reader
         :scope: Public


         :returns:
         :returntype: DB_Reader


      .. vba:vbfunc:: New_KW_from_date(myday As Date) As Kalenderwoche
         :scope: Public


         :arg Date myday:
         :returns:
         :returntype: Kalenderwoche


      .. vba:vbfunc:: New_KW_from_text(mytext$) As Kalenderwoche
         :scope: Public


         :arg $ mytext:
         :returns:
         :returntype: Kalenderwoche


      .. vba:vbfunc:: New_XML_Toolbox() As XML_Toolbox
         :scope: Public


         :returns:
         :returntype: XML_Toolbox


      .. vba:vbfunc:: New_Projekt_record() As Projekt_record
         :scope: Public


         :returns:
         :returntype: Projekt_record


      .. vba:vbfunc:: New_Pos_unterpos_records(myQM_XML_Doc As QM_XML_Doc, search$) As Pos_unterpos_records
         :scope: Public


         :arg QM_XML_Doc myQM_XML_Doc:
         :arg $ search:
         :returns:
         :returntype: Pos_unterpos_records


      .. vba:vbfunc:: New_QM_XML_Doc() As QM_XML_Doc
         :scope: Public


         :returns:
         :returntype: QM_XML_Doc


   .. vba:vbclass:: XML_Toolbox


      .. vba:vbvar:: cls_xmlDoc As DOMDocument60
         :scope: Private


      .. vba:vbvar:: cls_xmlRoot As IXMLDOMElement
         :scope: Private


      .. vba:vbprop:: xmlRoot As IXMLDOMElement
         :scope: Public


      .. vba:vbprop:: xmldoc As DOMDocument60
         :scope: Public


      .. vba:vbsub:: open_Doc(ByVal XmlDateiMitPfad As String)
         :scope: Public


         :arg String XmlDateiMitPfad:


      .. vba:vbsub:: create_Doc()
         :scope: Public




      .. vba:vbsub:: save_Doc(file_name$)
         :scope: Public


         :arg $ file_name:


      .. vba:vbfunc:: get_attribute_value(base_node As IXMLDOMElement, att_name$)
         :scope: Public


         :arg IXMLDOMElement base_node:
         :arg $ att_name:


      .. vba:vbfunc:: search_for_node(base_node As IXMLDOMElement, xpathsearch_str$) As IXMLDOMElement
         :scope: Public


         :arg IXMLDOMElement base_node:
         :arg $ xpathsearch_str:
         :returns:
         :returntype: IXMLDOMElement


      .. vba:vbfunc:: search_for_nodes(base_node As IXMLDOMElement, xpathsearch_str$) As IXMLDOMNodeList
         :scope: Public


         :arg IXMLDOMElement base_node:
         :arg $ xpathsearch_str:
         :returns:
         :returntype: IXMLDOMNodeList


   .. vba:vbmodule:: XL_Tools


      .. vba:vbsub:: Abbruchmeldung(msg$)


         :arg $ msg:


      .. vba:vbfunc:: Oeffne_Excel(name$, Pfad$) As Workbook


         :arg $ name:
         :arg $ Pfad:
         :returns:
         :returntype: Workbook


      .. vba:vbfunc:: Waehle_Datei(Optional msg$ = "", Optional path$ = "", Optional filter$ = "") As Variant


         :arg $ msg:
         :returns:
         :returntype: Variant


      .. vba:vbsub:: write_header(mysheet As Worksheet, start_cell, headertxt)


         :arg Worksheet mysheet:
         :arg  start_cell:
         :arg  headertxt:


      .. vba:vbfunc:: hole_zeilen(myrange As Range) As Long


         :arg Range myrange:
         :returns:
         :returntype: Long


      .. vba:vbfunc:: FileExists(ByVal File As String) As Boolean


         :arg String File:
         :returns:
         :returntype: Boolean


   .. vba:vbmodule:: QM2XL_Tools


      .. vba:vbvar:: cls_record As record
         :scope: Private


      .. vba:vbvar:: cls_parent As QM_XML_Doc
         :scope: Private


      .. vba:vbsub:: fill_from_XML_Doc(parent_QM_XML_Doc As QM_XML_Doc)
         :scope: Public


         :arg QM_XML_Doc parent_QM_XML_Doc:


      .. vba:vbsub:: testprint2sheet(Optional myrange As Range)
         :scope: Public


         :arg Range myrange:


      .. vba:vbfunc:: value(key$) As String
         :scope: Public


         :arg $ key:
         :returns:
         :returntype: String


      .. vba:vbfunc:: items() As Variant
         :scope: Public


         :returns:
         :returntype: Variant


      .. vba:vbfunc:: keys() As Variant
         :scope: Public


         :returns:
         :returntype: Variant


   .. vba:vbclass:: Pos_unterpos_records


      .. vba:vbvar:: cls_UPos_record As record
         :scope: Private


      .. vba:vbvar:: cls_Pos_record As record
         :scope: Private


      .. vba:vbvar:: cls_pos_upos_nodes As IXMLDOMNodeList
         :scope: Private


      .. vba:vbvar:: cls_parent As QM_XML_Doc
         :scope: Private


      .. vba:vbprop:: pos_record As record
         :scope: Public


      .. vba:vbprop:: Upos_record As record
         :scope: Public


      .. vba:vbprop:: node_count As Integer
         :scope: Public


      .. vba:vbsub:: init(myQM_XML_Doc As QM_XML_Doc, search$)


         :arg QM_XML_Doc myQM_XML_Doc:
         :arg $ search:


      .. vba:vbsub:: make_record_current(id%)
         :scope: Public


         :arg % id:


      .. vba:vbsub:: testprint_cur_record2sheet(Optional myrange As Range)
         :scope: Public


         :arg Range myrange:


      .. vba:vbfunc:: cur_rec_field(typ$, key$)
         :scope: Public


         :arg $ typ:
         :arg $ key:


   .. vba:vbclass:: record


      .. vba:vbvar:: cls_record As Dictionary
         :scope: Private


      .. vba:vbprop:: record As record
         :scope: Public


      .. vba:vbsub:: fill_from_XML_Doc(myXMLnode As IXMLDOMElement)
         :scope: Public


         :arg IXMLDOMElement myXMLnode:


      .. vba:vbfunc:: count() As Integer
         :scope: Public


         :returns:
         :returntype: Integer


      .. vba:vbfunc:: items() As Variant
         :scope: Public


         :returns:
         :returntype: Variant


      .. vba:vbfunc:: keys() As Variant
         :scope: Public


         :returns:
         :returntype: Variant


      .. vba:vbfunc:: value(key$) As String
         :scope: Public


         :arg $ key:
         :returns:
         :returntype: String


      .. vba:vbsub:: testprint2sheet(headline$, Optional myrange As Range)
         :scope: Public


         :arg $ headline:
         :arg Range myrange:


   .. vba:vbclass:: QM_XML_Doc


      .. vba:vbvar:: cls_xmlDoc As DOMDocument60
         :scope: Private


      .. vba:vbvar:: cls_XML_Toolbox As XML_Toolbox
         :scope: Private


      .. vba:vbvar:: cls_Projekt_record As Projekt_record
         :scope: Private


      .. vba:vbvar:: cls_pump_records As Pos_unterpos_records
         :scope: Private


      .. vba:vbvar:: cls_dok_date As Date
         :scope: Private


      .. vba:vbvar:: cls_dok_typ$
         :scope: Private


      .. vba:vbvar:: cls_dok_rev$
         :scope: Private


      .. vba:vbvar:: cls_dok_proj_nr$
         :scope: Private


      .. vba:vbprop:: XML_Toolbox As Variant
         :scope: Public


      .. vba:vbprop:: xmlRoot As IXMLDOMElement
         :scope: Public


      .. vba:vbprop:: xmldoc As DOMDocument60
         :scope: Public


      .. vba:vbprop:: Projekt_record As Projekt_record
         :scope: Public


      .. vba:vbprop:: pump_count As Integer
         :scope: Public


      .. vba:vbprop:: Pump_records As Pos_unterpos_records
         :scope: Public


      .. vba:vbprop:: dok_date As Date
         :scope: Public


      .. vba:vbprop:: dok_typ As String
         :scope: Public


      .. vba:vbprop:: dok_rev As String
         :scope: Public


      .. vba:vbprop:: dok_proj_nr As String
         :scope: Public


      .. vba:vbsub:: open_Single_Doc(Optional default_dir$ = "", Optional ByVal fileToOpen As String = "")
         :scope: Public


         :arg $ default_dir:


      .. vba:vbfunc:: get_document_tag(tag_path$) As Variant
         :scope: Private


         :arg $ tag_path:
         :returns:
         :returntype: Variant


      .. vba:vbsub:: search_pumps()
         :scope: Private




      .. vba:vbsub:: keys2sheet(Optional myrange As Range)
         :scope: Public


         :arg Range myrange:


      .. vba:vbsub:: testprint2sheet(Optional myrange As Range)
         :scope: Public


         :arg Range myrange:


      .. vba:vbfunc:: cur_rec_field(typ$, key$)
         :scope: Public


         :arg $ typ:
         :arg $ key:


      .. vba:vbfunc:: keys(typ$) As Variant
         :scope: Public


         :arg $ typ:
         :returns:
         :returntype: Variant


   .. vba:vbclass:: DB_Reader


      .. vba:vbvar:: locAdoConnection As ADODB.Connection
         :scope: Private


      .. vba:vbvar:: locRecordset As ADODB.Recordset
         :scope: Private


      .. vba:vbprop:: rs As Recordset
         :scope: Public


      .. vba:vbprop:: Connection As ADODB.Connection
         :scope: Public


      .. vba:vbprop:: xl_recordset As Recordset
         :scope: Public


      .. vba:vbprop:: txt_recordset As Recordset
         :scope: Public


      .. vba:vbfunc:: hole_recordset(sql$) As Recordset
         :scope: Public


         :arg $ sql:
         :returns:
         :returntype: Recordset


      .. vba:vbfunc:: open_rs_retry(sql$) As Recordset
         :scope: Public


         :arg $ sql:
         :returns:
         :returntype: Recordset


      .. vba:vbfunc:: open_rs(sql$) As Recordset
         :scope: Public


         :arg $ sql:
         :returns:
         :returntype: Recordset


      .. vba:vbfunc:: sql_cmd_no_output(sql$) As Long
         :scope: Public


         :arg $ sql:
         :returns:
         :returntype: Long


      .. vba:vbsub:: recordset_2_sheet(myrange As Range, Optional myrs As Recordset, Optional clear As Boolean, Optional header As Boolean)


         :arg Range myrange:
         :arg Recordset myrs:
         :arg Boolean clear:
         :arg Boolean header:


      .. vba:vbsub:: append_recordset_2_sheet(myrange As Range, Optional myrs As Recordset)


         :arg Range myrange:
         :arg Recordset myrs:


      .. vba:vbsub:: header_2_sheet(myrange As Range, Optional myrs As Recordset)
         :scope: Public


         :arg Range myrange:
         :arg Recordset myrs:


      .. vba:vbsub:: test_output(Optional myrs As Recordset)
         :scope: Public


         :arg Recordset myrs:


      .. vba:vbfunc:: Anzahl(sql$) As Long
         :scope: Public


         :arg $ sql:
         :returns:
         :returntype: Long


      .. vba:vbsub:: Open_Informix_Connection()
         :scope: Public




      .. vba:vbsub:: Open_SQLite_Connection(db_path$)
         :scope: Public


         :arg $ db_path:


      .. vba:vbsub:: Open_Excel_Connection(path_2_workbook$)
         :scope: Public


         :arg $ path_2_workbook:


      .. vba:vbsub:: Open_Txt_Connection(data_source_dir$)


         :arg $ data_source_dir:


      .. vba:vbsub:: Class_Terminate()
         :scope: Private




      .. vba:vbfunc:: sql_cmd_with_output(sql$) As Recordset
         :scope: Public


         :arg $ sql:
         :returns:
         :returntype: Recordset


   .. vba:vbclass:: Kalenderwoche

      !!!!!!!!!!!!!! Fehler ?? letzte Tage am Jahresende werden zu  KW1 im nächsten Jahr

      .. vba:vbvar:: locWednesday As Date
         :scope: Private

         !!!!!!!!!!!!!! Fehler ?? letzte Tage am Jahresende werden zu  KW1 im nächsten Jahr

      .. vba:vbvar:: locKW%
         :scope: Private


      .. vba:vbprop:: Mittwoch As Date
         :scope: Public


      .. vba:vbprop:: KW_txt As String
         :scope: Public


      .. vba:vbprop:: KW_int As Integer
         :scope: Public


      .. vba:vbprop:: Anfang As Date
         :scope: Public


      .. vba:vbprop:: Ende As Date
         :scope: Public


      .. vba:vbfunc:: Mittwoch_der_KW(myKW_txt As String) As Date
         :scope: Public


         :arg String myKW_txt:
         :returns:
         :returntype: Date


      .. vba:vbfunc:: Mittwoch_gleiche_Woche(myday As Date) As Date
         :scope: Public


         :arg Date myday:
         :returns:
         :returntype: Date


      .. vba:vbfunc:: greater(testKW$) As Boolean
         :scope: Public


         :arg $ testKW:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: greater_eq(testKW$) As Boolean
         :scope: Public


         :arg $ testKW:
         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: KW_plus_1_as_Text(old_KW_txt$) As String
         :scope: Public


         :arg $ old_KW_txt:
         :returns:
         :returntype: String


   .. vba:vbmodule:: Datum


      .. vba:vbfunc:: odbc_xl_date(mydate As Date) As String
         :scope: Public


         :arg Date mydate:
         :returns:
         :returntype: String


      .. vba:vbfunc:: odbc_csv_datetime(mydate As Date) As String
         :scope: Public


         :arg Date mydate:
         :returns:
         :returntype: String


      .. vba:vbfunc:: odbc_csv_date(mydate As Date) As String
         :scope: Public


         :arg Date mydate:
         :returns:
         :returntype: String


      .. vba:vbfunc:: KW(tag As Date) As Integer
         :scope: Public


         :arg Date tag:
         :returns:
         :returntype: Integer


      .. vba:vbfunc:: KWstr(tag As Date) As String
         :scope: Public


         :arg Date tag:
         :returns:
         :returntype: String


   .. vba:vbmodule:: UNIPPS2Excel_Tools


      .. vba:vbconst:: f_auftragkopf = "FROM ( " & "( " & " ( " & "f_auftragkopf INNER JOIN auftragpos " & "ON f_auftragkopf.auftr_pos = auftragpos.ident_nr2 AND f_auftragkopf.auftr_nr = auftragpos.ident_nr1 " & ") " & "INNER JOIN auftragkopf ON f_auftragkopf.auftr_nr = auftragkopf.ident_nr " & ") " & "INNER JOIN kunde ON auftragkopf.kunde = kunde.ident_nr " & ") "          & "INNER JOIN adresse ON kunde.adresse = adresse.ident_nr "
         :scope: Public


      .. vba:vbconst:: f_auftragkopf_auftragkopf_auftragpos = "FROM ( " & "f_auftragkopf INNER JOIN auftragpos " & "ON f_auftragkopf.auftr_pos = auftragpos.ident_nr2 AND f_auftragkopf.auftr_nr = auftragpos.ident_nr1 " & ") " & "INNER JOIN auftragkopf ON f_auftragkopf.auftr_nr = auftragkopf.ident_nr "
         :scope: Public


      .. vba:vbconst:: auftragkopf_auftragpos_teil = "FROM ( " & "auftragkopf INNER JOIN auftragpos " & "ON auftragkopf.ident_nr = auftragpos.ident_nr1 " & ") " & "INNER JOIN teil ON auftragpos.t_tg_nr = teil.ident_nr "
         :scope: Public


      .. vba:vbfunc:: sql_ersatz_Etiketten_nur_ID(start_datum As Date)
         :scope: Public


         :arg Date start_datum:


      .. vba:vbfunc:: sql_ersatz_Etiketten_nur_ID_per_ID(auftragkopf_ident_nr As Long)
         :scope: Public


         :arg Long auftragkopf_ident_nr:


      .. vba:vbfunc:: sql_ersatz_Etiketten_per_ID(auftragkopf_ident_nr As Long)
         :scope: Public


         :arg Long auftragkopf_ident_nr:


      .. vba:vbfunc:: sql_ersatz_Etiketten(start_datum As Date)
         :scope: Public


         :arg Date start_datum:


      .. vba:vbfunc:: sql_ersatz()
         :scope: Public




      .. vba:vbfunc:: sql_offene_Pumpen()
         :scope: Public




      .. vba:vbfunc:: sql_offen_und_fgm_seit_datum(start_datum As Date)
         :scope: Public


         :arg Date start_datum:


      .. vba:vbfunc:: sql_offen_und_fgm_nach_Lieferkw_seit_Lieferkw(start_datum As Date)
         :scope: Public


         :arg Date start_datum:


      .. vba:vbfunc:: sql_reparatur()
         :scope: Public




      .. vba:vbfunc:: sql_ersatz_kumuliert()
         :scope: Public




      .. vba:vbfunc:: sql_pumpen_FA(start_datum As Date) As String
         :scope: Public


         :arg Date start_datum:
         :returns:
         :returntype: String


      .. vba:vbfunc:: sql_pumpen_KA(start_datum As Date)
         :scope: Public


         :arg Date start_datum:


      .. vba:vbfunc:: sql_pumpen_KA_fgm(start_datum As Date)
         :scope: Public


         :arg Date start_datum:


      .. vba:vbfunc:: UNIPPS_Import(sql$, target_rng As Range) As Long


         :arg $ sql:
         :arg Range target_rng:
         :returns:
         :returntype: Long


      .. vba:vbsub:: get_KW(myrange As Range)


         :arg Range myrange:

