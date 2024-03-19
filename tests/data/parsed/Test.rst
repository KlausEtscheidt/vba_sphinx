Test
====

.. vba:module:: Test


   .. vba:class:: Pos_unterpos_records


      .. vba:var:: cls_UPos_record


      .. vba:var:: cls_Pos_record

         was immer

      .. vba:var:: cls_pos_upos_nodes

         alle Unterknoten die gefunden wurden

      .. vba:var:: cls_parent


      .. vba:property:: Mittwoch

         Quasi-Konstruktor der KW aus Datum:
         Mittwoch der Woche von myday als Bezugsdatum der KW setzen und Ordinalzahl der KW bestimmen und speichern
         getter docstr

      .. vba:property:: Upos_record


      .. vba:property:: node_count

         zaehle knoten

      .. vba:sub:: init(myQM_XML_Doc As QM_XML_Doc, search$)


      .. vba:sub:: testprint_cur_record2sheet(Optional myrange As Range)


      .. vba:function:: cur_rec_field(typ$, key$)


   .. vba:vbmodule:: sql_create_table

      Modul mit SQL zum Anlegen einiger Tabellen

      .. vba:const:: sql_auftragsposition

         erzeuge Tabelle Auftragspos

      .. vba:const:: sql_teile

         erzeuge Tabelle Teile

      .. vba:const:: sql_dispobestand

         erzeuge Tabelle Teile_Dispobestand

   .. vba:form:: Vorauswahl_frm

      Ein Formular_Anzeigen
      lalaal

      .. vba:var:: ok_pressed

         speichert Status des Formulars

      .. vba:sub:: ESC_btn_Click()

         Ereignis-Routine
         schee

   .. vba:xl_object:: Tabelle2

      Sollte nur Ereignis handler enthalten
      ODER
      ODER nicht

      .. vba:sub:: Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

         Doppel Klicker

      .. vba:function:: hole_recordset(sql$)

         Was macht die hier ?
         gehört nicht hier hin

   .. vba:vbmodule:: Modul1

      Modul1
      Beschreibung

      .. vba:sub:: Formular_Anzeigen()

         Formular_Anzeigen
         Beschreibung

      .. vba:sub:: Wochen_einlesen()

         Wochen_einlesen

   .. vba:xl_object:: Tabelle5

      T5

      .. vba:sub:: Worksheet_Activate()

         Worksheet_Activate

