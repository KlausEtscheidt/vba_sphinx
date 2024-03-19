Ersatz_Auftragsverfolgung_Makros
================================

.. vba:module:: Ersatz_Auftragsverfolgung_Makros


   .. vba:vb_office_obj:: DieseArbeitsmappe


      .. vba:vbsub:: Workbook_Open()
         :scope: Private




   .. vba:vbmodule:: ExcelAusgabe

      Modul zum Lesen von Ersatzaufträgen incl der Positionen und der Lagersituation

      .. vba:vbsub:: mainFormShow()
         :scope: Public

         Modul zum Lesen von Ersatzaufträgen incl der Positionen und der Lagersituation



      .. vba:vbsub:: lies_KA_aus_Access()
         :scope: Public

         Lies Daten aus Access



      .. vba:vbsub:: WochenAusgabeNachExcel()

         Gibt die Wochenübersicht (Anzahl,wert und Anzahl Pos) der Aufträge in Excel aus



      .. vba:vbsub:: ListenAusgabeNachExcel()

         Gibt für die gewählten Zeitbereiche und den gewählten Umfang
         die Daten der Kundeaufträge in Excel aus



      .. vba:vbsub:: DetailAusgabeNachExcel(ka_id$)

         Gibt die Positionen eunes Auftrags in Excel aus

         :arg $ ka_id:


      .. vba:vbsub:: WochenAusgabeVorbereiten(Zeile As Long)

         Vorbereiten der Ausgabe nach Excel

         :arg Long Zeile:


      .. vba:vbsub:: ListenAusgabeVorbereiten(Zeile As Long)

         Vorbereiten der KA-Listen Ausgabe nach Excel

         :arg Long Zeile:


      .. vba:vbsub:: DetailAusgabeVorbereiten(Zeile As Long)

         Vorbereiten der KA-Detail-Ausgabe nach Excel

         :arg Long Zeile:


   .. vba:vbclass:: KundenAuftrag

      Id des KA

      .. vba:vbvar:: AuftragsNr As Long
         :scope: Public

         Id des KA

      .. vba:vbvar:: Termin$
         :scope: Public

         Liefertermin als Datum (immer Mittwoch der KW)

      .. vba:vbvar:: Termin_KW$
         :scope: Public


      .. vba:vbvar:: KundenName$
         :scope: Public


      .. vba:vbvar:: Wertindex As Double
         :scope: Public

         UNIPPS-NettoGesamtpreis des Auftrags geteilt durch 200.000 €

      .. vba:vbvar:: AnzPos%
         :scope: Public

         Anzahl der Positionen des Auftrags

      .. vba:vbvar:: Zahlungsbed%
         :scope: Public

         Zahlungsbedingungen (Detail s. UNIPPS)

      .. vba:vbvar:: AllesAufLager As Boolean
         :scope: Public

         Flag=True wenn alle Positionen ausreichend auf Lager

      .. vba:vbvar:: AllesAufLagerDatum As Date
         :scope: Public


      .. vba:vbvar:: AllesAuf100erLager As Boolean
         :scope: Public

         Flag=True wenn alle Positionen aus 100'er Lagerorten stammen

      .. vba:vbvar:: Status%
         :scope: Public


      .. vba:vbvar:: StatusErsatz$
         :scope: Public


      .. vba:vbvar:: StatusErsatzDatum As Date
         :scope: Public


      .. vba:vbvar:: FertigDatum As Date
         :scope: Public


      .. vba:vbvar:: KaPositionen As Collection
         :scope: Public

         Liste der Positionen des KA

      .. vba:vbvar:: locRs As ADODB.Recordset
         :scope: Private


      .. vba:vbsub:: hole_positionen()
         :scope: Public

         Alle Positionen eines KA lesen



      .. vba:vbsub:: ExcelOut(Target As Range, ByRef Zeile As Long)
         :scope: Public

         Kundenauftrag nach Excel (mit oder ohne Positionen)

         :arg Range Target:
         :arg Long Zeile:


      .. vba:vbfunc:: isKaStatusOK() As Boolean
         :scope: Public


         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: isKaInDateRange() As Boolean
         :scope: Public


         :returns:
         :returntype: Boolean


      .. vba:vbsub:: Init(rs As ADODB.Recordset)
         :scope: Public


         :arg ADODB.Recordset rs:


   .. vba:vbclass:: Auftragsposition

      Klasse zum Speichern einer Auftragsposition

      .. vba:vbvar:: loc_hatFehlbestand As Boolean
         :scope: Private

         Klasse zum Speichern einer Auftragsposition

      .. vba:vbvar:: Auftragsnummer As Long
         :scope: Public


      .. vba:vbvar:: PosNr$
         :scope: Public


      .. vba:vbvar:: t_tg_nr$
         :scope: Public


      .. vba:vbvar:: Lager_frei As Double
         :scope: Public


      .. vba:vbvar:: Lager_reserviert As Double
         :scope: Public


      .. vba:vbvar:: VorzugsLagerOrt$
         :scope: Public


      .. vba:vbvar:: Bedarf_auftrag As Double
         :scope: Public


      .. vba:vbvar:: Bedarf_pos As Double
         :scope: Public


      .. vba:vbvar:: Bedarf_dispo As Double
         :scope: Public


      .. vba:vbvar:: hatFehlbestand As Boolean
         :scope: Public


      .. vba:vbvar:: Fehlbestands_art%
         :scope: Public


      .. vba:vbvar:: lagernd_seit As Date
         :scope: Public


      .. vba:vbvar:: hatUnterpos As Boolean
         :scope: Public


      .. vba:vbsub:: Init(rs As Recordset)
         :scope: Public

         Quasi-Konstruktor: Holt alle Daten der Pos aus Recordset
         Muss als erstes nach NEW aufgerufen werden

         :arg Recordset rs:


      .. vba:vbsub:: ExcelOut(Target As Range, ByRef Zeile As Long, startcol%)
         :scope: Public

         Ausgabe KA-Position in neue Zeile

         :arg Range Target:
         :arg Long Zeile:
         :arg % startcol:


   .. vba:vbform:: mainForm


      .. vba:vbvar:: Einzel_KW_Woche%
         :scope: Public


      .. vba:vbvar:: Einzel_KW_Jahr%
         :scope: Public


      .. vba:vbsub:: AusgabeBtn_Click()
         :scope: Private




      .. vba:vbsub:: EscBtn_Click()
         :scope: Private




      .. vba:vbsub:: UNIPPSImportBtn_Click()
         :scope: Private




      .. vba:vbsub:: UserForm_Activate()
         :scope: Private




      .. vba:vbfunc:: Check_KW_Eingabe() As Boolean
         :scope: Private


         :returns:
         :returntype: Boolean


      .. vba:vbfunc:: xxxCheck_KW_Eingabe() As Boolean
         :scope: Private


         :returns:
         :returntype: Boolean


   .. vba:vbmodule:: Globals

      Hier werden globale Variable definiert und mit set_globals1 bzw set_globals2 gesetzt
      Pfad zur Excel-Liste mit Ersatzstatus

      .. vba:vbconst:: EStatus_Pfad = "V:\"
         :scope: Global

         Hier werden globale Variable definiert und mit set_globals1 bzw set_globals2 gesetzt
         Pfad zur Excel-Liste mit Ersatzstatus

      .. vba:vbconst:: EStatus_Name = "Ersatzkommissionen.xls"
         :scope: Global


      .. vba:vbconst:: DB_Pfad = "V:\Tools\Excel Makros\"
         :scope: Global

         Pfad zur access-Datenbank

      .. vba:vbconst:: DB_Name = "Ersatz_mit_Bestand.accdb"
         :scope: Global


      .. vba:vbconst:: target_sheet_name_Wochen = "Wochen"
         :scope: Global


      .. vba:vbconst:: target_sheet_name_KA_Liste = "KA_Liste"
         :scope: Global


      .. vba:vbconst:: target_sheet_name_KA_Detail = "KA_Detail"
         :scope: Global


      .. vba:vbvar:: accApp As Object
         :scope: Global

         Access-Objekt um direkt Access-Befehle zu nutzen

      .. vba:vbvar:: DbConn As ADODB.Connection
         :scope: Global

         ODBC Datenbankverbindung zu Access

      .. vba:vbvar:: force_access_read As Boolean
         :scope: Global

         Erzwingen, das jedes mal neu eingelesen wird (nur fuer Tests True)

      .. vba:vbvar:: target_sheet_Wochen As Worksheet
         :scope: Global


      .. vba:vbvar:: target_sheet_KA_Liste As Worksheet
         :scope: Global


      .. vba:vbvar:: target_sheet_KA_Detail As Worksheet
         :scope: Global


      .. vba:vbvar:: Target As Range
         :scope: Global


      .. vba:vbvar:: Mittwoch_dieser_KW As Date
         :scope: Global


      .. vba:vbvar:: KaListenGelesen As Boolean
         :scope: Global


      .. vba:vbsub:: set_globals()

         Für Modul KA_mit_Pos



   .. vba:vbmodule:: Tests


      .. vba:vbvar:: KAListe As KA_Liste
         :scope: Global


      .. vba:vbsub:: Test1()




      .. vba:vbsub:: force_read()




      .. vba:vbsub:: unforce_read()




      .. vba:vbsub:: Wochenstatus()
         :scope: Public




   .. vba:vbclass:: KA_Liste


      .. vba:vbvar:: Liste As Collection
         :scope: Public


      .. vba:vbsub:: Init()
         :scope: Public




      .. vba:vbsub:: NachExcel(ByRef Zeile As Long)
         :scope: Public


         :arg Long Zeile:


      .. vba:vbsub:: Wochenstatus(MittwochKW As Date)
         :scope: Public


         :arg Date MittwochKW:


   .. vba:vbmodule:: Common

      ODBC-Verbindung zu Access

      .. vba:vbsub:: connect_Access()
         :scope: Public

         ODBC-Verbindung zu Access



      .. vba:vbsub:: disconnect_Access()
         :scope: Public

         ODBC-Verbindung zu Access abbauen



      .. vba:vbsub:: Open_Access()
         :scope: Public

         Access als Anwendung oeffnen



      .. vba:vbfunc:: hole_recordset(sql$) As ADODB.Recordset
         :scope: Public


         :arg $ sql:
         :returns:
         :returntype: ADODB.Recordset


      .. vba:vbsub:: DropTable(tablename$)
         :scope: Public

         Entfernt tabelle aus access

         :arg $ tablename:


   .. vba:vbmodule:: Import_Unipps_Abfragen

      Teile zu den Auftragspositionen lesen

      .. vba:vbsub:: Teile_2Access()
         :scope: Public

         Teile zu den Auftragspositionen lesen



      .. vba:vbsub:: Teile_Lagerbestand_2Access()
         :scope: Public

         Lagerbestand der Teile zu den Auftragspositionen lesen



      .. vba:vbsub:: Teile_Lagerbestands_Summen_2Access()
         :scope: Public

         Lagerbestände der Teile nach t_tg_nr und auftr_nr summieren



      .. vba:vbsub:: Apos_2Access()
         :scope: Public

         Auftragspositionen lesen



      .. vba:vbsub:: Apos_Superpos_2Access()
         :scope: Public

         Auftragspositionen markieren, die Unterpositionen haben



      .. vba:vbsub:: KA_2Access()
         :scope: Public

         Kundenauftraege einlesen; neue und neu fertiggemeldete markieren



      .. vba:vbsub:: KA_AnzPos_2Access()
         :scope: Public

         Für Kundenauftraege Anzahl der Positionen bestimmen



      .. vba:vbsub:: Dispobedarfe_2Access()
         :scope: Public

         Disponierte Bestaende der Teile zu den Auftragspositionen lesen



      .. vba:vbsub:: Dispobestand_res_Lager_2Access()
         :scope: Public

         Für Teil-Auftrags-Kombination reservierte Lagerbestände lesen



      .. vba:vbsub:: Dispobestand_Summe_2Access()
         :scope: Public

         Disponierte Bedarfe aufsummieren
         Nur Positionen ohne reservierten Lagerbestand berücksichtigen



      .. vba:vbsub:: Apos_Gesamtbedarfe_2Access()
         :scope: Public

         Gesamt-Bedarfe eines Auftrages an einem Teil in Tabelle Auftragspos
         Wichtig fuer KA die ein Teil auf mehreren Positionen enthalten



      .. vba:vbsub:: Apos_Dispo_Bedarfe_2Access()
         :scope: Public

         Disponierte Bedarfe ohne Reservierung in Tabelle Auftragspos



      .. vba:vbsub:: Apos_res_Lagerbestand_2Access()
         :scope: Public

         reservierte Lagerbestaende zur Tabelle Auftragspos dazu



      .. vba:vbsub:: Apos_freier_Lagerbestand_2Access()
         :scope: Public

         freie Lagerbestaende zur Tabelle Auftragspos dazu



      .. vba:vbsub:: Apos_Lagerort_2Access()
         :scope: Public

         Lagerorte zu Auftragspositionen dazu



      .. vba:vbsub:: Apos_FehlBedarfe_2Access()
         :scope: Public

         Fehlbedarfe Bedarfe in Tabelle Auftragspos
         
         Die Prüfung findet in mehreren Stufen statt.
         Als Ergebnis wird jeweils das Flag "fehlbestand" gesetzt und die Art der Prüfung als "fehlbest_status" gesetzt.
         "fehlbest_status"=0 (default) heißt ungeprüft
         Alle Prüfungen werden nur für bisher ungeprüfte Positionen durchgeführt
         => Sobald einmal das Flag "fehlbestand" gesetzt ist (True oder False), wird es nicht mehr neu gesetzt,
         da zugleich "fehlbest_status" mit einem Wert > 0 besetzt wird.
         Da die Prüfungen aufeinander aufbauen, ist die Reihenfolge wichtig
         
         Prüfung 1: Hat die Auftragsposition Unterpositionen ?
         Ja: Es gibt keinen Fehlbestand, da dieser für die Unterpositionen geprüft wird
         fehlbest_status=1 ; fehlbestand=False
         
         Prüfung 2.1: Gibt es für die Auftragsposition einen reservierten Bestand (Lager_res>0)
         und ist der GRÖSSER als der Gesamtbedarf an diesem Teil für diesen Auftrag (Lager_res>=bedarf_auftrag)
         Grund: Ein Teil kann auf mehreren Positionen eines Auftrags vorkommen.
         Der reservierte Bestand "Lager_res" gilt immer für alle Positionen zusammen.
         "bedarf_auftrag" ist daher die Summe der Bedarfs-Mengen eines Teils im Auftrag
         Ja: Es gibt keinen Fehlbestand; der reservierte Lagerbestand ist größer als der Bedarf des Auftrags
         fehlbest_status=2 ; fehlbestand=False
         
         Prüfung 2.2: Gibt es für die Auftragsposition einen reservierten Bestand (Lager_res>0)
         und ist der KLEINER als der Gesamtbedarf an diesem Teil für diesen Auftrag (Lager_res<bedarf_auftrag)
         Ja: Es gibt einen Fehlbestand; der reservierte Lagerbestand zu klein
         fehlbest_status=91 ; fehlbestand=True
         
         Prüfung 3.1: Gibt es KEINEN freien Lagerbestand für die Auftragsposition (Lager_frei=0)
         Grund: Die Positionen mit reserviertem Bestand wurden schon geprüft (s. oben)
         Für diese Position gibt es keinen reservierten, es muss der freie Bestand reichen
         Ja: es gibt KEINEN freien Lagerbestand
         fehlbest_status=92 ; fehlbestand=True
         
         Prüfung 3.2: Gibt es freien Lagerbestand (Lager_frei=0)
         und ist der GRÖSSER als der disponierte Bedarf für die Auftragsposition (Lager_frei>=bedarf_dispo)
         Grund: der disponierte Bedarf "bedarf_dispo" enthält die Bedarfe aller Aufträge deren Termin
         kleiner oder gleich dem Termin unseres "Prüf"-Auftrags ist,
         jedoch ohne die Bedarfe, für die es reservierten Bestand gibt.
         Reservierter Bestand fließt weder in "bedarf_dispo" noch in "Lager_frei" ein
         Ja: Der freie Lagerbestand ist >= als der disponierte Bedarf => kein Fehlbestand
         fehlbest_status=3 ; fehlbestand=False
         
         Prüfung 3.3: Gibt es freien Lagerbestand (Lager_frei=0)
         und ist der KLEINER als der disponierte Bedarf für die Auftragsposition (Lager_frei<bedarf_dispo)
         Grund: wie 3.2
         Ja: Der freie Lagerbestand ist kleiner als der disponierte Bedarf =>  Fehlbestand
         fehlbest_status=93 ; fehlbestand=True



      .. vba:vbsub:: KA_FehlBedarfe_2Access()
         :scope: Public

         Flag fuer FehlBedarfe in Tabelle KA_Zusatzdaten setzen



      .. vba:vbsub:: KA_Lager100_2Access()
         :scope: Public

         Check ob alle Teile des KA in Lagerbereich 100 liegen
         Flag in Tabelle KA_Zusatzdaten setzen



      .. vba:vbsub:: KA_Ersatzstatus_2Access()
         :scope: Public

         Ersatzstatus in Tabelle KA_Zusatzdaten



   .. vba:vbmodule:: Historie

      Berechnet durch neuen Import entstande Unterschiede (Vorher-Nachher-Vergleich)
      und legt diese mit Datum in Historie_xx-Tabellen ab
      Läuft schnell, daher ohne Fortschrittsnazeige

      .. vba:vbfunc:: Access_Historie_aktualisieren()
         :scope: Public

         Berechnet durch neuen Import entstande Unterschiede (Vorher-Nachher-Vergleich)
         und legt diese mit Datum in Historie_xx-Tabellen ab
         Läuft schnell, daher ohne Fortschrittsnazeige



      .. vba:vbsub:: HistorieAuftragsposLagerndStatus()
         :scope: Public

         Auftragspositionen, die seit dem letzten Einlesen erstmals ausreichend auf Lager liegen, in Tabelle Historie_Auftragspos eintragen



      .. vba:vbsub:: HistorieKAlagerndStatus()
         :scope: Public

         Historie für KA alle Teile auf Lager aktualisieren



      .. vba:vbsub:: HistorieErsatzStatus()
         :scope: Public

         Historie für Ersatzstatus aktualisieren



   .. vba:vbmodule:: Import_Ablauf


      .. vba:vbconst:: maxSchritt = 11
         :scope: Public


      .. vba:vbsub:: Import2Access()
         :scope: Public

         Liest neuen Datenstand aus UNIPPS und Ersatzkommissionen.xls nach Access
         hauptroutine des Imports



      .. vba:vbsub:: Unipps_2Access()
         :scope: Public




      .. vba:vbfunc:: hole_letzten_Datenstand()
         :scope: Public

         Lies das Datum des letzten UNIPPS-Imports aus Access-Tabelle ProgParameter



      .. vba:vbfunc:: DB_locked() As Boolean
         :scope: Public


         :returns:
         :returntype: Boolean


      .. vba:vbsub:: schreibe_Lock(verriegeln As Boolean)
         :scope: Public


         :arg Boolean verriegeln:


      .. vba:vbfunc:: schreibe_letzten_Datenstand(datum As Date)
         :scope: Public


         :arg Date datum:


      .. vba:vbfunc:: FortschrittZeigen(Schrittnr%, maxSchritt%, Text$)
         :scope: Public


         :arg % Schrittnr:
         :arg % maxSchritt:
         :arg $ Text:


   .. vba:vbmodule:: Import_Ersatz_Excelsheet

      Modul zum Lesen des Bearbeitungsstands in der Ersatzabteilung aus Ersatzkommissionen.xls
      Es wird zunächst die kleinste und die höchste Id aller nicht gelieferten Kundenaufträge ermittelt,
      deren Liefertermin in der Vergangenheit oder bis 3 Wochen in der Zukunft liegt
      Mit diesen Id's first_KA, last_KA werden in Excel die Blätter gelesen, die diesen ID-Bereich enthalten
      Dazu werden die ersten 3 Zeichen der Id mit den ersten 3 Zeichen der Sheetnamen verglichen
      Alle Paaren aus nicht leeren Auftragsnummern und Status dieser Blätter, werden in Access gespeichert
      Darunter sind in der Regel auch Id's, die nicht aus dem Bereich first_KA, last_KA stammen
      Umgekehrt kann der Bereich Id'S enthalten, die nocht nicht in Excel eingetragen sind

      .. vba:vbsub:: lies_Ersatz_Status_aus_Excel()
         :scope: Public

         Modul zum Lesen des Bearbeitungsstands in der Ersatzabteilung aus Ersatzkommissionen.xls
         Es wird zunächst die kleinste und die höchste Id aller nicht gelieferten Kundenaufträge ermittelt,
         deren Liefertermin in der Vergangenheit oder bis 3 Wochen in der Zukunft liegt
         Mit diesen Id's first_KA, last_KA werden in Excel die Blätter gelesen, die diesen ID-Bereich enthalten
         Dazu werden die ersten 3 Zeichen der Id mit den ersten 3 Zeichen der Sheetnamen verglichen
         Alle Paaren aus nicht leeren Auftragsnummern und Status dieser Blätter, werden in Access gespeichert
         Darunter sind in der Regel auch Id's, die nicht aus dem Bereich first_KA, last_KA stammen
         Umgekehrt kann der Bereich Id'S enthalten, die nocht nicht in Excel eingetragen sind



      .. vba:vbsub:: Durchsuche_Excel_Blatt(mysheet As Worksheet, rs As Recordset)
         :scope: Private

         Speichert alle Einträge "Kundenauftragsid"/"Status" eines Excel-Sheets in Access-Tabelle

         :arg Worksheet mysheet:
         :arg Recordset rs:


      .. vba:vbfunc:: hole_sql_KA_3Wochen() As String
         :scope: Public

         ermittelt SQl um gewünschten Bereich von Aufträgen aus Access zu lesen

         :returns:
         :returntype: String


   .. vba:vbmodule:: sql_create_table

      Modul mit SQL zum Anlegen einiger Tabellen

      .. vba:vbconst:: sql_auftragsposition = "CREATE TABLE Auftragspos ( " & "id_apos INTEGER CONSTRAINT pk PRIMARY KEY," & "ident_nr1 INTEGER, " & "ident_nr2 INTEGER," & "ueb_nr INTEGER," & "pos CHAR," & "t_tg_nr CHAR," & "besch_art INTEGER, " & "Lager_ort CHAR," & "Lager100 BIT DEFAULT 0," & "Lager_frei INTEGER DEFAULT 0," & "Lager_res INTEGER DEFAULT 0, " & "bedarf_auftrag INTEGER DEFAULT 0, " & "bedarf_pos INTEGER DEFAULT 0, " & "bedarf_dispo INTEGER DEFAULT 0, " & "ist_super_pos BIT DEFAULT 0, " & "fehlbestand BIT DEFAULT -1, " & "fehlbest_status INTEGER DEFAULT 0"                                  & ");"
         :scope: Public

         Modul mit SQL zum Anlegen einiger Tabellen

      .. vba:vbconst:: sql_teile = "CREATE TABLE Teile ( " & "t_tg_nr CHAR CONSTRAINT pk PRIMARY KEY," & "v_ort_frei CHAR," & "Lager_100 BIT DEFAULT 0," & "Lagerbestand INTEGER DEFAULT 0"                                  & ");"
         :scope: Public


      .. vba:vbconst:: sql_dispobestand = "CREATE TABLE Teile_Dispobestand ( " & "id INTEGER CONSTRAINT pk PRIMARY KEY," & "t_tg_nr CHAR," & "art INTEGER, " & "datum DATETIME," & "beleg_nr INTEGER," & "beleg_pos INTEGER," & "kunde INTEGER," & "auftr_nr INTEGER DEFAULT 0," & "verurs_nr INTEGER DEFAULT 0," & "menge INTEGER DEFAULT 0," & "res_Lagerbestand INTEGER DEFAULT 0 "                                  & ");"
         :scope: Public


   .. vba:vbclass:: Wochenstatus

      Anzahl aller KA einer Woche

      .. vba:vbvar:: nKA%
         :scope: Public

         Anzahl aller KA einer Woche

      .. vba:vbvar:: nerl%
         :scope: Public

         Anzahl der erledigten KA (Status=> 4)

      .. vba:vbvar:: noffen%
         :scope: Public

         Anzahl der offenen KA (Status < 4)

      .. vba:vbvar:: nLager100%
         :scope: Public

         Anzahl der offenen KA, deren Teile alle auf 100'er lagerorten liegen

      .. vba:vbvar:: nErsatz%
         :scope: Public

         Anzahl der offenen KA, deren Teile nicht alle auf 100'er lagerorten liegen

      .. vba:vbvar:: nVersand%
         :scope: Public

         Anzahl der der offenen KA, die im Versand bereit stehen

      .. vba:vbvar:: nFehlteil%
         :scope: Public

         Anzahl der der offenen KA, für die Teile fehlen

      .. vba:vbsub:: BerechneStatus(KAListe As KA_Liste, MittwochKW As Date)
         :scope: Public


         :arg KA_Liste KAListe:
         :arg Date MittwochKW:

