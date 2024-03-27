DatenbankDoku
=============

.. vba:module:: DatenbankDoku


   .. vba:vbmodule:: Modul1

      .. vba:vbconst:: dbpath = "C:\Users\Etscheidt\Documents\Embarcadero\Studio\Projekte\Zoll\LieferErklaer\db\"

      .. vba:vbconst:: dbname = "LieferErklaer.accdb"

      .. vba:vbvar:: Tabellen As Collection
         :scope: Global

      .. vba:vbsub:: main()
         :scope: Public



      .. vba:vbsub:: ShowTables(ausgabetyp$)
         :scope: Public

         :arg $ ausgabetyp:


      .. vba:vbsub:: GetTables(db As Database)
         :scope: Public

         :arg Database db:


   .. vba:vbclass:: Tabelle

      .. vba:vbvar:: Name$
         :scope: Public

      .. vba:vbvar:: tdef As TableDef
         :scope: Private

      .. vba:vbvar:: felder As Collection
         :scope: Public

      .. vba:vbvar:: Indizes As Collection
         :scope: Public

      .. vba:vbvar:: LenFeldName%
         :scope: Public

      .. vba:vbvar:: LenDefault%
         :scope: Public

      .. vba:vbsub:: hole_Daten(my_tdef As TableDef)
         :scope: Public

         :arg TableDef my_tdef:


      .. vba:vbsub:: holeFelder()
         :scope: Public



      .. vba:vbsub:: holeIndizes()
         :scope: Public



      .. vba:vbsub:: Ausgabe(ausgabetyp$)
         :scope: Public

         :arg $ ausgabetyp:


      .. vba:vbsub:: AusgabeFelder(ausgabetyp$)
         :scope: Public

         :arg $ ausgabetyp:


      .. vba:vbsub:: AusgabeIndizes(ausgabetyp$)
         :scope: Public

         :arg $ ausgabetyp:


   .. vba:vbclass:: Feld

      .. vba:vbvar:: fielddef As Field
         :scope: Private

      .. vba:vbvar:: Parent As Tabelle
         :scope: Public

      .. vba:vbvar:: Name$
         :scope: Public

      .. vba:vbvar:: Default As Variant
         :scope: Public

      .. vba:vbvar:: Size As Long
         :scope: Public

      .. vba:vbvar:: erforderlich As Boolean
         :scope: Public

      .. vba:vbvar:: Inhalt$
         :scope: Public

      .. vba:vbvar:: FType%
         :scope: Public

      .. vba:vbsub:: hole_Daten(myfield As Field)
         :scope: Public

         :arg Field myfield:


      .. vba:vbsub:: Ausgabe(ausgabetyp$)
         :scope: Public

         :arg $ ausgabetyp:


      .. vba:vbsub:: AusgabeRst()
         :scope: Public



      .. vba:vbfunc:: Feldtyp(typid%)
         :scope: Private

         :arg % typid:


      .. vba:vbfunc:: FilledText(OriText$, SollLaenge%)
         :scope: Private

         :arg $ OriText:
         :arg % SollLaenge:


   .. vba:vbclass:: TabellenIndex

      .. vba:vbvar:: Indexdef As Index
         :scope: Private

      .. vba:vbvar:: Parent As Tabelle
         :scope: Public

      .. vba:vbvar:: Name$
         :scope: Public

      .. vba:vbvar:: Primary As Boolean
         :scope: Public

      .. vba:vbvar:: Required As Boolean
         :scope: Public

      .. vba:vbvar:: Unique As Boolean
         :scope: Public

      .. vba:vbvar:: Feldliste As Collection
         :scope: Public

      .. vba:vbsub:: hole_Daten(meineIndexdef As Index)
         :scope: Public

         :arg Index meineIndexdef:


      .. vba:vbsub:: Ausgabe(ausgabetyp$)
         :scope: Public

         :arg $ ausgabetyp:


      .. vba:vbsub:: AusgabeRst()
         :scope: Public



      .. vba:vbfunc:: feldnamensliste() As String
         :scope: Private

         :returns:
         :returntype: String

