DatenbankDoku
=============

.. vba:module:: DatenbankDoku
   :scope: 
   :withevents:
   :static:


   .. vba:vbmodule:: Modul1

      .. vba:vbconst:: dbpath = "C:\Users\Etscheidt\Documents\Embarcadero\Studio\Projekte\Zoll\LieferErklaer\db\"
         :scope: 
         :withevents:
         :static:


      .. vba:vbconst:: dbname = "LieferErklaer.accdb"
         :scope: 
         :withevents:
         :static:


      .. vba:vbvar:: Tabellen As Collection
         :scope: Global
         :withevents:

      .. vba:vbsub:: main()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: ShowTables(ausgabetyp$)
         :scope: Public
         :withevents:
         :static:


         :arg $ ausgabetyp:


      .. vba:vbsub:: GetTables(db As Database)
         :scope: Public
         :withevents:
         :static:


         :arg Database db:


   .. vba:vbclass:: Tabelle

      .. vba:vbvar:: Name$
         :scope: Public
         :withevents:

      .. vba:vbvar:: tdef As TableDef
         :scope: Private
         :withevents:

      .. vba:vbvar:: felder As Collection
         :scope: Public
         :withevents:

      .. vba:vbvar:: Indizes As Collection
         :scope: Public
         :withevents:

      .. vba:vbvar:: LenFeldName%
         :scope: Public
         :withevents:

      .. vba:vbvar:: LenDefault%
         :scope: Public
         :withevents:

      .. vba:vbsub:: hole_Daten(my_tdef As TableDef)
         :scope: Public
         :withevents:
         :static:


         :arg TableDef my_tdef:


      .. vba:vbsub:: holeFelder()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: holeIndizes()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbsub:: Ausgabe(ausgabetyp$)
         :scope: Public
         :withevents:
         :static:


         :arg $ ausgabetyp:


      .. vba:vbsub:: AusgabeFelder(ausgabetyp$)
         :scope: Public
         :withevents:
         :static:


         :arg $ ausgabetyp:


      .. vba:vbsub:: AusgabeIndizes(ausgabetyp$)
         :scope: Public
         :withevents:
         :static:


         :arg $ ausgabetyp:


   .. vba:vbclass:: Feld

      .. vba:vbvar:: fielddef As Field
         :scope: Private
         :withevents:

      .. vba:vbvar:: Parent As Tabelle
         :scope: Public
         :withevents:

      .. vba:vbvar:: Name$
         :scope: Public
         :withevents:

      .. vba:vbvar:: Default As Variant
         :scope: Public
         :withevents:

      .. vba:vbvar:: Size As Long
         :scope: Public
         :withevents:

      .. vba:vbvar:: erforderlich As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: Inhalt$
         :scope: Public
         :withevents:

      .. vba:vbvar:: FType%
         :scope: Public
         :withevents:

      .. vba:vbsub:: hole_Daten(myfield As Field)
         :scope: Public
         :withevents:
         :static:


         :arg Field myfield:


      .. vba:vbsub:: Ausgabe(ausgabetyp$)
         :scope: Public
         :withevents:
         :static:


         :arg $ ausgabetyp:


      .. vba:vbsub:: AusgabeRst()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbfunc:: Feldtyp(typid%)
         :scope: Private
         :withevents:
         :static:


         :arg % typid:


      .. vba:vbfunc:: FilledText(OriText$, SollLaenge%)
         :scope: Private
         :withevents:
         :static:


         :arg $ OriText:
         :arg % SollLaenge:


   .. vba:vbclass:: TabellenIndex

      .. vba:vbvar:: Indexdef As Index
         :scope: Private
         :withevents:

      .. vba:vbvar:: Parent As Tabelle
         :scope: Public
         :withevents:

      .. vba:vbvar:: Name$
         :scope: Public
         :withevents:

      .. vba:vbvar:: Primary As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: Required As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: Unique As Boolean
         :scope: Public
         :withevents:

      .. vba:vbvar:: Feldliste As Collection
         :scope: Public
         :withevents:

      .. vba:vbsub:: hole_Daten(meineIndexdef As Index)
         :scope: Public
         :withevents:
         :static:


         :arg Index meineIndexdef:


      .. vba:vbsub:: Ausgabe(ausgabetyp$)
         :scope: Public
         :withevents:
         :static:


         :arg $ ausgabetyp:


      .. vba:vbsub:: AusgabeRst()
         :scope: Public
         :withevents:
         :static:




      .. vba:vbfunc:: feldnamensliste() As String
         :scope: Private
         :withevents:
         :static:


         :returns:
         :returntype: String

