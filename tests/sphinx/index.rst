VBA-Aufbau
==========

VBA-Module
----------

   Ein **vba:module** entspricht ungefähr einem Pythonmodul.\

   Es fasst Code aus einer Quell-Datei (z.B Excel-Workbook) zusammen.
   Alle Folgeeinträge gehören zu diesem module.
   Es hat keine weiteren Eigenschaften ??!

Container
---------

   Ein `vba:module` kann beliebig viele `Container` beinhalten,
   von denen es vier Arten gibt:

   - Formulare `vba:form`

   - Klassen `vba:class` 

   - Excel Objekte `vba:xl_object:` (Workbooks, etc)

   - Code-Module `vba:vbmodule` 

Software-Bausteine
------------------

Jeder Container kann die folgenden, dokumentierbaren Elemente enthalten:

   - Methoden: Sub's `vba:sub` oder Functions `vba:function`
   - Konstante `vba:const` (als Klassen- oder globale Konstante)
   - Variable `vba:var` (als Klassen- oder globale Variable)
   - Properties `vba:property` (nur in Klassen)

Links
-----

Link Modul/File :vba:mod:`Modul1`

Link VBA-Modul  :vba:vbmod:`Modul1.EinVBModul`

Link Formular   :vba:vbmod:`Modul1.mainform`

Link Klasse     :vba:vbmod:`Modul1.vbaclass`

Link XL-Objekt  :vba:vbmod:`Modul1.Tabelle2`

Link Sub:       :vba:vbproc:`Modul1.vbaclass.mysub`

Link Function:  :vba:vbproc:`Modul1.vbaclass.myfunction`

Link Konstante: :vba:vbdata:`Modul1.vbaclass.EStatus_Pfad`

Link Variable:  :vba:vbdata:`Modul1.vbaclass.AuftragsNr`

Link Property:  :vba:vbdata:`Modul1.vbaclass.myprop`

Indizes
=======

* :ref:`genindex`
* :ref:`vba-procedureindex`

Inhalt
======

.. toctree::

   module_ersatz
   Testmappe
   Ersatz_Etiketten Win10


%   Ersatz_Auftragsverfolgung_Makros


% mod_ersatz_makros