rst tests
=================

vba
----

.. vba:module:: Modul1

   Ein `vba:module` kann weiteren Inhalt enthalten.

.. vba:vbclass:: vbaclass

   Meine vb klasse

   .. vba:vbsub:: mysub(sddf as int,b as boolean)
      :scope: Public
      :static:

      laber rhabarber

   .. vba:vbconst:: EStatus_Pfad

      Pfad zur Excel-Liste mit Ersatzstatus

   .. vba:vbprop:: myprop

      eine prop

   .. vba:vbvar:: AuftragsNr

      Id des KA


   .. vba:vbfunc:: myfunction()
      
      :returns: nix


   .. vba:vbsub:: vbasub
      :scope: Private
      :single-line-parameter-list:

      volle signatur Private Sub(asak jjjjjjjjjjjjjjjjjjjj kkkkkkkkkkkkkkkkkkkkk lllllllllllllllllllllllllllll, öööööööööööööööööööööööööööööö , hhhhhhhhhhhhhhhhhhhhh)

      :arg ketyp kevar: text zu kevar sssssssssssssssssssssssssssssss dddddd 1234567890b1234567890b1234567890b
      :arg kevar2: text zu kevar2
      :type kevar2: int
      :arg % i: ein int
      :returns: was geben wir zurück
      :returntype: Typ von was geben wir zurück 

.. vba:vbform:: mainform

   Mein formular

   .. vba:vbsub:: vbaformsub(wert$)

      Private formsub(asak) 

.. vba:vbmodule:: EinVBModul

   .. vba:vbfunc:: vbamodform()
      
      :returns: nix


.. vba:vb_office_obj:: Tabelle2

   Ein Excel-Objekt

Python Referenz
---------------

.. py:module:: PythonModul

.. py:class:: my_py_class

   .. py:function:: a_py_function(ddd,  eeeeeeeeeeeeee)

      :param SomeClass foo2: description of parameter foo2
      :param int foo3: parameter foo2

      :single-line-parameter-list:


