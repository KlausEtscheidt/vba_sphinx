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


   .. vba:vbsub:: vbasub(myvar1 As Double, i%)
      :scope: Private

      Here you can describe the sub, 
      with as much of text as it doesn't bother your reader.

      :arg Double myvar1: description of myvar1
      :arg % i: description of integer variable i
      :returns: description of what is returned
      :rtype: type of what is returned


   .. vba:vbsub:: vbasub2(myvar1 As Double, myvar2 As specialtype, i%)
      :scope: Private

      Here you can describe the sub, with as much of text as it doesn't bother your reader.

      :arg Double myvar1: description of myvar1
      :arg myvar2: description of myvar2
      :type myvar2: specialtype
      :arg % i: description of integer variable i
      :returns: description of what is returned
      :rtype: type of what is returned

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
      :param int foo3: parameter foo3
      :param foo4: parameter foo4
      :type foo4: atype

      :single-line-parameter-list:


