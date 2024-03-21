Testmappe
=========

.. vba:module:: Testmappe

   .. vba:vb_office_obj:: Tabelle1

      Software in Tabelle

      .. vba:vbsub:: Worksheet_SelectionChange(ByVal Target As Range, i%)
         :scope: Private

         SelectionChange

         :arg Range Target:
         :arg % i:


   .. vba:vbmodule:: main

      Dies ist ein test Modul
      wir testen den Exel parser und die Sphinx-rst-Syntax

      .. vba:vbconst:: mconst% = 10 
         :scope: Global

         zahl 10

      .. vba:vbconst:: mconst2 As Double = 10 * 10 
         :scope: Public

         zahl 100

      .. vba:vbvar:: MeineVar%
         :scope: Public

         Globaler Speicher

      .. vba:vbvar:: MeinWb As Excel.Workbooks
         :scope: Public

         Test .-namen

      .. vba:vbsub:: testis(a%, b As Integer)
         :scope: Public
         :static:

         eine test sub

         :arg % a:
         :arg Integer b:


      .. vba:vbfunc:: myf(i As Double) As String

         eine test Function

         :arg Double i:
         :returns:
         :returntype: String


      .. vba:vbfunc:: myfunc2&()

         noch ne test Function

         :returns:
         :returntype: &


      .. vba:vbsub:: xx()
         :scope: Public



   .. vba:vbclass:: TestKlasse

      Testklasse zum Syntax test
      lalala
      ---------------------

      .. vba:vbvar:: Klassenvariable%
         :scope: Public

         Public var

      .. vba:vbvar:: Klassenvariable2 As String
         :scope: Public

         noch ne Var

      .. vba:vbvar:: withvar As Application
         :scope: Public
         :withevents:

         withvar test

      .. vba:vbprop:: Wert As Variant
         :scope: Public
         :static:

         Getter für Wert
         Letter für Wert

      .. vba:vbprop:: ABCssssssssssssssssssss
         :scope: Private

         eine boese prop

      .. vba:vbsub:: klassensub()
         :scope: Public

         eine sub der Klasse TestKlasse



   .. vba:vbform:: UserForm1

      ein test Formular

      .. vba:vbsub:: UserForm_Activate()
         :scope: Private

         Private Sub Beim Öffnen



      .. vba:vbsub:: UserForm_Click()
         :scope: Private

         beim Clicken


