Testmappe
=========

.. vba:module:: Testmappe
   :scope: 
   :withevents:
   :static:


   .. vba:vb_office_obj:: Tabelle1

      .. vba:vbvar:: i%
         :scope: Public
         :withevents:

      .. vba:vbvar:: test As String
         :scope: Public
         :withevents:

      .. vba:vbvar:: x As Double
         :scope: Private
         :withevents:

      .. vba:vbvar:: y%
         :scope: Private
         :withevents:

      .. vba:vbsub:: Worksheet_SelectionChange(ByVal Target As Range, i%)
         :scope: Private
         :withevents:
         :static:

         SelectionChange

         :arg Range Target:
         :arg % i:


   .. vba:vbmodule:: main

      .. vba:vbconst:: mconst% = 10 
         :scope: Global
         :withevents:
         :static:

         zahl 10

      .. vba:vbconst:: mconst2 As Double = 10 * 10 
         :scope: Public
         :withevents:
         :static:

         zahl 100

      .. vba:vbvar:: MeineVar%
         :scope: Public
         :withevents:

      .. vba:vbvar:: MeinWb As Excel.Workbooks
         :scope: Public
         :withevents:

      .. vba:vbsub:: testis(a%, b As Integer)
         :scope: Public
         :withevents:
         :static:

         eine test sub

         :arg % a:
         :arg Integer b:


      .. vba:vbfunc:: myf(i As Double) As String
         :scope: 
         :withevents:
         :static:

         eine test Function

         :arg Double i:
         :returns:
         :returntype: String


      .. vba:vbfunc:: myfunc2&()
         :scope: 
         :withevents:
         :static:

         noch ne test Function

         :returns:
         :returntype: &


      .. vba:vbsub:: xx()
         :scope: Public
         :withevents:
         :static:




   .. vba:vbclass:: TestKlasse

      .. vba:vbvar:: Klassenvariable%
         :scope: Public
         :withevents:

      .. vba:vbvar:: Klassenvariable2 As String
         :scope: Public
         :withevents:

      .. vba:vbvar:: withvar As Application
         :scope: Public
         :withevents:

      .. vba:vbprop:: Wert As Variant
         :scope: Public
         :withevents:
         :static:

         Getter für Wert
         Letter für Wert

      .. vba:vbprop:: ABCssssssssssssssssssss
         :scope: Private
         :withevents:
         :static:

         eine boese prop

      .. vba:vbsub:: klassensub()
         :scope: Public
         :withevents:
         :static:

         eine sub der Klasse TestKlasse



   .. vba:vbform:: UserForm1

      .. vba:vbsub:: UserForm_Activate()
         :scope: Private
         :withevents:
         :static:

         Private Sub Beim Öffnen



      .. vba:vbsub:: UserForm_Click()
         :scope: Private
         :withevents:
         :static:

         beim Clicken


