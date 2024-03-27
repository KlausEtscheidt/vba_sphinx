# Dokstrings
## Module
Mehrere Zeilen direkt am Modulanfang
Folgt danach ein Dokstring für Konst so muß dazwischen als Trenner ein normaler Kommentar mit mindestens einem % ('%)liegen
## Konst, Var, Prop, Methoden
Mehrere Zeilen vor oder ein String am Ende der Deklarations-Zeile

# Funktion des Parsers und Eingabedateien
## Ablauf
- Leerzeilen und komplette Kommentarzeilen (`'` am zeilenanfang) werden vor dem Parsen entfernt
- Ein `'%` Kommentar wird zu `'#######'`
- Die Gesamtdatei enthält VB-Module, die mit einer Headerzeile zwischen zwei `'='*60` Zeilen beginnen
- Am Dateiende steht `<EndofFile>`
- Innerhalb der Module wird alles außer
    - Statements für globale Variable
    - Statements für globale Konstante
    - Sub's und Function 
- überlesen
- Bei Sub's und Functions wird nur die Kopfzeile gelesen und der Rest bis zum End überlesen

## Einschränkungen des Parsers
- kein Inhalt der Methoden
- **Todo: exakter Abgleich mit VB-Syntax **

## Statements im Einzelnen

### argument list (arglist)
[ Optional ] [ ByVal | ByRef ] [ ParamArray ] varname [ ( ) ] [ As type ] [ = defaultvalue ]

### Sub
[sub](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)
[ Private | Public | Friend ] [ Static ] Sub name [ ( arglist ) ]


### Function
[func](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement)
[ Public | Private | Friend] [ Static ] Function name [ ( arglist ) ] [ As type ]

### properties
[ Public | Private | Friend ] [ Static ] Property Get name [ (arglist) ] [ As type ]
[ Public | Private | Friend ] [ Static ] Property Set name ( [ arglist ], reference )
[ Public | Private | Friend ] [ Static ] Property Let name ( [ arglist ], value )

### variable
[dim](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dim-statement)
[public](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/public-statement)
[private](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/private-statement)
Global keine Doku gefunden, anscheinend veraltet => wie Public 
Dim     [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ] [ , [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ]] . . .
Public  [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ] [ , [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ]] . . .
Private [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ] [ , [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ]] . . .

### subscripts
[ lower To ] upper [ , [ lower To ] upper ] . . ..

### const
[ Public | Private ] Const constname [ As type ] = expression

----

## Umsetzung

**variable** beim Komma in Einzelstatements auftrennen

**scope** statt bei allen Statements (VBA-Objekten)
[ Public | Private | Friend | Global | Dim] bei sub, func, prop, var, const
in rst als **option** wird zu keyword vor dem Namen

**static** bei sub, func, prop
in rst als **option** wird zu keyword vor dem Namen

**withevent** bei var
 [ WithEvents ]
in rst als **option** wird zu keyword vor dem Namen

**as_type** bei Func, prop, var, const
[ As type ]
in rst als **option** wird zu keyword hinter der Argumentliste
Außerdem bei Func als InfoField => ReturnType ?????
Konflikt zu Feldern im Docstring evtl konfigurierbar

**typchar** bei Func, prop, var, const
letzter Buchstabe des Namens 
bleibt beim Namen
Außerdem bei Func als InfoField => ReturnType

**subscripts**
( [ subscripts ] ) bei var hinter Namen lassen und ausgeben

**Properties**
Let, Get und Set werden zu einer Ausgabe zusammengefasst
Docstrings werden zusammengefasst
Wenn Get vorhanden von dort Typ holen typechar oder as_type 
Typ Behandlung wie oben

**arglist** bei Sub und Func
1:1 Ausgabe hinter Namen in Klammern 
zerlegen und optional als InfoField ausgeben
