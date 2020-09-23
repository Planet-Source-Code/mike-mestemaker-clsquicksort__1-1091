<div align="center">

## clsQuickSort


</div>

### Description

Generic sort class. Works with any in memory structure and will sort in any order. Does this by exposing two simple to code events: isLess and SwapItems.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Mestemaker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-mestemaker.md)
**Level**          |Unknown
**User Rating**    |4.2 (165 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-mestemaker-clsquicksort__1-1091/archive/master.zip)





### Source Code

```
Option Explicit
' QuickSort class
'
' To use this class, you must do a bit of planning: First,
' in a form or other object module (not a .bas module),
' create an object like this:
'
'  Private WithEvents TestSort as clsQuickSort
'
' Next, define a list of values. This list can be
' disk-based (table) or memory-based (array).
' Regardless, this list MUST be numerically indexed
' with no gaps in the numbering. The indexing can
' start from any number and go up to any number.
'
' Then, create code for the two events defined by this
' class: isLess and swapItems. The isLess event will
' pass three variables to you: ndx1, ndx2 and Result.
' Look at element ndx1 and ndx2 in your array (or
' however you've implemented storage). If element
' ndx1 is less than element ndx2, set the Result
' variable to -1; if element ndx1 is greater than
' element ndx2, set Result to 1; else set it to 0.
'
' To sort in descending order, reverse that logic.
' i.e. If element ndx1 is less than element ndx2,
' set the Result variable to 1; if element ndx1 is
' greater than element ndx2, set Result to -1; else
' set it to 0.
'
' If the "key" of your data is of type String, you
' can use the StrComp function in your isLess function:
'    Result = StrComp(ar(ndx1), ar(ndx2))
'
' The swapItems event will pass you two variables:
' ndx1 and ndx2. Within that code, do whatever is needed
' to swap those two items within your storage area.
'
' Within your code, when you wish to sort your list,
' call the .Sort method passing it the number of the
' last element and the number of the first element.
' If you omit the first element's index, it will
' default to 1.
'
' Upon completion, the property .RunTime will contain
' the number of seconds the routine ran.
'
' Sample code that sorts 100 random numbers is listed
' below at the end of the class code.
Public Event isLess _
  (ByVal ndx1 As Long, _
  ByVal ndx2 As Long, _
  Result As Integer)
Public Event SwapItems _
  (ByVal ndx1 As Long, _
  ByVal ndx2 As Long)
Public runTime As Long
Private Function Partition _
  (ByVal lb As Long, ByVal hb As Long) As Variant
  Dim pivot As Long
  Dim Result As Integer
  Dim lbi As Long
  Dim hbi As Long
  hbi = hb
  lbi = lb
  If hb <= lb Then
    Partition = Null
    Exit Function
  End If
  If hb - lb = 1 Then
    Result = 0
    RaiseEvent isLess(lb, hb, Result)
    If Result > 0 Then
      RaiseEvent SwapItems(lb, hb)
    End If
    Partition = Null
    Exit Function
  End If
  pivot = lbi
  Do While lbi < hbi
    Result = 0
    RaiseEvent isLess(pivot, hbi, Result)
    Do While Result <= 0 And hbi > lbi
      hbi = hbi - 1
      Result = 0
      RaiseEvent isLess(pivot, hbi, Result)
    Loop
    If hbi <> pivot Then
      RaiseEvent SwapItems(lbi, hbi)
      If lbi = pivot Then pivot = hbi
    End If
    Result = 0
    RaiseEvent isLess(lbi, pivot, Result)
    Do While Result < 0 And lbi < hbi
      lbi = lbi + 1
      Result = 0
      RaiseEvent isLess(lbi, pivot, Result)
    Loop
    If lbi <> pivot Then
      RaiseEvent SwapItems(lbi, hbi)
      If pivot = hbi Then pivot = lbi
    End If
  Loop
  Partition = pivot
End Function
Private Sub SortIt _
  (ByVal lastNdx As Long, _
  Optional ByVal firstNdx As Long = 1)
  Dim pivot As Variant
  If firstNdx < lastNdx Then
    pivot = Partition(firstNdx, lastNdx)
    If Not IsNull(pivot) Then
      Call SortIt(pivot - 1, firstNdx)
      Call SortIt(lastNdx, pivot + 1)
    End If
  End If
End Sub
Public Sub Sort _
  (ByVal lastNdx As Long, _
  Optional ByVal firstNdx As Long = 1)
  Dim startTime As Long
  startTime = Timer
  SortIt lastNdx, firstNdx
  runTime = Timer - startTime
  Do While runTime < 0
    runTime = runTime + 86400
  Loop
End Sub
Private Sub Class_Initialize()
  runTime = 0
End Sub
' SAMPLE CODE:
'Private ar(100) As Long
'Private WithEvents arSort As clsQuickSort
'Private Sub arSort_isLess _
  (ByVal ndx1 As Long, ByVal ndx2 As Long, _
  Result As Integer)
'
'  If ar(ndx1) = ar(ndx2) Then
'    Result = 0
'  Elseif ar(ndx1) < ar(ndx2) then
'    Result = -1
'  Else
'    Result = 1
'  End If
'End Sub
'Private Sub arSort_SwapItems _
  (ByVal ndx1 As Long, ByVal ndx2 As Long)
'
'  Dim tmp As Long
'  tmp = ar(ndx1)
'  ar(ndx1) = ar(ndx2)
'  ar(ndx2) = tmp
'End Sub
'  Randomize
'
'  Set arSort = New clsQuickSort
'  Dim i As Long
'  For i = LBound(ar) To UBound(ar)
'    ar(i) = Int(Rnd * 100 + 1)
'  Next i
'  arSort.Sort UBound(ar), LBound(ar)
'  Debug.Print "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'  For i = LBound(ar) To UBound(ar)
'    Debug.Print ar(i)
'  Next i
'  Debug.Print "XXXXXXXXXXXXXXXXXXXXXXXXXXXX"
'  Debug.Print "Sort time = "; arSort.runTime
```

