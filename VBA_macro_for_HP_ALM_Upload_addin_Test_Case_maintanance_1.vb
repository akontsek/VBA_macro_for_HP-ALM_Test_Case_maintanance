Sub Rename()

'adding new sheet for your Test Cases

Sheets.Add(After:=ActiveSheet).Name = "Test_Cases_for_HP-ALM_Upload"

'declaring your arrays

Dim arraypath
arraypath = Array("\TestCase_02", "\TestCase_03", "\TestCase_you-name-it")

Dim arrayname
arrayname = Array("TestName_new_A", "TestName_new_B", "TestName_new_C")

Dim arrayinput
arrayinput = Array("bbb", "ccc", "ddd")

Dim arraydata
arraydata = Array("222", "333", "444")



'the for loop is determined by the size of the arrays, and in our case it is the first array
'in this section you have to specify with values the macro has to replace

For i = 0 To UBound(arrayinput)


Sheets("Sample_blank_TestCases").Select
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Application.CutCopyMode = False
Selection.Copy
Sheets("Test_Cases_for_HP-ALM_Upload").Select
ActiveSheet.Paste
ActiveSheet.Range("a1").End(xlDown).Offset(1, 0).Select
Cells.Replace What:="\TestCase_01", Replacement:=arraypath(i), LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Cells.Replace What:="TestName_01", Replacement:=arrayname(i), LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Cells.Replace What:="aaa", Replacement:=arrayinput(i), LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
Cells.Replace What:="111", Replacement:=arraydata(i), LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False


Next

'moves the ready block one row down

    Sheets("Test_Cases_for_HP-ALM_Upload").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Cut Destination:=Range("A2")

'copies the header into the new sheet

    Sheets("Sample_blank_TestCases").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets("Test_Cases_for_HP-ALM_Upload").Select
    Range("A1").Select
    ActiveSheet.Paste


End Sub







