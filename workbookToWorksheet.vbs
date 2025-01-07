' Excel Ranges to MathCAD Worksheet
' Procedure:
'   - Build a spreadsheet named "input" with named ranges: eg range1, range2, ...
'   - Be sure to include a range for units named "UNITS"
'   - Populate the following list (TODO: read files automatically from directory)

Dim mathCADfiles
mathCADfiles = Split("panelPointL0",",")

'   - Build a template MathCAD Worksheet with input variables
'	  - Make copies for each desired iteration with the names of the list above.
'     - Naming Convention for those variables: range1_range2_range3 ...
'     - Optional: put those variables to the side, using them to assign better 
'                 variable names as needed.
'     - Optional: use "areas" to collapse certain checks when not applicable. 
'         - Optional: Include a message function that returns "Does not apply" 
'                     after the area if conditions met.
'     - This script will go thru and parse input variable names, 
'       and will fill in with the corresponding intersection of fields
'     - This script will look for the "UNITS" row or column and use that.
'	  - This script will look for the range with the same name as the file 
'       and include that in the intersection.
'   - Place this script in the same folder as the Spreadsheet and the Worksheets
'   - Run this script
' Notes:
'	- labelling inputs in woorksheets that don't correspond to the excel will throw an error
'	- if a range returns multiple cells, only the first cell is read
'	- 

Sub Msg(m)
	WScript.Echo(m)
End Sub

Dim fileSystemObject
Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
Dim dirScript
dirScript = fileSystemObject.GetParentFolderName(WScript.ScriptFullName)

' Get the Mathcad Prime application object:
Set mathcad = CreateObject("MathcadPrime.Application")
mathcad.Visible = true
mathcad.Activate()

' Read an Excel Spreadsheet
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(dirScript & "\input.xlsx")

' For Error Handling
Function RangeExists(R)
	Dim Test
	On Error Resume Next
	Set Test = objExcel.ActiveSheet.Range(R)
	RangeExists = Err.Number = 0
End Function

Dim splitString
Dim rangeArray(29)
Dim myAddress
Dim myValue
Dim myUnits
Dim thisCol
Dim thisRow
Dim dimCols
Dim dimRows
Dim dimCol
Dim dimRow



For k = 0 To UBound(mathCADfiles)
	' Open the worksheet:
	Set worksheet = mathcad.Open(fileSystemObject.BuildPath(dirScript,mathCADfiles(k) & ".mcdx"))
	'worksheet.SetTitle("Title from VB Script")
	' Get Inputs:
	Set inputs = worksheet.Inputs
	countInputs = inputs.Count
	' For Each Input
	For j = 0 To countInputs - 1
		thisInput = inputs.GetAliasByIndex(j)
		splitString = ""
		splitString = Split(thisInput,"_")
		'Make a list of 30 ranges
		For i = 0 To 29
			If i <= UBound(splitString) Then
				rangeArray(i) = splitString(i)
				If Not RangeExists(rangeArray(i)) Then
					rangeArray(i) = "A1:ZZ2000"
				End If
			Else
				rangeArray(i) = "A1:ZZ2000"
			End If
		Next
		If RangeExists(mathCADfiles(k)) Then
			rangeArray(29) = mathCADfiles(k)
		End If
		myValue = objExcel.Intersect(objExcel.Range(rangeArray(0)),objExcel.Range(rangeArray(1)),objExcel.Range(rangeArray(2)),objExcel.Range(rangeArray(3)),objExcel.Range(rangeArray(4)),objExcel.Range(rangeArray(5)),objExcel.Range(rangeArray(6)),objExcel.Range(rangeArray(7)),objExcel.Range(rangeArray(8)),objExcel.Range(rangeArray(9)),objExcel.Range(rangeArray(10)),objExcel.Range(rangeArray(11)),objExcel.Range(rangeArray(12)),objExcel.Range(rangeArray(13)),objExcel.Range(rangeArray(14)),objExcel.Range(rangeArray(15)),objExcel.Range(rangeArray(16)),objExcel.Range(rangeArray(17)),objExcel.Range(rangeArray(18)),objExcel.Range(rangeArray(19)),objExcel.Range(rangeArray(20)),objExcel.Range(rangeArray(21)),objExcel.Range(rangeArray(22)),objExcel.Range(rangeArray(23)),objExcel.Range(rangeArray(24)),objExcel.Range(rangeArray(25)),objExcel.Range(rangeArray(26)),objExcel.Range(rangeArray(27)),objExcel.Range(rangeArray(28)),objExcel.Range(rangeArray(29))).Cells(1,1)
		myAddress = objExcel.Intersect(objExcel.Range(rangeArray(0)),objExcel.Range(rangeArray(1)),objExcel.Range(rangeArray(2)),objExcel.Range(rangeArray(3)),objExcel.Range(rangeArray(4)),objExcel.Range(rangeArray(5)),objExcel.Range(rangeArray(6)),objExcel.Range(rangeArray(7)),objExcel.Range(rangeArray(8)),objExcel.Range(rangeArray(9)),objExcel.Range(rangeArray(10)),objExcel.Range(rangeArray(11)),objExcel.Range(rangeArray(12)),objExcel.Range(rangeArray(13)),objExcel.Range(rangeArray(14)),objExcel.Range(rangeArray(15)),objExcel.Range(rangeArray(16)),objExcel.Range(rangeArray(17)),objExcel.Range(rangeArray(18)),objExcel.Range(rangeArray(19)),objExcel.Range(rangeArray(20)),objExcel.Range(rangeArray(21)),objExcel.Range(rangeArray(22)),objExcel.Range(rangeArray(23)),objExcel.Range(rangeArray(24)),objExcel.Range(rangeArray(25)),objExcel.Range(rangeArray(26)),objExcel.Range(rangeArray(27)),objExcel.Range(rangeArray(28)),objExcel.Range(rangeArray(29))).Address
		thisCol = objExcel.Range(myAddress).Columns(0).Column
		thisRow = objExcel.Range(myAddress).Rows(0).Row
		dimCols = objExcel.Range("UNITS").Columns.Count
		dimRows = objExcel.Range("UNITS").Rows.Count
		If dimCols > 1 Then
			If dimRows > 1 Then
				Msg "error, dims field should be single row or column"
			Else
				dimRow = objExcel.Range("UNITS").Rows(0).Row
				myUnits = objExcel.Cells(dimRow+1,thisCol+1)
			End If
		Else
			dimCol = objExcel.Range("UNITS").Columns(0).Column
			myUnits = objExcel.Cells(thisRow+1,dimCol+1)
		End If
		' Change first Input Value and Units:
		Call worksheet.SetRealValue(thisInput, myValue, myUnits)
	Next
Next
objExcel.Quit
mathcad.CloseAll(1)
mathcad.Quit(2)
