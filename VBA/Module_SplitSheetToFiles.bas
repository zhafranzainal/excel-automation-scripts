Attribute VB_Name = "Module_SplitSheetToFiles"

Sub SplitIntoMultipleFiles()

	' Variables
	Dim ws As Worksheet
	Dim totalRows As Long
	Dim rowsPerFile As Long
	Dim totalFiles As Long
	Dim i As Long
	Dim startingRow As Long
	Dim endingRow As Long
	Dim newWorkbook As Workbook

	' Set Worksheet
	Set ws = ThisWorkbook.Sheets("sheet1")
	
	' Number of rows per file
	rowsPerFile = 10000
	' Get number of rows in file
	totalRows = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
	' Calculate the number of files needed
	totalFiles = Application.WorksheetFunction.Ceiling(totalRows / rowsPerFile, 1)
	' Init variables
	startingRow = 2
	
	' Loop to create files
	For i = 1 To totalFiles

		' Calculate row range for each file
		endingRow = startingRow + rowsPerFile - 1

		If endingRow > totalRows Then
			endingRow = totalRows
		End If

		' Create new book
		Set newWorkbook = Workbooks.Add
		' Copy header row in new file
		ws.Rows(1).EntireRow.Copy newWorkbook.Sheets(1).Rows(1)
		' Copy rows to new file
		ws.Rows(startingRow & ":" & endingRow).EntireRow.Copy newWorkbook.Sheets(1).Rows(2)
		' Change workseet name
		newWorkbook.Sheets(1).Name = ws.Name
		' Save file in the same workbook path as current file
		newWorkbook.SaveAs ThisWorkbook.Path & "\" & "Table-" & i & ".xlsx" '
		' Close new book
		newWorkbook.Close SaveChanges : = False
		' Update starting row for next file
		startingRow = endingRow + 1

	Next i

End Sub