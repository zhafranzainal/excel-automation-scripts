Attribute VB_Name = "Module_SplitSheetToFiles"

Sub SplitIntoMultipleFiles()

	' Variables
	Dim ws As Worksheet
	Dim wb As Workbook
	Dim startingRow As Long
	Dim rowsPerSheet As Long
	Dim totalRows As Long
	Dim totalFiles As Long
	Dim fileNumber As Long
	Dim lastRow As Long
	
	' Set worksheet
	Set ws = ThisWorkbook.Sheets("sheet1")
	' First row after header
	startingRow = 2
	' Maximum number of rows needed per file
	rowsPerSheet = 10000
	' Total number of rows in file
	totalRows = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
	' Calculate number of files needed
	totalFiles = Application.WorksheetFunction.Ceiling(totalRows / rowsPerSheet, 1)
	
	' Loop to create files
	For fileNumber = 1 To totalFiles

		' Calculate row range for each file
		lastRow = startingRow + rowsPerSheet - 1

		If lastRow > totalRows Then
			lastRow = totalRows
		End If

		' Create a new workbook
		Set wb = Workbooks.Add
		' Copy header row in new file
		ws.Rows(1).EntireRow.Copy wb.Sheets(1).Rows(1)
		' Copy rows to new file
		ws.Rows(startingRow & ":" & lastRow).EntireRow.Copy wb.Sheets(1).Rows(2)
		' Change worksheet name
		wb.Sheets(1).Name = ws.Name
		' Save file in the same workbook path as current file
		wb.SaveAs ThisWorkbook.Path & "\Table-" & fileNumber & ".xlsx"
		' Close new book
		wb.Close SaveChanges : = False
		' Update starting row for next file
		startingRow = lastRow + 1

	Next fileNumber

	MsgBox "Done! " & totalFiles & " files have been saved to " & ThisWorkbook.Path
	
End Sub