'1 Master workbook with all the data
'1 criteria/category needed to split the master workbook
'Existing sub file if u want to paste into exisiting worksheet and use pivot table with the sub file

Sub Split()

Dim wswb As String
Dim wssh As String
Dim LR As Long 

wswb = ActiveWorkbook.Name 'Master workbook Name
wssh = ActiveSheet.Name 'Master Worksheet Name
Vcolumn = Columns("A") 'state the column that has your ccriterial/category needed to split the master file
LR = Range("A" & Rows.Count).End(xlUp).Row 'find last row where excel will copy the cells that are visible after auto filter

Columns("A").Copy 'Column A is the Column with the category to split
'copy the column needed to split the data into another worksheet
Sheets.Add
ActiveSheet.Name = "List of Categories" 'new worksheet created to list the category in the copied column
Range("A1").PasteSpecial
Columns("A").RemoveDuplicates Columns:=1, Header:=xlYes 'remove the duplicates of categories in column A, with this you will know how many unique categories
'This will tell you how many split excel files you will get at the end

vcounter = Range("A" & Rows.Count).End(xlUp).Row 'get the last row of the categories in the new worksheet

For i = 2 To vcounter
vfilter = Sheets("List of Categories").Cells(i, 1) 'get the first category to filter from new worksheet
Sheets(wssh).Activate
ActiveSheet.Columns.AutoFilter field:=Columns("A").Column, Criteria1:=vfilter 'filter the Master ws based on the vfilter from new worksheet, Criteria1 is written as Criteria'one'
Range("A1:F" & LR).SpecialCells(xlCellTypeVisible).Select 'Copy the filtered rows only, limit the columns, so it does not copy to much. too many cells to copy will delay the macro time.
Selection.Copy

'if you want to create a new file:
'workbooks.add
'range("A1").pastespecial
'If vfilter <> "" Then
'activeworkbook.saveas thisworkbook.path & "\" & vfilter & "_List.xlsx"
'Else
'ActiveWorkbook.Saveas thisworkbook.path & "\" & "empty"
'End If

'if not for exisiting sub workbook use:
Workbooks.Open ThisWorkbook.Path & "\" & vfilter & "_List.xlsx" 'open the existing sub file you have
Sheets(vfilter).Range("A1").PasteSpecial Paste:=xlPasteAll 'paste the selection into the worksheet that has the same name is your criteria
If vfilter <> "" Then
ActiveWorkbook.Save
Else
ActiveWorkbook.Save
End If

ActiveWorkbook.Close
Workbooks(wswb).Activate

Next i
ActiveSheet.ShowAllData
Sheets("List of Categories").Delete

End Sub

'if you want to work with pivot tables, you will have to run this macro once, to get data inside the sub file first
'open the sub file to get the range of the data in the pivot table.
'set the pivot table range with the exact columns but set the rows ar the way to the bottom.
'example: Range(A:F)
'arrange your fields to how u would like the pivot table to look like
'right click on your pivot table to 'pivot table options' and go to the data tab to set the pivot table to refresh everytime you open the excel.
