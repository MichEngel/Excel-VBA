‘1 Master workbook with all the data
‘1 criteria needed to split the master workbook
‘Existing sub file if u want to paste into existing worksheet and use pivot table with the sub file
‘2 worksheets in 1 master workbook

Option Compare Text ‘make uppercase and lowercase identical

Sub Split()

Dim wswb, subwbpath, subwbname, wssh1, wssh2 As String
Dim LR As Long

wswb = "Desktop Split.xlsm" 'Master workbook Name
wssh1 = "SHS 1" 'SHS Worksheet Name (first WS the macro will run)
wssh2 = "NNA 1" 'NNA Worksheet Name (second WS the macro will run)
Vcolumn = Columns("E") 'Column with the criteria to split the master wb
LR = Range("A" & Rows.Count).End(xlUp).Row 'last row & column to copy to new WB

'SHS sheet macro
Sheets(wssh1).Activate
Columns("E").Copy 'Column E is the Column with the criteria to split the master WB
Sheets.Add ‘Add a new sheet with the column of criteria copied
ActiveSheet.Name = "SHS Service Line" ‘Name the new WS created
Range("A1").PasteSpecial
Columns("A").RemoveDuplicates Columns:=1, Header:=xlYes ‘make the criteria unique

vcounter = Range("A" & Rows.Count).End(xlUp).Row 'get the last row of criteria

For i = 2 To vcounter
vfilter = Sheets("SHS Service Line").Cells(i, 1) 'get the first criteria to filter from SHS WS
Sheets(wssh1).Activate
ActiveSheet.Columns.AutoFilter field:=Columns("E").Column, Criteria1:=vfilter 'filter the Master ws based on the vfilter from SHS Service Line WS {Criteria1 written as Criteria'one'}
Range("A1:S" & LR).SpecialCells(xlCellTypeVisible).Select ‘Select the autofiltered cells only
Selection.Copy

‘Location of where the new SUB WB will be located
subwbpath = "C:\Users\mychel.PRIMESTAFF\PrimeStaff Management Services Pte Ltd\TAATs - General\VBA Split\SHS 1\"
subwbname = "*" & vfilter & "*_List.xlsx" ‘file name of new SUB WB

If Dir(subwbpath & subwbname) = "" Then ‘check if a SUB WB exist
‘SUB WB does not exist
Workbooks.Add ‘Create a SUB WB
ActiveSheet.Name = vfilter ‘Name the new WS
Range("A1").PasteSpecial Paste:=xlPasteAll ‘Paste the selected cell from master WS

  If vfilter <> "" Then 'save new SUB WB
  ActiveWorkbook.SaveAs ThisWorkbook.Path & "\SHS 1\" & vfilter & "_List.xlsx"
  Else
  ActiveWorkbook.SaveAs ThisWorkbook.Path & "\SHS 1\_Empty"
  End If

Else

Workbooks.Open ThisWorkbook.Path & "\SHS 1\" & vfilter & "_List.xlsx" 'SUB WB exists
Sheets(vfilter).Range("A1").PasteSpecial Paste:=xlPasteAll ‘Paste the selected cell from master WS
  
  If vfilter <> "" Then
  ActiveWorkbook.Save 'save current wb
  Else
  ActiveWorkbook.SaveAs ThisWorkbook.Path & "\SHS 1\_Empty"
  End If
  
End If

ActiveWorkbook.Close
Workbooks(wswb).Activate
Next i


'NNA sheet macro
Sheets(wssh2).Activate
Columns("E").Copy 'Column E is the Column with the criteria to split the master WB
Sheets.Add ‘Add a new sheet with the column of criteria copied
ActiveSheet.Name = "NNA Service Line" ‘Name the new WS created
Range("A1").PasteSpecial
Columns("A").RemoveDuplicates Columns:=1, Header:=xlYes ‘make the criteria unique

vcounter = Range("A" & Rows.Count).End(xlUp).Row 'get the last row of criteria

For i = 2 To vcounter
vfilter = Sheets("NNA Service Line").Cells(i, 1) 'get the first criteria to filter from SHS WS
Sheets(wssh2).Activate
ActiveSheet.Columns.AutoFilter field:=Columns("E").Column, Criteria1:=vfilter 'filter the Master ws based on the vfilter from SHS Service Line WS {Criteria1 written as Criteria'one'}
Range("A1:S" & LR).SpecialCells(xlCellTypeVisible).Select ‘Select the autofiltered cells only
Selection.Copy

‘Location of where the new SUB WB will be located
subwbpath = "C:\Users\mychel.PRIMESTAFF\PrimeStaff Management Services Pte Ltd\TAATs - General\VBA Split\NNA 1\"
subwbname = "*" & vfilter & "*_List.xlsx" ‘file name of new SUB WB

If Dir(subwbpath & subwbname) = "" Then ‘check if a SUB WB exist
‘SUB WB does not exist
Workbooks.Add ‘Create a SUB WB
ActiveSheet.Name = vfilter ‘Name the new WS
Range("A1").PasteSpecial Paste:=xlPasteAll ‘Paste the selected cell from master WS

  If vfilter <> "" Then 'save new SUB WB
  ActiveWorkbook.SaveAs ThisWorkbook.Path & "\NNA 1\" & vfilter & "_List.xlsx"
  Else
  ActiveWorkbook.SaveAs ThisWorkbook.Path & "\NNA 1\_Empty"
  End If

Else

Workbooks.Open ThisWorkbook.Path & "\NNA 1\" & vfilter & "_List.xlsx" 'SUB WB exists
Sheets(vfilter).Range("A1").PasteSpecial Paste:=xlPasteAll ‘Paste the selected cell from master WS
  
  If vfilter <> "" Then
  ActiveWorkbook.Save 'save current wb
  Else
  ActiveWorkbook.SaveAs ThisWorkbook.Path & "\NNA 1\_Empty"
  End If
  
End If

ActiveWorkbook.Close
Workbooks(wswb).Activate
Next i

Sheets(wssh2).ShowAllData
Sheets(wssh1).ShowAllData
Sheets("SHS Service Line").Delete
Sheets("NNA Service Line").Delete

End Sub


'if you want to work with pivot tables, you will have to run this macro once, to get data inside the sub file first
'open the sub file to get the range of the data in the pivot table.
'Set the pivot table range with the exact columns but set the rows ar the way to the bottom.
'example: Range(A:F)
'arrange your fields to how u would like the pivot table to look like
'right click on your pivot table to 'pivot table options' and go to the data tab to set the pivot table to refresh every time you open the excel.
