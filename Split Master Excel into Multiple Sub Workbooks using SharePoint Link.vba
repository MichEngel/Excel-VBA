‘Used for 1 Criteria
‘Used for Sharepoint link files
‘Used for 2 master worksheets

Option Compare Text 'make upper and lower case identical

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, _
ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As LongPtr

Sub Split()

Dim wswb, wssh1, wssh2 As String
Dim filepath, fullpath, tempfile As String
Dim result As LongPtr
Dim LR As Long

wswb = "Desktop Split.xlsm" 'Master workbook Name
wssh1 = "SHS" 'SHS Worksheet Name
wssh2 = "NNA" 'NNA Worksheet Name
Vcolumn = Columns("E") 'Column with the criteria to split the master wb
LR = Range("A" & Rows.Count).End(xlUp).Row 'last row to copy to new sub wb

'SHS sheet macro
Sheets(wssh1).Activate
Columns("E").Copy 'Column E is used to split the master wb
Sheets.Add ‘add a new ws to copy the copied column
ActiveSheet.Name = "SHS Service Line" 'know how many different criteria there are
Range("A1").PasteSpecial
Columns("A").RemoveDuplicates Columns:=1, Header:=xlYes
‘This tells you how many sub wb you will get

'get the last row of criteria from SHS Service Line ws
vcounter = Range("A" & Rows.Count).End(xlUp).Row

‘Loops from first to last criteria
For i = 2 To vcounter
vfilter = Sheets("SHS Service Line").Cells(i, 1) ‘criteria to filter for SHS ws
Sheets(wssh1).Activate
'filter the Master ws based on the vfilter from SHS SERVICE LINE WS 
‘Criteria1 written as Criteria'one'
ActiveSheet.Columns.AutoFilter field:=Columns("E").Column, Criteria1:=vfilter 
‘Select the auto filtered cells only
Range("A1:R" & LR).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

'set file URL and temporary file path
filepath = "https://primestaffmanagement.sharepoint.com/sites/Project-PS/Shared%20Documents/General/" '(CHANGE TO FOLDER PATH WHERE YOU WILL WANT THE SUB EXCEL FILE TO BE)
‘include the filename to create full path of the sub wb
fullpath = filepath & vfilter & "_List.xlsx" 

‘Create a path for the tempfile to be downloaded at
tempfile = Environ("TEMP") & "\" & vfilter & "_Lists.xlsx"

'attempt to download the file
result = URLDownloadToFile(0, fullpath, tempfile, 0, 0)

'check if file exists;
‘If file exists, result = 0
‘Open existing wb
‘paste selected cells into the criteria ws
‘Save
‘Else {result <> 0}
‘Create new wb
‘Paste selected cells into the criteria ws
‘Save as

If result = 0 Then

  Workbooks.Open ThisWorkbook.Path & "\" & vfilter & "_List.xlsx"
  Sheets(vfilter).Range("A1").PasteSpecial Paste:=xlPasteAll
  
  If vfilter <> "" Then
  ActiveWorkbook.Save 'save current wb
  Else
  ActiveWorkbook.SaveAs (ThisWorkbook.Path & "\_Empty.xlsx")
  End If
  
Else

  Workbooks.Add
  ActiveSheet.Name = vfilter
  Range("A1").PasteSpecial Paste:=xlPasteAll

    If vfilter <> "" Then 'save new sub wb
      ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & vfilter & "_List.xlsx"
    Else
      ActiveWorkbook.SaveAs ThisWorkbook.Path & "\_Empty.xlsx"
    End If

End If

'delete the tempfile
On Error Resume Next
Kill tempfile
On Error GoTo 0

ActiveWorkbook.Close
Workbooks(wswb).Activate
Next i


'NNA sheet macro
Sheets(wssh2).Activate
Columns("E").Copy 'Column E is used to split the master wb
Sheets.Add ‘add a new ws to copy the copied column
ActiveSheet.Name = "NNA Service Line" 'know how many different criteria there are
Range("A1").PasteSpecial
Columns("A").RemoveDuplicates Columns:=1, Header:=xlYes
‘This tells you how many sub wb you will get

'get the last row of criteria from NNA Service Line ws
vcounter = Range("A" & Rows.Count).End(xlUp).Row

‘Loops from first to last criteria
For i = 2 To vcounter
vfilter = Sheets("NNA Service Line").Cells(i, 1) ‘criteria to filter for SHS ws
Sheets(wssh2).Activate
'filter the Master ws based on the vfilter from NNA SERVICE LINE WS 
‘Criteria1 written as Criteria'one'
ActiveSheet.Columns.AutoFilter field:=Columns("E").Column, Criteria1:=vfilter 
‘Select the auto filtered cells only
Range("A1:R" & LR).SpecialCells(xlCellTypeVisible).Select
Selection.Copy

'set file URL and temporary file path
filepath = "https://primestaffmanagement.sharepoint.com/sites/Project-PS/Shared%20Documents/General/" '(CHANGE TO FOLDER PATH WHERE YOU WILL WANT THE SUB EXCEL FILE TO BE)
‘include the filename to create full path of the sub wb
fullpath = filepath & vfilter & "_List.xlsx" 

‘Create a path for the tempfile to be downloaded at
tempfile = Environ("TEMP") & "\" & vfilter & "_Lists.xlsx"

'attempt to download the file
result = URLDownloadToFile(0, fullpath, tempfile, 0, 0)

'check if file exists;
‘If file exists, result = 0
‘Open existing wb
‘paste selected cells into the criteria ws
‘Save
‘Else {result <> 0}
‘Create new wb
‘Paste selected cells into the criteria ws
‘Save as

If result = 0 Then

  Workbooks.Open ThisWorkbook.Path & "\" & vfilter & "_List.xlsx"
  Sheets(vfilter).Range("A1").PasteSpecial Paste:=xlPasteAll
  
  If vfilter <> "" Then
  ActiveWorkbook.Save 'save current wb
  Else
  ActiveWorkbook.SaveAs (ThisWorkbook.Path & "\_Empty.xlsx")
  End If
  
Else

  Workbooks.Add
  ActiveSheet.Name = vfilter
  Range("A1").PasteSpecial Paste:=xlPasteAll

    If vfilter <> "" Then 'save new sub wb
      ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & vfilter & "_List.xlsx"
    Else
      ActiveWorkbook.SaveAs ThisWorkbook.Path & "\_Empty.xlsx"
    End If

End If

'delete the tempfile
On Error Resume Next
Kill tempfile
On Error GoTo 0

ActiveWorkbook.Close
Workbooks(wswb).Activate
Next i

Sheets(wssh2).ShowAllData
Sheets(wssh1).ShowAllData
Sheets("SHS Service Line").Delete
Sheets("NNA Service Line").Delete

End Sub
