‘Make sure you have a button to run this macro
‘Make sure there is a cell to indicate which row to run the macro

Private Sub Mail Merge Document() 'Extension Cross Dock Contract

Dim ContractRow, ContractCol, LastRow As Long ‘lastrow may not be necessary
Dim CandName, WordTempLoc, TagName, TagValue, FileName As String
Dim Sent As Date
Dim WordDoc, WordApp As Object
Dim WordContent As Word.Range

With Sheet1 ‘use the real name of the sheet instead of the rename of the sheet

‘Make sure the reference of VBA MS WORD 16.0 OBJECT LIBRARY is available
On Error Resume Next
‘Open word application in background to prepare for mail merging
  Set WordApp = GetObject("Word.Application") 
    If Err.Number <> 0 Then
      Err.Clear
      Set WordApp = CreateObject("word.application")
      WordApp.Visible = True
     End If
    
Sent = .Range("A" & ContractRow).Value ‘cell that will refresh when the macro is ran for that row
CandName = .Range("M" & ContractRow).Value ‘candidate name that will be run by the macro
ContractRow = .Range("I1").Value ‘row that that the macro will run the mail merge
LastRow = .Range("I9999").End(xlUp).Row ‘may not be necessary

  If ContractRow = .Range("I1").Value And Sent = "" Then ‘condition for the macro to run
    Set WordDoc = WordApp.Documents.Open(FileName:=("A:\Temp & Contract\About Projects\Project - Delivery Hero\Template\LOA - EXTENSION CONTRACT.docx"), ReadOnly:=False)
‘Exact document that will be mail merged, open the document up
    For ContractCol = 3 To 21 ‘columns with values that will be merged into the word document
 ‘tags that will be in the word document will have to match these TagName in the excel’s 3rd row
      TagName = Sheet1.Cells(3, ContractCol).Value
‘values in excel that will be replaced in the word document
      TagValue = Sheet1.Cells(ContractRow, ContractCol).Value
        With WordDoc.Content.Find ‘running the mail merge
          .Text = TagName ‘match the TagName with the <<mail merge name>> in the word doc
          .Replacement.Text = TagValue ‘replace the TagName with the TagValue from the excel
          .Wrap = wdFindContinue ‘go to the next TagName
          .Execute Replace:=wdReplaceAll
        End With
    Next ContractCol
    FileName = "A:\Temp & Contract\About Projects\Project - Delivery Hero\STORES SINGAPORE\CROSS DOCK - Project\A) ALL LOA & ADDENDEMS for Cross Dock" & "\" & .Range("M" & ContractRow).Value & " - LOA (EXTENSION)" & ".docx" ‘set the path and file name of the word doc
    WordDoc.SaveAs FileName ‘saveas the word doc, so that the word doc with TagName can be reused
    .Range("A" & ContractRow) = Now  ‘macro is completed, put the timestamp into sent column
  End If
End With
End Sub
