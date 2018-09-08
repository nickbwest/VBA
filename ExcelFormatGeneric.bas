Attribute VB_Name = "Module2"
Option Compare Database

Option Explicit

Public Sub ExcelFormatGeneric(sFile As String)
On Error GoTo Err_FormatExcelExports

Dim xlApp As Object
Dim xlsheet As Object

Set xlApp = CreateObject("Excel.Application")
Set xlsheet = xlApp.Workbooks.Open(sFile).Sheets(1)

With xlApp
    
    .Application.cells.Select
    .Application.Selection.Font.Name = "Lucinda Sans Unicode"
    .Application.Selection.Font.Size = 8
    .Application.Selection.HorizontalAlignment = xlLeft
    .Application.Selection.Rows.AutoFit
    .Application.Selection.Columns.AutoFit
    .Application.Range("C3").Select
    .Application.ActiveWindow.FreezePanes = True
    .Application.ActiveWorkbook.Save
    .Application.ActiveWorkbook.Close
    .Quit
     
End With
'use this if you want to set the workbook as readonly
 'SetAttr sFile, vbReadOnly

Set xlApp = Nothing

Set xlsheet = Nothing

Exit_FormatExcelExports:
    
    Exit Sub
  

Err_FormatExcelExports:


    MsgBox Err.Number & " - " & Err.Description
     

    Resume Exit_FormatExcelExports

   

 End Sub




