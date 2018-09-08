Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Public Function ExportObj()
    On Error GoTo HandleError

    Dim oDB         As DAO.Database
    Dim oRS         As DAO.Recordset
    Dim sSQL        As String
    Dim iMsg        As Integer
    Dim sDate       As String
    Dim sFile       As String
    Dim FSO         As Scripting.FileSystemObject
    Dim txs         As Scripting.TextStream

    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmDebugWindow

    ' -- Retrieve list of reports from tblReports -- '
    Set oDB = CurrentDb()
    sSQL = "Select rptID, rptDisplayName, rptObjectName, rptFormat, rptExportLocation, Daily, Monday, Tuesday, Wednesday, Thursday, Friday " & _
           "from tblReports order by rptID"
    Set oRS = oDB.OpenRecordset(sSQL)
    oRS.MoveFirst
    
    ' -- Any reports in set? -- '
    If oRS.BOF = True And oRS.EOF = True Then
        iErr = MsgBox("No reports found.", vbOKOnly, "Report Execution Error")
       GoTo Proc_Exit
    End If

    '-- Set date/timestamp. --'
    sDate = "_" & Format(Now, "yyyymmdd")

    Debug.Print "Report Run Started: " & Now
    
    '-- Execute Reports Start -- '
    Do While Not oRS.EOF
    If oRS.Fields("Daily") = True Then
    GoTo reportrun
    ElseIf oRS.Fields("Monday") = True And Weekday(Date) = 2 Then
    GoTo reportrun
    ElseIf oRS.Fields("Tuesday") = True And Weekday(Date) = 3 Then
    GoTo reportrun
    ElseIf oRS.Fields("Wednesday") = True And Weekday(Date) = 4 Then
    GoTo reportrun
    ElseIf oRS.Fields("Thursday") = True And Weekday(Date) = 5 Then
    GoTo reportrun
    ElseIf oRS.Fields("Friday") = True And Weekday(Date) = 6 Then
    GoTo reportrun
    Else
    oRS.MoveNext
    End If
    If oRS.EOF Then
    GoTo Proc_Exit
    End If
Loop

    '-- Execute MS Object Output -- '
reportrun:
    Debug.Print oRS.Fields("rptID") & " " & oRS.Fields("rptDisplayName") & " " & "Started:" & " " & Now
    
    DoCmd.SetWarnings False
    sFile = oRS.Fields("rptExportLocation") & "\" & oRS.Fields("rptDisplayName") & sDate & ".xlsx"
    Select Case oRS.Fields("rptFormat")
        Case "E" - -reportObject Is Excel
            DoCmd.TransferSpreadsheet acExport, , oRS.Fields("rptObjectName"), sFile
            Call ExcelFormatGeneric(sFile) '****this function will format the excel output ****
        Case "M" '-- reportObject is a Macro
            DoCmd.RunMacro oRS.Fields("rptObjectName")
        Case "Q" '-- reportObject is a Query
            DoCmd.OpenQuery oRS.Fields("rptObjectName")
        Case "R" '-- reportObject is a Report
            DoCmd.OutputTo acOutputReport, oRS.Fields("rptObjectName"), acFormatSNP, Replace(sFile, ".xlsx", ".snp")
    End Select
    Debug.Print oRS.Fields("rptID") & " '" & oRS.Fields("rptDisplayName") & "' completed: " & Now
    oRS.MoveNext
Loop

'-- Complete.
oRS.Close
oDB.Close
Set txs = Nothing
Set FSO = Nothing

DoCmd.SetWarnings = True

Proc_Exit:
    Debug.Print "Report Run FInished: " & Now
    DoCmd.SetWarnings True
    Exit Function

HandleError:

    sErr = "Error: " & Err.Number & vbclrf & "Description:  " & Err.Description & vbCrLf & "Function: exportReports()"
    iErr = MsgBox(sContact & sErr, vbOKOnly, "Report Execution Error")
    
GoTo Proc_Exit

End Function
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

