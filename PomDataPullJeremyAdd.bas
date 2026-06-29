Attribute VB_Name = "PomDataPullJeremyAdd"
Option Explicit
Public Sub PbeDatabasePull()
' ***INITIALIZE***
    ' Database Objects
    Dim SqlDb As New ADODB.Connection
    Dim SqlData As New ADODB.Recordset
    Dim SqlConn As String
    SqlConn = "Provider=SQLOLEDB;Data Source=pbebdsmessier.nos.boeing.com;User ID=SvcIedwhExcel ;Password=0x9F4DE51F3477553AE46A752DD6B66FCE;"
    ' Line Number String Creation Objects
    Dim Row As Integer
    Dim Col As Integer
    Dim LineNumberString As String
    ' Testing Can delete these three if you dont want the message box
    Dim StartTime As Double
    Dim msgAlert As Variant
    StartTime = Timer
' ***CREATE LINE NUMBER STRING
    Sheets("QUERY").Range("B2") = "" ' Reset Value
    With Sheets("Airplane List")
        For Row = 2 To 21 ' Loop and Recreate String
            For Col = 1 To 2 ' Get both columns
                If .Cells(Row, Col) <> "" Then
                    If LineNumberString = "" Then
                        LineNumberString = "('" & .Cells(Row, Col) & "'" ' Populate for Blank
                    Else
                        LineNumberString = LineNumberString & ", '" & .Cells(Row, Col) & "'" ' Populate not blank
                    End If
                End If
            Next
        Next
        If LineNumberString <> "" Then
            LineNumberString = LineNumberString & ")" ' Trailing Parenthesis
        Else
            LineNumberString = "('NONE')" ' Default to avoid error
        End If
    End With
    Sheets("QUERY").Range("B2") = LineNumberString ' Populate final string in Query Tab
' ***PULL DATA***
    SqlDb.Open SqlConn ' Connect to DB
    SqlData.Open Sheets("QUERY").Range("B1").Value & Sheets("QUERY").Range("B2").Value, SqlDb ' Run Query
' ***LOAD DATA
    With Sheets("FileImport")
        .Rows("2:10000").ClearContents ' Clear Current Data
        .Range("A2").CopyFromRecordset SqlData: SqlData.Close ' Load table
    End With
' ***WRAP UP***
    SqlDb.Close ' Disconnect from DB
' *** RUN TIME MESSAGE - DELETE TO GET RID OF POPUUP - TESTING ONLY
    msgAlert = MsgBox("Data Updated with Run Time of " & Round(Timer - StartTime, 1) & " second", vbOKOnly, "Data Pull Complete")
End Sub

