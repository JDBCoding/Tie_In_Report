Attribute VB_Name = "Module1"
Dim nextRun As Date ' this is for Sub UpdateTimestamp

Private Sub Workbook_Open() ' This runs the update timestamp sub when the workbook is opened
    ' Call the subroutine to start updating the timestamp
    'Call StartUpdatingTimestamp
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Call the subroutine to stop updating the timestamp
    'Call StopUpdatingTimestamp
End Sub

Sub abuildTieIn()

    StopUpdatingTimestamp ' Stops timestamp update sub so there are no conflicts
    ClearTieInData ' This clears sheet FileImport, FileExport, and FileConfig to prepare for data processing
    ImportData ' This pulls in the live data to the FileImport
    ReportTimeStamp ' add in a timestamp to fill out the Report generated later
    EmptyImportCheck ' This checks the POM outputted data, if no data is found it kills the program
    ImportFileTimestamps ' this takes the timestamps off of existing New Data and Old Data and paste it in FileConfig
    DataTrimTieIn ' This sets up the fileExport sheet
    CreateArchiveFolder ' This makes an archive folder based on existing New Data time stamp
    CompareTimestamps ' this checks if the data is older than 4 hours
    With ThisWorkbook.Sheets("FileConfig") ' Replace Old data if its older than 4 hours, keep if its younger
        If .Range("B5").Value = "Replace" Then
            Call DeleteOldDataWorkbooks ' This deletes the current old Data file
            Call NewToOld ' this turns the current new Data to old Data
            Call ExportUpdatedNewData ' this exports the updated NewData files to the active folder
        ElseIf .Range("B5").Value = "Keep" Then
            Call DeleteNewDataWorkbooks ' this deletes the current new Data book to make room for the next one
            Call ExportUpdatedNewData ' this exports the updated NewData files to the active folder
        End If
    End With ' Exit If/Then statment for how to handle Old Data
    ArchiveData ' this copys the new and old data and puts it in the archive folder
    NewOldDataImport ' Import current newData and oldData files to their respective sheets
    CopyDataToCompare ' Combine new and old data on one sheet for comparison checks
    CheckForNewComments ' if the comment was updated or job added put comment in row X of Compare
    emptyCommentUpdate ' deletes the comment flag if the flag was placed from comment being deleted, this prevents issues down the line
    CheckForNewJobs ' mark any new jobs with the word New on column V of Compare
    CheckForSells ' mark any new jobs with the word sell on column W of Compare
    MoveToTotalList ' Moves data to sheet TotalList so we only have one list of data to work with
    LnCorrect ' Change all VH and VZ numbers to their Line number, groups boom and AP jobs together
    lostToolPlanUpdate ' changes the plan number to the fake lost tool number to sort them into their own section
    AssignSortNumberToJob ' Uses the Job Sort list to number each job so it can be ordered correctly
    IncreaseChildSortNumber 'Increases sort number by 1 so children show under their parent job
    AddJobType ' adds the Job type to each job, this is used later for sorting
    SortDataByAirplaneAndJob ' Groups first by airplane, then by the assigned job number
    UpdateJobTittle ' Use the MESCI tittle for the job if the user POM notes are empty
    flightOrderCount ' This creates the flight order count martix on sheet TotalList, logic added to exclude sold orders
    flightOrderCleanListBuild 'This makes a list of the flight orders to keep for the main tie in
    flightOrderCleanJobDelete ' This deletes all flight jobs that are duplicates of ones we already have
    MoveDataToDataClean ' move essential data to DataClean for final formatting
    AddColon ' adds : to the end of each job
    UpdateSoldNotes ' Overwrite any notes on sold jobs with "Sold", also addes "Sold" to jobs without notes
    TextColorUpdate ' Changes the text colors according on their job type
    InsertLinesBetweenAirplanes ' This puts 3 lines between each airplane
    CreateAirplaneHeader 'Creates the LN header between each airplane
    InsertLinesBetweenJobTypes ' inserts a single line break between each job type
    CreateJobHeader ' set up the header for each job type
    MoveComments ' move the comments to below each job that has a comment
    CreateHeaderSection 'Builds the header section at the top
    TextFormat ' reformat the text so it looks nice in the email
    addFlightOrderLeftCommentPreflight ' This add the comment how many preflights remain
    addFlightOrderLeftCommentPostflight 'This adds the comment how many post flights are left
    addFlightOrderLeftCommentLMI 'This adds the comment how many LMI orders are left
    addFlightOrderLeftCommentQuickturn 'This adds the comment how many Quickturn orders are left
    OpenEmailTieIn ' this copies the data from excel and paste it as a table in outlook
    StartUpdatingTimestamp ' resumes the update timestamp sub for POM data

End Sub

Sub aBuildTieInManualFileSelect()

    StopUpdatingTimestamp ' Stops timestamp update sub so there are no conflicts
    ClearTieInData ' This clears sheet FileImport, FileExport, and FileConfig to prepare for data processing
    SelectNewData ' This lets the user pick a file to use as the NewData
    SelectOldData ' This lets the user pick a file to use as the OldData
    ReportTimeStamp ' add in a timestamp to fill out the Report generated later
    CopyDataToCompare ' Combine new and old data on one sheet for comparison checks
    CheckForNewComments ' if the comment was updated or job added put comment in row X of Compare
    CheckForNewJobs ' mark any new jobs with the word New on column V of Compare
    emptyCommentUpdate ' deletes the comment flag if the flag was placed from comment being deleted, this prevents issues down the line
    CheckForSells ' mark any new jobs with the word sell on column W of Compare
    MoveToTotalList ' Moves data to sheet TotalList so we only have one list of data to work with
    LnCorrect ' Change all VH and VZ numbers to their Line number, groups boom and AP jobs together
    lostToolPlanUpdate ' changes the plan number to the fake lost tool number to sort them into their own section
    AssignSortNumberToJob ' Uses the Job Sort list to number each job so it can be ordered correctly
    IncreaseChildSortNumber 'Increases sort number by 1 so children show under their parent job
    AddJobType ' adds the Job type to each job, this is used later for sorting
    SortDataByAirplaneAndJob ' Groups first by airplane, then by the assigned job number
    UpdateJobTittle ' Use the MESCI tittle for the job if the user POM notes are empty
    flightOrderCount ' This creates the flight order count martix on sheet TotalList, logic added to exclude sold orders
    flightOrderCleanListBuild 'This makes a list of the flight orders to keep for the main tie in
    flightOrderCleanJobDelete ' This deletes all flight jobs that are duplicates of ones we already have
    MoveDataToDataClean ' move essential data to DataClean for final formatting
    AddColon ' adds : to the end of each job
    UpdateSoldNotes ' Overwrite any notes on sold jobs with "Sold", also addes "Sold" to jobs without notes
    TextColorUpdate ' Changes the text colors according on their job type
    InsertLinesBetweenAirplanes ' This puts 3 lines between each airplane
    CreateAirplaneHeader 'Creates the LN header between each airplane
    InsertLinesBetweenJobTypes ' inserts a single line break between each job type
    CreateJobHeader ' set up the header for each job type
    MoveComments ' move the comments to below each job that has a comment
    CreateHeaderSection 'Builds the header section at the top
    TextFormat ' reformat the text so it looks nice in the email
    addFlightOrderLeftCommentPreflight ' This add the comment how many preflights remain
    addFlightOrderLeftCommentPostflight 'This adds the comment how many post flights are left
    addFlightOrderLeftCommentLMI 'This adds the comment how many LMI orders are left
    addFlightOrderLeftCommentQuickturn 'This adds the comment how many Quickturn orders are left
    OpenEmailTieIn ' this copies the data from excel and paste it as a table in outlook
    StartUpdatingTimestamp ' resumes the update timestamp sub for POM data
    
End Sub

Sub aFirstShiftDar()

    StopUpdatingTimestamp ' Stops timestamp update sub so there are no conflicts
    ClearFirstDar ' clears all the old data out of working sheets
    ImportData ' This pulls in the live data to the FileImport
    EmptyImportCheck ' This checks the POM outputted data, if no data is found it kills the program
    DataTrimDar ' this configures the FileExport to the expected job format and column order
    FirstDarCreateArchive ' This creates an archive folder and puts the raw data in it
    UpdateJobTitleDAR ' This uses defult MESCI title if the POM notes are blank
    LnCorrectDAR 'Change all VH and VZ numbers to their Line number
    CombineWOAndJobDescription ' combines WO and Job title into one cell
    SeparateNames ' This breaks out each tech assigned to a job in their own cell
    FillFirstDar ' This uses Key values to move jobs to each person in the dar
    FillFirstDarPriorityBoard ' this puts one instance of each job assigned on the priority board
    UpdateFirstDarTimeStamp ' shows the updated tab with when the POM data was created
    StartUpdatingTimestamp ' resumes the update timestamp sub for POM data
    ThisWorkbook.Sheets("FirstDar").Activate

End Sub

Sub aSecondShiftDar()

    StopUpdatingTimestamp ' Stops timestamp update sub so there are no conflicts
    ClearSecondDar ' clears all the old data out of working sheets
    NicknameRemove ' Removes nicknames so the data process works properly
    ImportData ' This pulls in the live data to the FileImport
    EmptyImportCheck ' This checks the POM outputted data, if no data is found it kills the program
    DataTrimDar ' this configures the FileExport to the expected job format and column order
    SecondDarCreateArchive ' This creates an archive folder and puts the raw data in it
    UpdateJobTitleDAR ' This uses defult MESCI title if the POM notes are blank
    LnCorrectDAR 'Change all VH and VZ numbers to their Line number
    CombineWOAndJobDescription ' combines WO and Job title into one cell
    SeparateNames ' This breaks out each tech assigned to a job in their own cell
    FillSecondDar ' This uses Key values to move jobs to each person in the dar
    FillSecondDarPriorityBoard ' this puts one instance of each job assigned on the priority board
    UpdateSecondDarTimeStamp ' shows the updated tab with when the POM data was created
    NicknameAdd ' Adds perferred nicknames back into the final DAR
    StartUpdatingTimestamp ' resumes the update timestamp sub for POM data
    ThisWorkbook.Sheets("SecondDar").Activate

End Sub

Sub aThirdShiftDar()

    StopUpdatingTimestamp ' Stops timestamp update sub so there are no conflicts
    ClearThirdDar ' clears all the old data out of working sheets
    ImportData ' This pulls in the live data to the FileImport
    EmptyImportCheck ' This checks the POM outputted data, if no data is found it kills the program
    DataTrimDar ' this configures the FileExport to the expected job format and column order
    ThirdDarCreateArchive ' This creates an archive folder and puts the raw data in it
    UpdateJobTitleDAR ' This uses defult MESCI title if the POM notes are blank
    LnCorrectDAR 'Change all VH and VZ numbers to their Line number
    CombineWOAndJobDescription ' combines WO and Job title into one cell
    SeparateNames ' This breaks out each tech assigned to a job in their own cell
    FillThirdDar ' This uses Key values to move jobs to each person in the dar
    FillThirdDarPriorityBoard ' this puts one instance of each job assigned on the priority board
    UpdateThirdDarTimeStamp ' shows the updated tab with when the POM data was created
    StartUpdatingTimestamp ' resumes the update timestamp sub for POM data
    ThisWorkbook.Sheets("ThirdDar").Activate

End Sub

Sub aStaticList()

    StopUpdatingTimestamp ' Stops timestamp update sub so there are no conflicts
    ClearTieInData ' This clears sheet FileImport, FileExport, and FileConfig to prepare for data processing
    ImportData ' This pulls in the live data to the FileImport
    EmptyImportCheck ' This checks the POM outputted data, if no data is found it kills the program
    DataTrimTieIn ' This sets up the fileExport sheet
    joblistDataTransfer ' This moves data straight from file export to TotalList
    LNcorrectJoblist ' This updates the airplane line numbers
    lostToolPlanUpdate ' changes the plan number to the fake lost tool number to sort them into their own section
    AssignSortNumberToJob ' Uses the Job Sort list to number each job so it can be ordered correctly
    IncreaseChildSortNumber 'Increases sort number by 1 so children show under their parent job
    AddJobType ' adds the Job type to each job, this is used later for sorting
    SortDataByAirplaneAndJob ' Groups first by airplane, then by the assigned job number
    UpdateJobTittle ' Use the MESCI tittle for the job if the user POM notes are empty
    MoveDataToDataClean ' move essential data to DataClean for final formatting
    AddColon ' adds : to the end of each job
    UpdateSoldNotes ' Overwrite any notes on sold jobs with "Sold", also addes "Sold" to jobs without notes
    TextColorUpdate ' Changes the text colors according on their job type
    InsertLinesBetweenAirplanes ' This puts 3 lines between each airplane
    CreateAirplaneHeader 'Creates the LN header between each airplane
    InsertLinesBetweenJobTypes ' inserts a single line break between each job type
    CreateJobHeader ' set up the header for each job type
    MoveComments ' move the comments to below each job that has a comment
    CreateHeaderSection 'Builds the header section at the top
    TextFormat ' reformat the text so it looks nice in the email
    openWordTieIn ' pastes the tie in into a word doc so it can be viewed
    StartUpdatingTimestamp ' resumes the update timestamp sub for POM data

End Sub

Sub updateRefreshAP()

' this code is no longer used

End Sub

Sub bFirstPrioritySort()
    Dim wsCompare As Worksheet
    Dim wsFirstDar As Worksheet
    Dim lastRow As Long

    ' Set references to the sheets
    Set wsCompare = ThisWorkbook.Sheets("Compare")
    Set wsFirstDar = ThisWorkbook.Sheets("FirstDar")

    ' Clear all data on sheet "Compare"
    wsCompare.Cells.Clear

    ' Clear all data in column AG on sheet FirstDar
    wsFirstDar.Columns("AG").Clear

    ' Write "sort" in cell T1 of the Compare sheet
    wsCompare.Range("T1").Value = "sort"

    ' Copy range C24:S38 from sheet FirstDar and paste starting at A2 in Compare
    wsFirstDar.Range("C24:S38").Copy wsCompare.Range("A2")

    ' Copy range V24:V38 from sheet FirstDar and paste starting at T2 in Compare
    wsFirstDar.Range("V24:V38").Copy wsCompare.Range("T2")

    ' Unmerge all cells on sheet Compare
    wsCompare.Cells.UnMerge

    ' Sort the data on sheet Compare from smallest to largest based on column T
    lastRow = wsCompare.Cells(wsCompare.Rows.count, "A").End(xlUp).Row
    wsCompare.Range("A2:T" & lastRow).Sort Key1:=wsCompare.Range("T2"), Order1:=xlAscending, Header:=xlNo

    ' Copy data from sheet Compare and paste it in sheet FirstDar starting on row 24
    wsFirstDar.Range("C24:C" & (24 + lastRow - 2)).Value = wsCompare.Range("A2:A" & lastRow).Value
    wsFirstDar.Range("D24:D" & (24 + lastRow - 2)).Value = wsCompare.Range("B2:B" & lastRow).Value
    wsFirstDar.Range("J24:J" & (24 + lastRow - 2)).Value = wsCompare.Range("H2:H" & lastRow).Value
    wsFirstDar.Range("R24:R" & (24 + lastRow - 2)).Value = wsCompare.Range("P2:P" & lastRow).Value
    wsFirstDar.Range("S24:S" & (24 + lastRow - 2)).Value = wsCompare.Range("Q2:Q" & lastRow).Value
    wsFirstDar.Range("V24:V" & (24 + lastRow - 2)).Value = wsCompare.Range("T2:T" & lastRow).Value

    ' Copy values from J24:J38 and paste them in column AG
    wsFirstDar.Range("AG24:AG38").Value = wsFirstDar.Range("J24:J38").Value

    ' Set column AG to wrap text
    wsFirstDar.Columns("AG").WrapText = True
End Sub

Sub bSecondPrioritySort()
    Dim wsCompare As Worksheet
    Dim wsSecondDar As Worksheet
    Dim lastRow As Long

    ' Set references to the sheets
    Set wsCompare = ThisWorkbook.Sheets("Compare")
    Set wsSecondDar = ThisWorkbook.Sheets("SecondDar")

    ' Clear all data on sheet "Compare"
    wsCompare.Cells.Clear

    ' Clear all data in column AG on sheet SecondDar
    wsSecondDar.Columns("AG").Clear

    ' Write "sort" in cell T1 of the Compare sheet
    wsCompare.Range("T1").Value = "sort"

    ' Copy range C26:S40 from sheet SecondDar and paste starting at A2 in Compare
    wsSecondDar.Range("C26:S40").Copy wsCompare.Range("A2")

    ' Copy range U26:U40 from sheet SecondDar and paste starting at T2 in Compare
    wsSecondDar.Range("U26:U40").Copy wsCompare.Range("T2")

    ' Unmerge all cells on sheet Compare
    wsCompare.Cells.UnMerge

    ' Sort the data on sheet Compare from smallest to largest based on column T
    lastRow = wsCompare.Cells(wsCompare.Rows.count, "A").End(xlUp).Row
    wsCompare.Range("A2:T" & lastRow).Sort Key1:=wsCompare.Range("T2"), Order1:=xlAscending, Header:=xlNo

    ' Copy data from sheet Compare and paste it in sheet SecondDar starting on row 26
    wsSecondDar.Range("C26:C" & (26 + lastRow - 2)).Value = wsCompare.Range("A2:A" & lastRow).Value
    wsSecondDar.Range("D26:D" & (26 + lastRow - 2)).Value = wsCompare.Range("B2:B" & lastRow).Value
    wsSecondDar.Range("J26:J" & (26 + lastRow - 2)).Value = wsCompare.Range("H2:H" & lastRow).Value
    wsSecondDar.Range("R26:R" & (26 + lastRow - 2)).Value = wsCompare.Range("P2:P" & lastRow).Value
    wsSecondDar.Range("S26:S" & (26 + lastRow - 2)).Value = wsCompare.Range("Q2:Q" & lastRow).Value
    wsSecondDar.Range("U26:U" & (26 + lastRow - 2)).Value = wsCompare.Range("T2:T" & lastRow).Value

    ' Copy values from J26:J40 and paste them in column AG
    wsSecondDar.Range("AG26:AG40").Value = wsSecondDar.Range("J26:J40").Value

    ' Set column AG to wrap text
    wsSecondDar.Columns("AG").WrapText = True
End Sub

Sub bThirdPrioritySort()
    Dim wsCompare As Worksheet
    Dim wsThirdDar As Worksheet
    Dim lastRow As Long

    ' Set references to the sheets
    Set wsCompare = ThisWorkbook.Sheets("Compare")
    Set wsThirdDar = ThisWorkbook.Sheets("ThirdDar")

    ' Clear all data on sheet "Compare"
    wsCompare.Cells.Clear

    ' Clear all data in column AG on sheet ThirdDar
    wsThirdDar.Columns("AG").Clear

    ' Write "sort" in cell T1 of the Compare sheet
    wsCompare.Range("T1").Value = "sort"

    ' Copy range C17:S31 from sheet ThirdDar and paste starting at A2 in Compare
    wsThirdDar.Range("C17:S31").Copy wsCompare.Range("A2")

    ' Copy range U17:U31 from sheet ThirdDar and paste starting at T2 in Compare
    wsThirdDar.Range("U17:U31").Copy wsCompare.Range("T2")

    ' Unmerge all cells on sheet Compare
    wsCompare.Cells.UnMerge

    ' Sort the data on sheet Compare from smallest to largest based on column T
    lastRow = wsCompare.Cells(wsCompare.Rows.count, "A").End(xlUp).Row
    wsCompare.Range("A2:T" & lastRow).Sort Key1:=wsCompare.Range("T2"), Order1:=xlAscending, Header:=xlNo

    ' Copy data from sheet Compare and paste it in sheet ThirdDar starting on row 17
    wsThirdDar.Range("C17:C" & (17 + lastRow - 2)).Value = wsCompare.Range("A2:A" & lastRow).Value
    wsThirdDar.Range("D17:D" & (17 + lastRow - 2)).Value = wsCompare.Range("B2:B" & lastRow).Value
    wsThirdDar.Range("J17:J" & (17 + lastRow - 2)).Value = wsCompare.Range("H2:H" & lastRow).Value
    wsThirdDar.Range("R17:R" & (17 + lastRow - 2)).Value = wsCompare.Range("P2:P" & lastRow).Value
    wsThirdDar.Range("S17:S" & (17 + lastRow - 2)).Value = wsCompare.Range("Q2:Q" & lastRow).Value
    wsThirdDar.Range("U17:U" & (17 + lastRow - 2)).Value = wsCompare.Range("T2:T" & lastRow).Value

    ' Copy values from J17:J31 and paste them in column AG
    wsThirdDar.Range("AG17:AG31").Value = wsThirdDar.Range("J17:J31").Value

    ' Set column AG to wrap text
    wsThirdDar.Columns("AG").WrapText = True
End Sub

Sub bEmailFirstDar()
    Dim wsSource As Worksheet
    Dim wsConfig As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim currentDate As String
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim rngToCopy As Range
    
    ' Set the worksheets
    Set wsSource = ThisWorkbook.Sheets("FirstDar")
    Set wsConfig = ThisWorkbook.Sheets("FileConfig")
    
    ' Get the folder path from cell B4 in FileConfig
    folderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\DAR_Archive\First_Shift\" & wsConfig.Range("B4").Value & "\"
    
    ' Create the file name with the current date and time
    currentDate = Format(Now, "MMDDYY_HHMM")
    fileName = "First_Shift_DAR_" & currentDate & ".xlsx"
    
    ' Save a copy of the FirstDar sheet
    wsSource.Copy
    With ActiveWorkbook
        .SaveAs folderPath & fileName
        .Close False
    End With
    
    ' Set the range to copy
    Set rngToCopy = wsSource.Range("B2:AB39")
    
    ' Create a new Outlook application and mail item
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    ' Prepare the email
    With outlookMail
        .To = "DL-BDSKC46FieldShop1stshift@exchange.boeing.com" ' Set the recipient's email address
        .CC = "DL-BDSKC46FieldManagers@exchange.boeing.com; DL-BDSKC46FieldTeamLeads@exchange.boeing.com" ' Add CC recipient
        .Subject = "1st Shift DAR"
        .Body = "Please find the attached First Shift DAR."
        
        ' Copy the range to the clipboard
        rngToCopy.Copy
        
        ' Paste the range into the email body
        .Display ' Show the email before sending
        .GetInspector.WordEditor.Application.Selection.Paste
        
        ' Optionally, you can attach the saved file
        .Attachments.Add folderPath & fileName
        
    End With
    
    ' Clean up
    Set outlookMail = Nothing
    Set outlookApp = Nothing
End Sub

Sub bEmailSecondDar()
    Dim wsSource As Worksheet
    Dim wsConfig As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim currentDate As String
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim rngToCopy As Range
    
    ' Set the worksheets
    Set wsSource = ThisWorkbook.Sheets("SecondDar")
    Set wsConfig = ThisWorkbook.Sheets("FileConfig")
    
    ' Get the folder path from cell B4 in FileConfig
    folderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\DAR_Archive\Second_Shift\" & wsConfig.Range("B4").Value & "\"
    
    ' Create the file name with the current date and time
    currentDate = Format(Now, "MMDDYY_HHMM")
    fileName = "Second_Shift_DAR_" & currentDate & ".xlsx"
    
    ' Save a copy of the SecondDar sheet
    wsSource.Copy
    With ActiveWorkbook
        .SaveAs folderPath & fileName
        .Close False
    End With
    
    ' Set the range to copy
    Set rngToCopy = wsSource.Range("B2:T41")
    
    ' Create a new Outlook application and mail item
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    ' Prepare the email
    With outlookMail
        .To = "DL-KC46BDSQAInspection2nd3rdShiftsEDC@exchange.boeing.com; DL-BDSKC46FieldShop2ndshift@exchange.boeing.com" ' Set the recipient's email address
        .CC = "DL-BDSKC46FieldManagers@exchange.boeing.com; DL-BDSKC46FieldTeamLeads@exchange.boeing.com" ' Add CC recipient
        .Subject = "2nd Shift DAR"
        .Body = ""
        
        ' Copy the range to the clipboard
        rngToCopy.Copy
        
        ' Paste the range into the email body
        .Display ' Show the email before sending
        .GetInspector.WordEditor.Application.Selection.Paste
        
        ' Optionally, you can attach the saved file
        .Attachments.Add folderPath & fileName
        
    End With
    
    ' Clean up
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
End Sub

Sub bEmailThirdDar()
    Dim wsSource As Worksheet
    Dim wsConfig As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim currentDate As String
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim rngToCopy As Range
    
    ' Set the worksheets
    Set wsSource = ThisWorkbook.Sheets("ThirdDar") ' Change to ThirdDar
    Set wsConfig = ThisWorkbook.Sheets("FileConfig")
    
    ' Get the folder path from cell B4 in FileConfig
    folderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\DAR_Archive\Third_Shift\" & wsConfig.Range("B4").Value & "\"
    
    ' Create the file name with the current date and time
    currentDate = Format(Now, "MMDDYY_HHMM")
    fileName = "Third_Shift_DAR_" & currentDate & ".xlsx"
    
    ' Save a copy of the ThirdDar sheet
    wsSource.Copy
    With ActiveWorkbook
        .SaveAs folderPath & fileName
        .Close False
    End With
    
    ' Set the range to copy
    Set rngToCopy = wsSource.Range("B2:T41") ' Adjust the range if needed
    
    ' Create a new Outlook application and mail item
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    ' Prepare the email
    With outlookMail
        .To = "DL-KC46BDSQAInspection2nd3rdShiftsEDC@exchange.boeing.com; DL-BDSKC46FieldShop3rdshift@exchange.boeing.com" ' Set the recipient's email address
        .CC = "DL-BDSKC46FieldManagers@exchange.boeing.com; DL-BDSKC46FieldTeamLeads@exchange.boeing.com" ' Add CC recipient
        .Subject = "3rd Shift DAR" ' Update subject line
        .Body = ""
        
        ' Copy the range to the clipboard
        rngToCopy.Copy
        
        ' Paste the range into the email body
        .Display ' Show the email before sending
        .GetInspector.WordEditor.Application.Selection.Paste
        
        ' Optionally, you can attach the saved file
        .Attachments.Add folderPath & fileName
        
    End With
    
    ' Clean up
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
End Sub

Sub OpenEdgeWithURL()

' This code is no longer used
    
End Sub

Sub UpdateTimestamp()

    'This code is no longer used.

End Sub

Sub StartUpdatingTimestamp()
    ' Start the process of updating the timestamp
    'UpdateTimestamp
End Sub

Sub StopUpdatingTimestamp()
    ' Stop the scheduled update
   ' On Error Resume Next
   ' Application.OnTime nextRun, "UpdateTimestamp", , False
   ' On Error GoTo 0
End Sub

Sub ClearTieInData()
    Dim fileImportSheet As Worksheet
    Dim fileConfigSheet As Worksheet
    Dim fileExportSheet As Worksheet
    Dim newDataSheet As Worksheet
    Dim oldDataSheet As Worksheet
    Dim compareSheet As Worksheet
    Dim totalListSheet As Worksheet
    Dim dataCleanSheet As Worksheet

    Set fileImportSheet = ThisWorkbook.Sheets("FileImport")
    Set fileConfigSheet = ThisWorkbook.Sheets("FileConfig")
    Set fileExportSheet = ThisWorkbook.Sheets("FileExport")
    Set newDataSheet = ThisWorkbook.Sheets("NewData")
    Set oldDataSheet = ThisWorkbook.Sheets("OldData")
    Set compareSheet = ThisWorkbook.Sheets("Compare")
    Set totalListSheet = ThisWorkbook.Sheets("TotalList")
    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Clear contents and formatting for each sheet
    fileImportSheet.Cells.Clear
    fileConfigSheet.Columns("B").ClearContents
    fileExportSheet.Cells.Clear
    newDataSheet.Cells.Clear
    oldDataSheet.Cells.Clear
    compareSheet.Cells.Clear
    totalListSheet.Cells.Clear
    
    ' Clear all contents and formatting in DataClean
    With dataCleanSheet
        .Cells.Clear ' Clear all contents
        .Cells.ClearFormats ' Clear all formatting
    End With
End Sub

Sub ClearFirstDar()
    Dim fileImportSheet As Worksheet
    Dim fileExportSheet As Worksheet
    Dim firstDarSheet As Worksheet
    Dim compareSheet As Worksheet
    Dim totalListSheet As Worksheet
    Dim fileConfigSheet As Worksheet

    ' Set the worksheets
    Set fileImportSheet = ThisWorkbook.Sheets("FileImport")
    Set fileExportSheet = ThisWorkbook.Sheets("FileExport")
    Set firstDarSheet = ThisWorkbook.Sheets("FirstDar")
    Set compareSheet = ThisWorkbook.Sheets("Compare")
    Set totalListSheet = ThisWorkbook.Sheets("TotalList")
    Set fileConfigSheet = ThisWorkbook.Sheets("FileConfig")

    ' Clear the contents of the specified sheets
    fileImportSheet.Cells.Clear
    fileExportSheet.Cells.Clear
    compareSheet.Cells.Clear
    totalListSheet.Cells.Clear

    ' Clear all data in column B of the FileConfig sheet
    fileConfigSheet.Columns("B").ClearContents

    ' Clear specified ranges in the FirstDar sheet
    With firstDarSheet
        .Range("F5:U6").ClearContents
        .Range("T5:AA6").ClearContents
        .Range("F8:U21").ClearContents
        .Range("T8:AA21").ClearContents
        .Range("C24:S38").ClearContents
        .Columns("AG").ClearContents
        .Range("V24:W38").ClearContents
        .Rows("24:38").AutoFit
    End With
End Sub

Sub ClearSecondDar()
    Dim fileImportSheet As Worksheet
    Dim fileExportSheet As Worksheet
    Dim secondDarSheet As Worksheet
    Dim compareSheet As Worksheet
    Dim totalListSheet As Worksheet
    Dim fileConfigSheet As Worksheet

    ' Set the worksheets
    Set fileImportSheet = ThisWorkbook.Sheets("FileImport")
    Set fileExportSheet = ThisWorkbook.Sheets("FileExport")
    Set secondDarSheet = ThisWorkbook.Sheets("SecondDar")
    Set compareSheet = ThisWorkbook.Sheets("Compare")
    Set totalListSheet = ThisWorkbook.Sheets("TotalList")
    Set fileConfigSheet = ThisWorkbook.Sheets("FileConfig")

    ' Clear the contents of the specified sheets
    fileImportSheet.Cells.Clear
    fileExportSheet.Cells.Clear
    compareSheet.Cells.Clear
    totalListSheet.Cells.Clear

    ' Clear all data in column B of the FileConfig sheet
    fileConfigSheet.Columns("B").ClearContents

    ' Clear specified ranges in the SecondDar sheet
    With secondDarSheet
        .Range("F5:S6").ClearContents
        .Range("F8:S9").ClearContents
        .Range("F11:S15").ClearContents
        .Range("F17:S22").ClearContents
        .Range("C26:S40").ClearContents
        .Columns("AG").ClearContents
        .Range("U26:V40").ClearContents
        .Rows("26:40").AutoFit
    End With
End Sub

Sub ClearThirdDar()
    Dim fileImportSheet As Worksheet
    Dim fileExportSheet As Worksheet
    Dim thirdDarSheet As Worksheet
    Dim compareSheet As Worksheet
    Dim totalListSheet As Worksheet
    Dim fileConfigSheet As Worksheet

    ' Set the worksheets
    Set fileImportSheet = ThisWorkbook.Sheets("FileImport")
    Set fileExportSheet = ThisWorkbook.Sheets("FileExport")
    Set thirdDarSheet = ThisWorkbook.Sheets("ThirdDar")
    Set compareSheet = ThisWorkbook.Sheets("Compare")
    Set totalListSheet = ThisWorkbook.Sheets("TotalList")
    Set fileConfigSheet = ThisWorkbook.Sheets("FileConfig")

    ' Clear the contents of the specified sheets
    fileImportSheet.Cells.Clear
    fileExportSheet.Cells.Clear
    compareSheet.Cells.Clear
    totalListSheet.Cells.Clear

    ' Clear all data in column B of the FileConfig sheet
    fileConfigSheet.Columns("B").ClearContents

    ' Clear specified ranges in the SecondDar sheet
    With thirdDarSheet
        .Range("F5:S5").ClearContents
        .Range("F7:S13").ClearContents
        .Range("C17:S31").ClearContents
        .Columns("AG").ClearContents
        .Range("U17:V31").ClearContents
        .Rows("17:31").AutoFit
    End With
End Sub

Sub EmptyImportCheck()
    Dim fileImportSheet As Worksheet

    ' Set the FileImport sheet
    Set fileImportSheet = ThisWorkbook.Sheets("FileImport")

    ' Check if there is a value in cell A2 of FileImport
    If fileImportSheet.Cells(2, "A").Value = "" Then
        ' Display a message box if A2 is empty
        MsgBox "Problem with NewData, Program will close. Please re-open and try again after the next POM Data refresh.", vbExclamation
        
        ' Stop all scripts and close the workbook
        Application.DisplayAlerts = False ' Disable alerts to prevent prompts
        ThisWorkbook.Close SaveChanges:=False ' Close the workbook without saving
        Application.DisplayAlerts = True ' Re-enable alerts
    End If
End Sub

Sub ImportFileTimestamps()
    Dim configSheet As Worksheet
    Dim newFilePath As String
    Dim oldFilePath As String
    Dim newFileName As String
    Dim oldFileName As String
    Dim newFileDate As Date
    Dim oldFileDate As Date

    ' Set the config sheet
    Set configSheet = ThisWorkbook.Sheets("FileConfig")

    ' Define the file path
    Dim folderPath As String
    folderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\"

    ' Define the file names
    newFileName = "new_data_*.xlsx" ' Use wildcard to find the latest new data file
    oldFileName = "old_data_*.xlsx" ' Use wildcard to find the latest old data file

    ' Get the new data file timestamp
    newFilePath = Dir(folderPath & newFileName)
    If newFilePath <> "" Then
        newFileDate = FileDateTime(folderPath & newFilePath)
        configSheet.Range("B1").Value = newFileDate ' Store the new data timestamp
        configSheet.Range("B1").NumberFormat = "mm/dd/yyyy hh:mm" ' Format as date and time
    Else
        MsgBox "No new data file found.", vbExclamation
    End If

    ' Get the old data file timestamp
    oldFilePath = Dir(folderPath & oldFileName)
    If oldFilePath <> "" Then
        oldFileDate = FileDateTime(folderPath & oldFilePath)
        configSheet.Range("B2").Value = oldFileDate ' Store the old data timestamp
        configSheet.Range("B2").NumberFormat = "mm/dd/yyyy hh:mm" ' Format as date and time
    Else
        MsgBox "No old data file found.", vbExclamation
    End If
End Sub

Sub CreateArchiveFolder()
    Dim configSheet As Worksheet
    Dim baseFolderPath As String
    Dim yearFolderPath As String
    Dim monthFolderPath As String
    Dim archiveFolderPath As String
    Dim folderName As String
    Dim currentTime As Date
    Dim yearFolder As String
    Dim monthFolder As String

    ' Set the config sheet
    Set configSheet = ThisWorkbook.Sheets("FileConfig")

    ' Get the current date and time
    currentTime = Now

    ' Define the base folder path
    baseFolderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\Archive\"

    ' Get the year and month for folder creation
    yearFolder = Format(currentTime, "YYYY")
    monthFolder = Format(currentTime, "mmmm") ' Get the full month name

    ' Define the full paths for year and month folders
    yearFolderPath = baseFolderPath & yearFolder & "\"
    monthFolderPath = yearFolderPath & monthFolder & "\"

    ' Create the year folder if it does not exist
    On Error Resume Next
    MkDir yearFolderPath
    On Error GoTo 0

    ' Create the month folder if it does not exist
    On Error Resume Next
    MkDir monthFolderPath
    On Error GoTo 0

    ' Format the archive folder name
    folderName = "Archive_" & Format(currentTime, "DDMMYY_HHMM")

    ' Define the full path for the archive folder
    archiveFolderPath = monthFolderPath & folderName

    ' Create the archive folder
    On Error Resume Next ' Ignore errors if the folder already exists
    MkDir archiveFolderPath
    On Error GoTo 0 ' Resume normal error handling

    ' Check if the archive folder was created successfully
    If Dir(archiveFolderPath, vbDirectory) <> "" Then
        ' Place the folder name in cell B4
        configSheet.Range("B4").Value = folderName
    Else
    End If
End Sub

Sub FirstDarCreateArchive()
    Dim folderPath As String
    Dim folderName As String
    Dim subFolderName As String
    Dim timeStamp As String
    Dim newWorkbook As Workbook
    Dim fileExportSheet As Worksheet
    Dim fileConfigSheet As Worksheet
    Dim fileConfigTimestamp As String
    Dim fullFolderPath As String
    
    ' Set the base folder path
    folderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\DAR_Archive\First_Shift\"
    
    ' Create the folder name with current date and time
    timeStamp = Format(Now, "MMDDYY_HHMM")
    folderName = "Dar_Archive_" & timeStamp
    
    ' Create the full path for the main folder
    fullFolderPath = folderPath & folderName
    
    ' Create the new folder
    On Error Resume Next ' Ignore error if folder already exists
    MkDir fullFolderPath
    On Error GoTo 0 ' Resume normal error handling
    
    ' Set the worksheets
    Set fileExportSheet = ThisWorkbook.Sheets("FileExport")
    Set fileConfigSheet = ThisWorkbook.Sheets("FileConfig")
    
    ' Write the folder name to FileConfig cell B4
    fileConfigSheet.Range("B4").Value = folderName
    
    ' Get the timestamp from FileConfig cell B3
    fileConfigTimestamp = Format(fileConfigSheet.Range("B3").Value, "MMDDYY_HHMM")
    
    ' Create a new workbook
    Set newWorkbook = Workbooks.Add
    
    ' Copy data from FileExport sheet to the new workbook
    fileExportSheet.UsedRange.Copy Destination:=newWorkbook.Sheets(1).Range("A1")
    
    ' Save the new workbook in the specified main folder
    newWorkbook.SaveAs fullFolderPath & "\First_Raw_Data_" & fileConfigTimestamp & ".xlsx"
    
    ' Close the new workbook
    newWorkbook.Close SaveChanges:=False

End Sub

Sub SecondDarCreateArchive()
    Dim folderPath As String
    Dim folderName As String
    Dim subFolderName As String
    Dim timeStamp As String
    Dim newWorkbook As Workbook
    Dim fileExportSheet As Worksheet
    Dim fileConfigSheet As Worksheet
    Dim fileConfigTimestamp As String
    Dim fullFolderPath As String
    
    ' Set the base folder path
    folderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\DAR_Archive\Second_Shift\"
    
    ' Create the folder name with current date and time
    timeStamp = Format(Now, "MMDDYY_HHMM")
    folderName = "Dar_Archive_" & timeStamp
    
    ' Create the full path for the main folder
    fullFolderPath = folderPath & folderName
    
    ' Create the new folder
    On Error Resume Next ' Ignore error if folder already exists
    MkDir fullFolderPath
    On Error GoTo 0 ' Resume normal error handling
    
    ' Set the worksheets
    Set fileExportSheet = ThisWorkbook.Sheets("FileExport")
    Set fileConfigSheet = ThisWorkbook.Sheets("FileConfig")
    
    ' Write the folder name to FileConfig cell B4
    fileConfigSheet.Range("B4").Value = folderName
    
    ' Get the timestamp from FileConfig cell B3
    fileConfigTimestamp = Format(fileConfigSheet.Range("B3").Value, "MMDDYY_HHMM")
    
    ' Create a new workbook
    Set newWorkbook = Workbooks.Add
    
    ' Copy data from FileExport sheet to the new workbook
    fileExportSheet.UsedRange.Copy Destination:=newWorkbook.Sheets(1).Range("A1")
    
    ' Save the new workbook in the specified main folder
    newWorkbook.SaveAs fullFolderPath & "\Second_Raw_Data_" & fileConfigTimestamp & ".xlsx"
    
    ' Close the new workbook
    newWorkbook.Close SaveChanges:=False

End Sub

Sub ThirdDarCreateArchive()
    Dim folderPath As String
    Dim folderName As String
    Dim subFolderName As String
    Dim timeStamp As String
    Dim newWorkbook As Workbook
    Dim fileExportSheet As Worksheet
    Dim fileConfigSheet As Worksheet
    Dim fileConfigTimestamp As String
    Dim fullFolderPath As String
    
    ' Set the base folder path
    folderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\DAR_Archive\Third_Shift\"
    
    ' Create the folder name with current date and time
    timeStamp = Format(Now, "MMDDYY_HHMM")
    folderName = "Dar_Archive_" & timeStamp
    
    ' Create the full path for the main folder
    fullFolderPath = folderPath & folderName
    
    ' Create the new folder
    On Error Resume Next ' Ignore error if folder already exists
    MkDir fullFolderPath
    On Error GoTo 0 ' Resume normal error handling
    
    ' Set the worksheets
    Set fileExportSheet = ThisWorkbook.Sheets("FileExport")
    Set fileConfigSheet = ThisWorkbook.Sheets("FileConfig")
    
    ' Write the folder name to FileConfig cell B4
    fileConfigSheet.Range("B4").Value = folderName
    
    ' Get the timestamp from FileConfig cell B3
    fileConfigTimestamp = Format(fileConfigSheet.Range("B3").Value, "MMDDYY_HHMM")
    
    ' Create a new workbook
    Set newWorkbook = Workbooks.Add
    
    ' Copy data from FileExport sheet to the new workbook
    fileExportSheet.UsedRange.Copy Destination:=newWorkbook.Sheets(1).Range("A1")
    
    ' Save the new workbook in the specified main folder
    newWorkbook.SaveAs fullFolderPath & "\Third_Raw_Data_" & fileConfigTimestamp & ".xlsx"
    
    ' Close the new workbook
    newWorkbook.Close SaveChanges:=False

End Sub

Sub DeleteNewDataWorkbooks()
    Dim configSheet As Worksheet
    Dim sourceFolderPath As String
    Dim newDataFile As String
    Dim newDataFilePath As String

    ' Set the config sheet
    Set configSheet = ThisWorkbook.Sheets("FileConfig")

    ' Define the source folder path
    sourceFolderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\"

    ' Specify the new_data file pattern
    newDataFile = "new_data_*.xlsx" ' Use wildcard to find the latest new data file
    newDataFilePath = Dir(sourceFolderPath & newDataFile) ' Get the first matching file

    If newDataFilePath <> "" Then
        ' Delete the new data file
        Kill sourceFolderPath & newDataFilePath

    Else
        MsgBox "No new data file found in the source directory.", vbExclamation
    End If
End Sub

Sub DeleteOldDataWorkbooks()
    Dim configSheet As Worksheet
    Dim sourceFolderPath As String
    Dim oldDataFile As String
    Dim oldDataFilePath As String

    ' Set the config sheet
    Set configSheet = ThisWorkbook.Sheets("FileConfig")

    ' Define the source folder path
    sourceFolderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\"

    ' Specify the old_data file pattern
    oldDataFile = "old_data_*.xlsx" ' Use wildcard to find the old data files
    oldDataFilePath = Dir(sourceFolderPath & oldDataFile) ' Get the first matching file

    If oldDataFilePath <> "" Then
        ' Delete the old data file
        Do While oldDataFilePath <> ""
            Kill sourceFolderPath & oldDataFilePath
            oldDataFilePath = Dir ' Get the next matching file
        Loop
    Else
        MsgBox "No old data files found in the source directory.", vbExclamation
    End If
End Sub

Sub ArchiveData()
    Dim fso As Object
    Dim configSheet As Worksheet
    Dim sourceFolderPath As String
    Dim archivePath As String
    Dim newDataFile As String
    Dim oldDataFile As String
    Dim newDataFilePath As String
    Dim oldDataFilePath As String
    Dim yearFolder As String
    Dim monthFolder As String
    Dim currentTime As Date

    ' Set the config sheet
    Set configSheet = ThisWorkbook.Sheets("FileConfig")

    ' Define the source folder path
    sourceFolderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\"

    ' Get the current date and time
    currentTime = Now

    ' Get the year and month for folder creation
    yearFolder = Format(currentTime, "YYYY")
    monthFolder = Format(currentTime, "mmmm") ' Get the full month name

    ' Define the archive folder path
    archivePath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\Archive\" & yearFolder & "\" & monthFolder & "\" & configSheet.Range("B4").Value & "\"

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ensure the archive folder exists
    If Not fso.FolderExists(archivePath) Then
        fso.CreateFolder archivePath
    End If

    ' Copy new_data files
    newDataFile = "new_data_*.xlsx" ' Use wildcard to find new data files
    newDataFilePath = Dir(sourceFolderPath & newDataFile) ' Get the first matching file

    Do While newDataFilePath <> ""
        fso.CopyFile sourceFolderPath & newDataFilePath, archivePath & newDataFilePath
        newDataFilePath = Dir ' Get the next matching file
    Loop

    ' Copy old_data files
    oldDataFile = "old_data_*.xlsx" ' Use wildcard to find old data files
    oldDataFilePath = Dir(sourceFolderPath & oldDataFile) ' Get the first matching file

    Do While oldDataFilePath <> ""
        fso.CopyFile sourceFolderPath & oldDataFilePath, archivePath & oldDataFilePath
        oldDataFilePath = Dir ' Get the next matching file
    Loop
    
End Sub

Sub DataTrimTieIn()
    Dim wsExport As Worksheet
    Dim wsImport As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim checkRange As Range
    Dim cell As Range
    
    Set wsExport = ThisWorkbook.Sheets("FileExport")
    Set wsImport = ThisWorkbook.Sheets("FileImport")
    
    ' Clear existing headers and data in FileExport
    wsExport.Rows(1).ClearContents
    wsExport.Cells.ClearContents
    
    ' Define the headers
    Dim headers As Variant
    headers = Array("Line Number", "Job Number", "Parent Job Number", "Order Number", "Description", "Notes", "POM Comments")
    
    ' Loop through the headers array and add them to the first row
    For i = LBound(headers) To UBound(headers)
        wsExport.Cells(1, i + 1).Value = headers(i)
    Next i
    
    ' Find the last row in FileImport
    lastRow = wsImport.Cells(wsImport.Rows.count, "A").End(xlUp).Row
    
    ' Copy data from FileImport to FileExport
    wsImport.Range("A2:A" & lastRow).Copy wsExport.Range("A2") ' Copy column A to column A
    wsImport.Range("J2:J" & lastRow).Copy wsExport.Range("B2") ' Copy column J to column B
    wsImport.Range("M2:M" & lastRow).Copy wsExport.Range("C2") ' Copy column M to column C
    wsImport.Range("B2:B" & lastRow).Copy wsExport.Range("D2") ' Copy column B to column D
    wsImport.Range("C2:C" & lastRow).Copy wsExport.Range("E2") ' Copy column C to column E
    wsImport.Range("D2:D" & lastRow).Copy wsExport.Range("F2") ' Copy column D to column F
    wsImport.Range("F2:F" & lastRow).Copy wsExport.Range("G2") ' Copy column F to column G
    
    ' Check for "-" in the data and replace it with an empty string
    lastRow = wsExport.Cells(wsExport.Rows.count, "A").End(xlUp).Row ' Update lastRow after copying data
    For i = 2 To lastRow ' Loop through rows starting from the second row
        Set checkRange = wsExport.Range(wsExport.Cells(i, 1), wsExport.Cells(i, 7)) ' Check columns A to G
        For Each cell In checkRange
            If cell.Value = "-" Then
                cell.Value = "" ' Replace "-" with an empty string
            End If
        Next cell
    Next i
End Sub

Sub DataTrimDar()
    Dim fileImportSheet As Worksheet
    Dim fileExportSheet As Worksheet
    Dim lastRow As Long

    ' Set the worksheets
    Set fileImportSheet = ThisWorkbook.Sheets("FileImport")
    Set fileExportSheet = ThisWorkbook.Sheets("FileExport")

    ' Find the last row in the FileImport sheet based on column A
    lastRow = fileImportSheet.Cells(fileImportSheet.Rows.count, "A").End(xlUp).Row

    ' Copy data from FileImport to FileExport for columns A to H
    fileImportSheet.Range("A1:A" & lastRow).Copy Destination:=fileExportSheet.Range("A1")
    fileImportSheet.Range("B1:B" & lastRow).Copy Destination:=fileExportSheet.Range("B1")
    fileImportSheet.Range("C1:C" & lastRow).Copy Destination:=fileExportSheet.Range("C1")
    fileImportSheet.Range("D1:D" & lastRow).Copy Destination:=fileExportSheet.Range("D1")
    fileImportSheet.Range("E1:E" & lastRow).Copy Destination:=fileExportSheet.Range("E1")
    fileImportSheet.Range("I1:I" & lastRow).Copy Destination:=fileExportSheet.Range("F1") ' Changed F to I
    fileImportSheet.Range("F1:F" & lastRow).Copy Destination:=fileExportSheet.Range("G1") ' Changed G to F
    fileImportSheet.Range("G1:G" & lastRow).Copy Destination:=fileExportSheet.Range("H1") ' Changed H to G

End Sub

Sub joblistDataTransfer()
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Dim lastRow As Long

    ' Set references to the source and destination sheets
    Set sourceSheet = ThisWorkbook.Sheets("FileExport")
    Set destinationSheet = ThisWorkbook.Sheets("TotalList")

    ' Find the last row in the source sheet
    lastRow = sourceSheet.Cells(sourceSheet.Rows.count, "A").End(xlUp).Row

    ' Copy data from columns A to G
    sourceSheet.Range("A1:G" & lastRow).Copy

    ' Paste data into the destination sheet starting from cell A1
    destinationSheet.Range("A1").PasteSpecial Paste:=xlPasteValues

    ' Clear the clipboard to free up memory
    Application.CutCopyMode = False

End Sub

Sub ExportUpdatedNewData()
    Dim fileExportSheet As Worksheet
    Dim configSheet As Worksheet
    Dim exportPath As String
    Dim timeStamp As Date
    Dim formattedDate As String
    Dim newFileName As String
    Dim newWorkbook As Workbook

    ' Set the worksheet references
    Set fileExportSheet = ThisWorkbook.Sheets("FileExport")
    Set configSheet = ThisWorkbook.Sheets("FileConfig")

    ' Define the export path
    exportPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\"

    ' Get the timestamp from cell B3
    timeStamp = configSheet.Range("B3").Value

    ' Format the timestamp for the file name
    formattedDate = Format(timeStamp, "MMDDYY_HHMM")

    ' Define the file name for the new data
    newFileName = exportPath & "new_data_" & formattedDate & ".xlsx"

    ' Create a new workbook to export the data
    Set newWorkbook = Workbooks.Add

    ' Copy data from FileExport to the new workbook
    fileExportSheet.Cells.Copy newWorkbook.Sheets(1).Cells

    ' Save the new data file
    newWorkbook.SaveAs fileName:=newFileName, FileFormat:=xlOpenXMLWorkbook

    ' Close the new workbook
    newWorkbook.Close SaveChanges:=False
End Sub

Sub CompareTimestamps()
    Dim ws As Worksheet
    Dim timeC2 As Date
    Dim timeB3 As Date
    
    ' Set the worksheet to configSheet
    Set ws = ThisWorkbook.Sheets("FileConfig")
    
    ' Get the time values from cells C2 and B3
    timeC2 = ws.Range("C2").Value
    timeB3 = ws.Range("B3").Value
    
    ' Compare the timestamps
    If timeB3 > timeC2 Then
        ws.Range("B5").Value = "Replace"
    Else
        ws.Range("B5").Value = "Keep"
    End If
End Sub

Sub NewToOld()
    Dim configSheet As Worksheet
    Dim sourceFolderPath As String
    Dim newDataFile As String
    Dim newDataFilePath As String
    Dim oldDataFilePath As String

    ' Set the config sheet
    Set configSheet = ThisWorkbook.Sheets("FileConfig")

    ' Define the source folder path
    sourceFolderPath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\"

    ' Specify the new_data file pattern
    newDataFile = "new_data_*.xlsx" ' Use wildcard to find the latest new data file
    newDataFilePath = Dir(sourceFolderPath & newDataFile) ' Get the first matching file

    If newDataFilePath <> "" Then
        ' Construct the old data file path
        oldDataFilePath = Replace(newDataFilePath, "new_data_", "old_data_")

        ' Rename the new data file to old data file
        Name sourceFolderPath & newDataFilePath As sourceFolderPath & oldDataFilePath

    Else
        MsgBox "No new data file found in the source directory.", vbExclamation
    End If
End Sub

Sub SelectNewData()
    Dim selectedFile As Variant
    Dim newDataSheet As Worksheet
    Dim fileDialog As fileDialog

    ' Set the NewData sheet
    Set newDataSheet = ThisWorkbook.Sheets("NewData")

    ' Prompt the user to select a file
    MsgBox "Select New_Data file", vbInformation

    ' Create a FileDialog object as a File Picker dialog box
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)

    ' Set the initial directory
    fileDialog.InitialFileName = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\"

    ' Set the file filter
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "Excel Files", "*.xls; *.xlsx"

    ' Show the dialog box and get the selected file
    If fileDialog.Show = -1 Then ' If the user selects a file
        selectedFile = fileDialog.SelectedItems(1) ' Get the selected file path

        ' Import data from the selected file
        Workbooks.Open selectedFile
        With ActiveWorkbook
            ' Copy all data from the first sheet
            .Sheets(1).UsedRange.Copy Destination:=newDataSheet.Range("A1")
            .Close SaveChanges:=False
        End With
    Else
        MsgBox "No file selected. Operation canceled.", vbExclamation
    End If
End Sub

Sub SelectOldData()
    Dim selectedFile As Variant
    Dim oldDataSheet As Worksheet
    Dim fileDialog As fileDialog

    ' Set the OldData sheet
    Set oldDataSheet = ThisWorkbook.Sheets("OldData")

    ' Prompt the user to select a file
    MsgBox "Select Old_Data file", vbInformation

    ' Create a FileDialog object as a File Picker dialog box
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)

    ' Set the initial directory
    fileDialog.InitialFileName = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\"

    ' Set the file filter
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "Excel Files", "*.xls; *.xlsx"

    ' Show the dialog box and get the selected file
    If fileDialog.Show = -1 Then ' If the user selects a file
        selectedFile = fileDialog.SelectedItems(1) ' Get the selected file path

        ' Import data from the selected file
        Workbooks.Open selectedFile
        With ActiveWorkbook
            ' Copy all data from the first sheet
            .Sheets(1).UsedRange.Copy Destination:=oldDataSheet.Range("A1")
            .Close SaveChanges:=False
        End With
    Else
        MsgBox "No file selected. Operation canceled.", vbExclamation
    End If
End Sub

Sub ReportTimeStamp()
    Dim fileConfigSheet As Worksheet

    ' Set the FileConfig sheet
    Set fileConfigSheet = ThisWorkbook.Sheets("FileConfig")

    ' Put the current date and time in cell B3
    fileConfigSheet.Cells(3, "B").Value = Now
End Sub

Sub UpdateFirstDarTimeStamp()
    Dim wsFirstDar As Worksheet
    Dim wsFileConfig As Worksheet
    Dim timeStamp As Variant
    Dim formattedTimeStamp As String

    ' Set the worksheets
    Set wsFirstDar = ThisWorkbook.Sheets("FirstDar")
    Set wsFileConfig = ThisWorkbook.Sheets("FileConfig")

    ' Get the timestamp from FileConfig cell B3
    timeStamp = wsFileConfig.Range("B3").Value

    ' Format the timestamp
    formattedTimeStamp = "Updated: " & Format(timeStamp, "MM/DD/YYYY HH:MM")

    ' Update cell X3 in FirstDar with the formatted timestamp
    wsFirstDar.Range("X3").Value = formattedTimeStamp
End Sub

Sub UpdateSecondDarTimeStamp()
    Dim wsSecondDar As Worksheet
    Dim wsFileConfig As Worksheet
    Dim timeStamp As Variant
    Dim formattedTimeStamp As String

    ' Set the worksheets
    Set wsSecondDar = ThisWorkbook.Sheets("SecondDar")
    Set wsFileConfig = ThisWorkbook.Sheets("FileConfig")

    ' Get the timestamp from FileConfig cell B3
    timeStamp = wsFileConfig.Range("B3").Value

    ' Format the timestamp
    formattedTimeStamp = "Updated: " & Format(timeStamp, "MM/DD/YYYY HH:MM")

    ' Update cell X3 in FirstDar with the formatted timestamp
    wsSecondDar.Range("P3").Value = formattedTimeStamp
End Sub

Sub UpdateThirdDarTimeStamp()
    Dim wsSecondDar As Worksheet
    Dim wsFileConfig As Worksheet
    Dim timeStamp As Variant
    Dim formattedTimeStamp As String

    ' Set the worksheets
    Set wsSecondDar = ThisWorkbook.Sheets("ThirdDar")
    Set wsFileConfig = ThisWorkbook.Sheets("FileConfig")

    ' Get the timestamp from FileConfig cell B3
    timeStamp = wsFileConfig.Range("B3").Value

    ' Format the timestamp
    formattedTimeStamp = "Updated: " & Format(timeStamp, "MM/DD/YYYY HH:MM")

    ' Update cell X3 in FirstDar with the formatted timestamp
    wsSecondDar.Range("P3").Value = formattedTimeStamp
End Sub

Sub NewOldDataImport()
    Dim fso As Object
    Dim newDataSheet As Worksheet
    Dim oldDataSheet As Worksheet
    Dim sourcePath As String
    Dim newDataFile As String
    Dim oldDataFile As String

    ' Set the worksheet references
    Set newDataSheet = ThisWorkbook.Sheets("NewData")
    Set oldDataSheet = ThisWorkbook.Sheets("OldData")

    ' Define the source path
    sourcePath = "\\nw\data\NewGen_TKR_FC\Industrial Engineering\Server 5S Structure\23.0 EDC\POM_Final\Build_Data\"

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Copy data from new_data* files to NewData sheet
    newDataFile = Dir(sourcePath & "new_data*") ' Get the first matching file
    Do While newDataFile <> ""
        ' Open the workbook and copy data
        With Workbooks.Open(sourcePath & newDataFile)
            .Sheets(1).UsedRange.Copy newDataSheet.Cells(1, 1) ' Paste starting at row 1
            .Close SaveChanges:=False
        End With
        newDataFile = Dir ' Get the next matching file
    Loop

    ' Copy data from old_data* files to OldData sheet
    oldDataFile = Dir(sourcePath & "old_data*") ' Get the first matching file
    Do While oldDataFile <> ""
        ' Open the workbook and copy data
        With Workbooks.Open(sourcePath & oldDataFile)
            .Sheets(1).UsedRange.Copy oldDataSheet.Cells(1, 1) ' Paste starting at row 1
            .Close SaveChanges:=False
        End With
        oldDataFile = Dir ' Get the next matching file
    Loop
End Sub

Sub CopyDataToCompare()
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Dim lastRow As Long

    ' Set the destination sheet to "Compare"
    Set destinationSheet = ThisWorkbook.Sheets("Compare")

    ' Copy data from "New Data" sheet
    Set sourceSheet = ThisWorkbook.Sheets("NewData")
    lastRow = sourceSheet.Cells(sourceSheet.Rows.count, "A").End(xlUp).Row
    sourceSheet.Range("A1:G" & lastRow).Copy destinationSheet.Range("A1")

    ' Copy data from "Old Data" sheet
    Set sourceSheet = ThisWorkbook.Sheets("OldData")
    lastRow = sourceSheet.Cells(sourceSheet.Rows.count, "A").End(xlUp).Row
    sourceSheet.Range("A1:G" & lastRow).Copy destinationSheet.Range("I1")
End Sub

Sub CheckForNewComments()
    Dim compareSheet As Worksheet
    Dim newDataSheet As Worksheet
    Dim lastRowCompare As Long
    Dim lastRowL As Long
    Dim lastRowQ As Long
    Dim valueA1 As Variant
    Dim valueB1 As Variant
    Dim valueA2 As Variant
    Dim valueB2 As Variant
    Dim foundMatch As Boolean
    Dim i As Long, j As Long

    ' Set the Compare and New Data sheets
    Set compareSheet = ThisWorkbook.Sheets("Compare")
    Set newDataSheet = ThisWorkbook.Sheets("NewData")

    ' Find the last row in the Compare sheet for column D (Value A)
    lastRowCompare = compareSheet.Cells(compareSheet.Rows.count, "D").End(xlUp).Row

    ' Find the last row in the Compare sheet for column L (Value A)
    lastRowL = compareSheet.Cells(compareSheet.Rows.count, "L").End(xlUp).Row

    ' Clear previous results in column Q of New Data
    newDataSheet.Range("Q2:Q" & newDataSheet.Rows.count).ClearContents

    ' Loop through each pair in column D and G of the Compare sheet
    For i = 2 To lastRowCompare ' Start from row 2
        valueA1 = compareSheet.Cells(i, "D").Value ' Value A from column D
        valueB1 = compareSheet.Cells(i, "G").Value ' Value B from column G
        foundMatch = False

        ' Loop through each pair in column L and O of the Compare sheet
        For j = 2 To lastRowL ' Start from row 2
            valueA2 = compareSheet.Cells(j, "L").Value ' Value A from column L
            valueB2 = compareSheet.Cells(j, "O").Value ' Value B from column O

            ' Check if Value A matches
            If valueA1 = valueA2 Then
                foundMatch = True
                ' Check if Value B is different
                If valueB1 <> valueB2 Then
                    newDataSheet.Cells(i, "Q").Value = "Comment" ' Set to Comment
                Else
                    newDataSheet.Cells(i, "Q").Value = "" ' Leave blank
                End If
                Exit For ' Exit the inner loop once a match is found
            End If
        Next j

        ' If no match was found for Value A
        If Not foundMatch Then
            newDataSheet.Cells(i, "Q").Value = "Comment" ' Set to Comment
        End If
    Next i

    ' Move the comment marker from New Data to Compare
    lastRowQ = newDataSheet.Cells(newDataSheet.Rows.count, "Q").End(xlUp).Row

    ' Check if there is data to copy
    If lastRowQ > 0 Then
        ' Copy data from column Q of New Data to column X of Compare
        newDataSheet.Range("Q1:Q" & lastRowQ).Copy
        compareSheet.Range("X1").PasteSpecial Paste:=xlPasteValues
    Else
        MsgBox "No data found in column Q of New Data to copy.", vbExclamation
    End If
End Sub

Sub emptyCommentUpdate()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheet to NewData
    Set ws = ThisWorkbook.Sheets("NewData")

    ' Find the last row in column Q
    lastRow = ws.Cells(ws.Rows.count, "Q").End(xlUp).Row

    ' Loop through each row in column Q
    For i = 1 To lastRow
        ' Check if there is a value in column Q and if column G is empty
        If ws.Cells(i, "Q").Value <> "" And ws.Cells(i, "G").Value = "" Then
            ' Delete the value in column Q
            ws.Cells(i, "Q").Value = ""
        End If
    Next i
End Sub

Sub CheckForNewJobs()
    Dim compareSheet As Worksheet
    Dim lastRowCompare As Long
    Dim lastRowL As Long
    Dim numberToCheck As Variant
    Dim found As Boolean
    Dim i As Long, j As Long

    ' Set the Compare sheet
    Set compareSheet = ThisWorkbook.Sheets("Compare")

    ' Find the last row in the Compare sheet for column D
    lastRowCompare = compareSheet.Cells(compareSheet.Rows.count, "D").End(xlUp).Row

    ' Find the last row in the Compare sheet for column L
    lastRowL = compareSheet.Cells(compareSheet.Rows.count, "L").End(xlUp).Row

    ' Loop through each number in column D of the Compare sheet
    For i = 2 To lastRowCompare ' Start from row 2
        numberToCheck = compareSheet.Cells(i, "D").Value
        found = False

        ' Check if the number exists in column L of the Compare sheet
        For j = 2 To lastRowL ' Start from row 2
            If compareSheet.Cells(j, "L").Value = numberToCheck Then
                found = True
                Exit For
            End If
        Next j

        ' If the number is not found, put "New" in column V of the Compare sheet
        If Not found Then
            compareSheet.Cells(i, "V").Value = "New"
        End If
    Next i
End Sub

Sub CheckForSells()
    Dim compareSheet As Worksheet
    Dim lastRowL As Long
    Dim lastRowD As Long
    Dim numberToCheck As Variant
    Dim found As Boolean
    Dim i As Long, j As Long

    ' Set the Compare sheet
    Set compareSheet = ThisWorkbook.Sheets("Compare")

    ' Find the last row in the Compare sheet for column L
    lastRowL = compareSheet.Cells(compareSheet.Rows.count, "L").End(xlUp).Row

    ' Find the last row in the Compare sheet for column D
    lastRowD = compareSheet.Cells(compareSheet.Rows.count, "D").End(xlUp).Row

    ' Loop through each number in column L of the Compare sheet
    For i = 2 To lastRowL ' Start from row 2
        numberToCheck = compareSheet.Cells(i, "L").Value
        found = False

        ' Check if the number exists in column D of the Compare sheet
        For j = 2 To lastRowD ' Start from row 2
            If compareSheet.Cells(j, "D").Value = numberToCheck Then
                found = True
                Exit For
            End If
        Next j

        ' If the number is not found, put "sell" in column W of the Compare sheet
        If Not found Then
            compareSheet.Cells(i, "W").Value = "sell"
        End If
    Next i
End Sub

Sub MoveToTotalList()
    Dim compareSheet As Worksheet
    Dim totalListSheet As Worksheet
    Dim lastRowCompare As Long
    Dim lastRowTotalList As Long
    Dim i As Long

    ' Set the Compare and TotalList sheets
    Set compareSheet = ThisWorkbook.Sheets("Compare")
    Set totalListSheet = ThisWorkbook.Sheets("TotalList")

    ' Copy data from columns A through G from Compare to TotalList
    lastRowCompare = compareSheet.Cells(compareSheet.Rows.count, "A").End(xlUp).Row
    compareSheet.Range("A1:G" & lastRowCompare).Copy totalListSheet.Range("A1")

    ' Copy values from column V and X from Compare to columns J and K in TotalList
    compareSheet.Range("V1:V" & lastRowCompare).Copy totalListSheet.Range("J1")
    compareSheet.Range("X1:X" & lastRowCompare).Copy totalListSheet.Range("K1")

    ' Loop through each row in column W of Compare
    For i = 2 To lastRowCompare ' Start from row 2
        If compareSheet.Cells(i, "W").Value <> "" Then
            ' Find the next blank row in TotalList
            lastRowTotalList = totalListSheet.Cells(totalListSheet.Rows.count, "A").End(xlUp).Row + 1
            
            ' Copy data from columns I through O of Compare to TotalList
            compareSheet.Range(compareSheet.Cells(i, "I"), compareSheet.Cells(i, "O")).Copy
            totalListSheet.Cells(lastRowTotalList, "A").PasteSpecial Paste:=xlPasteValues
            
            ' Put the word "Sell" in column L of TotalList
            totalListSheet.Cells(lastRowTotalList, "L").Value = "Sell"
        End If
    Next i
End Sub

Sub LnCorrect()
    Dim totalListSheet As Worksheet
    Dim airplaneListSheet As Worksheet
    Dim lastRowTotalList As Long
    Dim lastRowAirplaneList As Long
    Dim i As Long, j As Long
    Dim valueToCheck As Variant
    Dim matchFound As Boolean

    ' Set the sheets
    Set totalListSheet = ThisWorkbook.Sheets("TotalList")
    Set airplaneListSheet = ThisWorkbook.Sheets("Airplane List")

    ' Find the last row in TotalList and Airplane List
    lastRowTotalList = totalListSheet.Cells(totalListSheet.Rows.count, "A").End(xlUp).Row
    lastRowAirplaneList = airplaneListSheet.Cells(airplaneListSheet.Rows.count, "A").End(xlUp).Row

    ' Loop through each row in TotalList starting from row 2
    For i = lastRowTotalList To 2 Step -1 ' Loop backwards to avoid skipping rows after deletion
        valueToCheck = totalListSheet.Cells(i, "A").Value ' Value in column A of TotalList
        matchFound = False ' Reset matchFound for each row

        ' Loop through each row in Airplane List
        For j = 2 To lastRowAirplaneList
            ' Check if the value matches column A or B of Airplane List
            If valueToCheck = airplaneListSheet.Cells(j, "A").Value Or valueToCheck = airplaneListSheet.Cells(j, "B").Value Then
                ' Replace the value in column A of TotalList with the value in column C of Airplane List
                totalListSheet.Cells(i, "A").Value = airplaneListSheet.Cells(j, "C").Value
                matchFound = True ' Set matchFound to True if a match is found
                Exit For ' Exit the inner loop once a match is found
            End If
        Next j

        ' If no match was found, delete the entire row in TotalList
        If Not matchFound Then
            totalListSheet.Rows(i).Delete
        End If
    Next i
End Sub

Sub LnCorrectDAR()
    Dim fileExportSheet As Worksheet
    Dim airplaneListSheet As Worksheet
    Dim lastRowFileExport As Long
    Dim lastRowAirplaneList As Long
    Dim i As Long, j As Long
    Dim valueToCheck As Variant

    ' Set the sheets
    Set fileExportSheet = ThisWorkbook.Sheets("FileExport")
    Set airplaneListSheet = ThisWorkbook.Sheets("Airplane List")

    ' Find the last row in FileExport and Airplane List
    lastRowFileExport = fileExportSheet.Cells(fileExportSheet.Rows.count, "A").End(xlUp).Row
    lastRowAirplaneList = airplaneListSheet.Cells(airplaneListSheet.Rows.count, "A").End(xlUp).Row

    ' Loop through each row in FileExport starting from row 2
    For i = 2 To lastRowFileExport
        valueToCheck = fileExportSheet.Cells(i, "A").Value ' Value in column A of FileExport
        ' Loop through each row in Airplane List
        For j = 2 To lastRowAirplaneList
            ' Check if the value matches column A or B of Airplane List
            If valueToCheck = airplaneListSheet.Cells(j, "A").Value Or valueToCheck = airplaneListSheet.Cells(j, "B").Value Then
                ' Replace the value in column A of FileExport with the value in column C of Airplane List
                fileExportSheet.Cells(i, "A").Value = airplaneListSheet.Cells(j, "C").Value
                Exit For ' Exit the inner loop once a match is found
            End If
        Next j
    Next i
End Sub

Sub LNcorrectJoblist()
    Dim totalListSheet As Worksheet
    Dim airplaneListSheet As Worksheet
    Dim totalListLastRow As Long
    Dim airplaneListLastRow As Long
    Dim i As Long, j As Long
    Dim totalListValue As String
    Dim matchFound As Boolean

    ' Set references to the sheets
    Set totalListSheet = ThisWorkbook.Sheets("TotalList")
    Set airplaneListSheet = ThisWorkbook.Sheets("Airplane List")

    ' Find the last rows in both sheets
    totalListLastRow = totalListSheet.Cells(totalListSheet.Rows.count, "A").End(xlUp).Row
    airplaneListLastRow = airplaneListSheet.Cells(airplaneListSheet.Rows.count, "A").End(xlUp).Row

    ' Loop through each value in column A of TotalList
    For i = 1 To totalListLastRow
        totalListValue = totalListSheet.Cells(i, 1).Value
        matchFound = False
        
        ' Loop through each value in columns A and B of Airplane List
        For j = 1 To airplaneListLastRow
            If totalListValue = airplaneListSheet.Cells(j, 1).Value Or totalListValue = airplaneListSheet.Cells(j, 2).Value Then
                ' If a match is found, overwrite the value in TotalList with the value from column C
                totalListSheet.Cells(i, 1).Value = airplaneListSheet.Cells(j, 3).Value
                matchFound = True
                Exit For ' Exit the inner loop once a match is found
            End If
        Next j
    Next i

End Sub

Sub lostToolPlanUpdate()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheet to TotalList
    Set ws = ThisWorkbook.Sheets("TotalList")

    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.count, "F").End(xlUp).Row

    ' Loop through each row in column F
    For i = 1 To lastRow
        ' Check if "lost tool" is anywhere in the cell (case insensitive)
        If InStr(1, UCase(ws.Cells(i, "F").Value), "LOST TOOL") > 0 Then
            ' Replace the value in column B with "842-LOSTTOOL-FAKEPLAN"
            ws.Cells(i, "B").Value = "842-LOSTTOOL-FAKEPLAN"
        End If
    Next i
End Sub

Sub AssignSortNumberToJob()
    Dim jobSortSheet As Worksheet
    Dim jobSortSourceSheet As Worksheet
    Dim lastRowJobSort As Long
    Dim lastRowJobSortSource As Long
    Dim valueA1 As Variant
    Dim valueA2 As Variant
    Dim valueB1 As Variant
    Dim foundMatch As Boolean
    Dim i As Long, j As Long

    ' Set the sheets
    Set jobSortSheet = ThisWorkbook.Sheets("TotalList")
    Set jobSortSourceSheet = ThisWorkbook.Sheets("JobSort")

    ' Find the last row in column B of the Job Sort sheet
    lastRowJobSort = jobSortSheet.Cells(jobSortSheet.Rows.count, "B").End(xlUp).Row

    ' Find the last row in column A of the JobSort sheet
    lastRowJobSortSource = jobSortSourceSheet.Cells(jobSortSourceSheet.Rows.count, "A").End(xlUp).Row

    ' Clear previous data in column N of Job Sort
    jobSortSheet.Range("N2:N" & lastRowJobSort).ClearContents

    ' Loop through each value in column B of the Job Sort sheet
    For i = 2 To lastRowJobSort ' Start from row 2
        valueA1 = jobSortSheet.Cells(i, "B").Value ' Get the value from column B
        foundMatch = False

        ' Loop through each value in column A of the JobSort sheet
        For j = 2 To lastRowJobSortSource ' Start from row 2
            valueA2 = jobSortSourceSheet.Cells(j, "A").Value ' Get the value from column A
            valueB1 = jobSortSourceSheet.Cells(j, "D").Value ' Get the corresponding value from column D

            ' Check if the values match
            If valueA1 = valueA2 Then
                ' If they match, copy the value from column D to column N of Job Sort
                jobSortSheet.Cells(i, "N").Value = valueB1
                foundMatch = True
                Exit For ' Exit the inner loop once a match is found
            End If
        Next j

        ' If no match was found, assign the value from cell N82 on the JobSort sheet to column N
        If Not foundMatch Then
            jobSortSheet.Cells(i, "N").Value = jobSortSheet.Range("N82").Value
        End If
    Next i
    
End Sub

Sub IncreaseChildSortNumber()
    Dim totalListSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valueD As String

    ' Set the TotalList sheet
    Set totalListSheet = ThisWorkbook.Sheets("TotalList")

    ' Find the last row in column D of the TotalList sheet
    lastRow = totalListSheet.Cells(totalListSheet.Rows.count, "D").End(xlUp).Row

    ' Loop through each value in column D
    For i = 2 To lastRow ' Start from row 2
        valueD = Trim(totalListSheet.Cells(i, "D").Value) ' Get the value from column D and trim spaces

        ' Check if the cell is not empty
        If Len(valueD) > 0 Then
            ' Check if the first character is "N" or "U"
            If Left(valueD, 1) = "N" Or Left(valueD, 1) = "U" Then
                ' Increase the number in column N by 1
                totalListSheet.Cells(i, "N").Value = totalListSheet.Cells(i, "N").Value + 1
            End If
        End If
    Next i
End Sub

Sub AddJobType()
    ' Set the sheets
    Dim jobSortSheet As Worksheet
    Dim jobSortSourceSheet As Worksheet
    Dim lastRowJobSort As Long
    Dim lastRowJobSortSource As Long
    Dim i As Long, j As Long
    Dim valueA1 As Variant
    Dim valueA2 As Variant
    Dim valueB2 As Variant
    Dim foundMatch As Boolean

    Set jobSortSheet = ThisWorkbook.Sheets("TotalList")
    Set jobSortSourceSheet = ThisWorkbook.Sheets("JobSort")

    ' Find the last row in column B of the Job Sort sheet
    lastRowJobSort = jobSortSheet.Cells(jobSortSheet.Rows.count, "B").End(xlUp).Row

    ' Find the last row in column A of the JobSort sheet
    lastRowJobSortSource = jobSortSourceSheet.Cells(jobSortSourceSheet.Rows.count, "A").End(xlUp).Row

    ' Loop through each value in column B of the Job Sort sheet
    For i = 2 To lastRowJobSort ' Start from row 2
        valueA1 = jobSortSheet.Cells(i, "B").Value ' Get the value from column B
        foundMatch = False

        ' Loop through each value in column A of the JobSort sheet
        For j = 2 To lastRowJobSortSource ' Start from row 2
            valueA2 = jobSortSourceSheet.Cells(j, "A").Value ' Get the value from column A
            valueB2 = jobSortSourceSheet.Cells(j, "E").Value ' Get the corresponding value from column E

            ' Check if the values match
            If valueA1 = valueA2 Then
                ' If they match, copy the value from column E to column O of Job Sort
                jobSortSheet.Cells(i, "O").Value = valueB2
                foundMatch = True
                Exit For ' Exit the inner loop once a match is found
            End If
        Next j

        ' Optional: If no match was found, you can leave column O blank or add a message
        If Not foundMatch Then
            jobSortSheet.Cells(i, "O").Value = "MISC" ' Clear the cell in column O if no match is found
        End If
    Next i
End Sub

Sub SortDataByAirplaneAndJob()
    ' Set the worksheet
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets("TotalList")
    
    ' Find the last row of data in the sheet
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ' Check if there is data to sort
    If lastRow < 2 Then
        MsgBox "No data to sort."
        Exit Sub
    End If

    ' Select the range to sort (from A2 to Q of the last row)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("A2:A" & lastRow), Order:=xlAscending ' Sort by column A
        .SortFields.Add key:=ws.Range("N2:N" & lastRow), Order:=xlAscending ' Subsort by column N
        .SetRange ws.Range("A2:Q" & lastRow) ' Adjust the range to include columns A through Q
        .Header = xlNo ' No header in the range
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub UpdateJobTittle()

    Dim jobSortSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set jobSortSheet = ThisWorkbook.Sheets("TotalList")

    ' Find the last row in column E of the TotalList
    lastRow = jobSortSheet.Cells(jobSortSheet.Rows.count, "E").End(xlUp).Row

    ' Loop through each row in column F
    For i = 2 To lastRow ' Start from row 2
        ' Check if the cell in column F is empty
        If jobSortSheet.Cells(i, "F").Value = "" Then
            ' If it is empty, copy the value from column E to column F
            jobSortSheet.Cells(i, "F").Value = jobSortSheet.Cells(i, "E").Value
        End If
    Next i
End Sub

Sub flightOrderCount()
    Dim ws As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim uniqueValues As Collection
    Dim cell As Range
    Dim i As Long
    Dim outputColumn As Long
    Dim count As Long
    Dim searchPattern As String
    Dim labels As Variant

    ' Set the worksheet to TotalList and output to DataClean
    Set ws = ThisWorkbook.Sheets("TotalList")
    Set wsOutput = ThisWorkbook.Sheets("DataClean")

    ' Find the last row in column A of TotalList
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' Initialize a collection to store unique values
    Set uniqueValues = New Collection

    ' Loop through each cell in column A starting from row 2
    On Error Resume Next ' Ignore errors for duplicate keys
    For Each cell In ws.Range("A2:A" & lastRow)
        If cell.Value <> "" Then
            uniqueValues.Add cell.Value, CStr(cell.Value) ' Use the value as the key
        End If
    Next cell
    On Error GoTo 0 ' Resume normal error handling

    ' Set the starting column for output
    outputColumn = 12 ' Column L is the 12th column

    ' Define the labels to be added
    labels = Array("Preflight", "Postflight", "LMI", "Quickturn", "ThruFlight", "ThruLMI")

    ' Loop through the unique values and paste them in row 1 of DataClean
    For i = 1 To uniqueValues.count
        wsOutput.Cells(1, outputColumn).Value = uniqueValues(i)
        
        ' Add labels in the rows below the unique value
        Dim j As Long
        For j = 0 To UBound(labels)
            wsOutput.Cells(j + 1, outputColumn + 1).Value = labels(j) ' Place labels in the column to the right
        Next j
        
        ' Initialize count for 842-PREFLIGHT-STC-XXX
        count = 0
        
        ' Define the search pattern for 842-PREFLIGHT-STC-XXX
        searchPattern = "842-PREFLIGHT-STC-*"

        ' Loop through column B to count occurrences matching the pattern
        For Each cell In ws.Range("B2:B" & lastRow)
            If cell.Value Like searchPattern And cell.Offset(0, -1).Value = uniqueValues(i) Then
                ' Exclude the specific value and check if column L is empty
                If cell.Value <> "842-PREFLIGHT-STC-FERRY" And ws.Cells(cell.Row, "L").Value = "" Then
                    count = count + 1
                End If
            End If
        Next cell
        
        ' Output the count in the 3rd column of the airplane's list row 1
        wsOutput.Cells(1, outputColumn + 2).Value = count
        
        ' Initialize count for 842-POSTFLIGHT-STC-XXX
        count = 0
        
        ' Define the search pattern for 842-POSTFLIGHT-STC-XXX
        searchPattern = "842-POSTFLIGHT-STC-*"

        ' Loop through column B to count occurrences matching the pattern
        For Each cell In ws.Range("B2:B" & lastRow)
            If cell.Value Like searchPattern And cell.Offset(0, -1).Value = uniqueValues(i) Then
                ' Exclude the specific value and check if column L is empty
                If cell.Value <> "842-POSTFLIGHT-STC-FERRY" And ws.Cells(cell.Row, "L").Value = "" Then
                    count = count + 1
                End If
            End If
        Next cell
        
        ' Output the count in the 3rd column of the airplane's list row 2
        wsOutput.Cells(2, outputColumn + 2).Value = count
        
        ' Initialize count for 842-PREFLTLMI-STC-XXX
        count = 0
        
        ' Define the search pattern for 842-PREFLTLMI-STC-XXX
        searchPattern = "842-PREFLTLMI-STC-*"

        ' Loop through column B to count occurrences matching the pattern
        For Each cell In ws.Range("B2:B" & lastRow)
            If cell.Value Like searchPattern And cell.Offset(0, -1).Value = uniqueValues(i) Then
                ' Exclude the specific value and check if column L is empty
                If cell.Value <> "842-PREFLTLMI-STC-FERRY" And ws.Cells(cell.Row, "L").Value = "" Then
                    count = count + 1
                End If
            End If
        Next cell
        
        ' Output the count in the 3rd column of the airplane's list row 3
        wsOutput.Cells(3, outputColumn + 2).Value = count
        
        ' Initialize count for 842-QUICKTURN-STC-XXX
        count = 0
        
        ' Define the search pattern for 842-QUICKTURN-STC-XXX
        searchPattern = "842-QUICKTURN-STC-*"

        ' Loop through column B to count occurrences matching the pattern
        For Each cell In ws.Range("B2:B" & lastRow)
            If cell.Value Like searchPattern And cell.Offset(0, -1).Value = uniqueValues(i) Then
                ' Exclude the specific value and check if column L is empty
                If cell.Value <> "842-QUICKTURN-STC-FERRY" And ws.Cells(cell.Row, "L").Value = "" Then
                    count = count + 1
                End If
            End If
        Next cell
        
        ' Output the count in the 3rd column of the airplane's list row 4
        wsOutput.Cells(4, outputColumn + 2).Value = count
        
        ' Initialize count for 842-THRUFLIGHT-EDC-XXX
        count = 0
        
        ' Define the search pattern for 842-THRUFLIGHT-EDC-XXX
        searchPattern = "842-THRUFLIGHT-EDC-*"

        ' Loop through column B to count occurrences matching the pattern
        For Each cell In ws.Range("B2:B" & lastRow)
            If cell.Value Like searchPattern And cell.Offset(0, -1).Value = uniqueValues(i) Then
                ' Exclude the specific value and check if column L is empty
                If cell.Value <> "842-THRUFLIGHT-EDC-FERRY" And ws.Cells(cell.Row, "L").Value = "" Then
                    count = count + 1
                End If
            End If
        Next cell
        
        ' Output the count in the 3rd column of the airplane's list row 5
        wsOutput.Cells(5, outputColumn + 2).Value = count
        
        ' Initialize count for 842-THRUFLTLMI-STC-XX
        count = 0
        
        ' Define the search pattern for 842-THRUFLTLMI-STC-XX
        searchPattern = "842-THRUFLTLMI-STC-*"

        ' Loop through column B to count occurrences matching the pattern
        For Each cell In ws.Range("B2:B" & lastRow)
            If cell.Value Like searchPattern And cell.Offset(0, -1).Value = uniqueValues(i) Then
                ' Exclude the specific value and check if column L is empty
                If cell.Value <> "842-THRUFLTLMI-STC-FERRY" And ws.Cells(cell.Row, "L").Value = "" Then
                    count = count + 1
                End If
            End If
        Next cell
        
        ' Output the count in the 3rd column of the airplane's list row 6
        wsOutput.Cells(6, outputColumn + 2).Value = count
        
        outputColumn = outputColumn + 3 ' Move to the next column (3 columns over)
    Next i
End Sub

Sub flightOrderCleanListBuild()
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim wsKeys As Worksheet
    Dim Col As Integer
    Dim outputRow As Integer
    Dim lastRow As Long
    Dim key As String
    Dim matchFound As Boolean
    Dim lowestMatch As String
    Dim lowestDigits As Long
    Dim currentDigits As Long
    Dim i As Long
    Dim lastThreeChars As String

    ' Set the source worksheet to TotalList
    Set wsSource = ThisWorkbook.Worksheets("TotalList")
    ' Set the output worksheet to DataClean
    Set wsOutput = ThisWorkbook.Worksheets("Compare")
    ' Set the keys worksheet to DataClean
    Set wsKeys = ThisWorkbook.Worksheets("DataClean")
    
    ' Initialize the output row
    outputRow = 1 ' Start at AE1 (row 1)
    
    ' Initialize the column for keys (starting from column L which is 12)
    Col = 12
    
    ' Loop through the columns (L, O, R, ...)
    Do While wsKeys.Cells(1, Col).Value <> ""
        ' Get the key from the current column in DataClean
        key = wsKeys.Cells(1, Col).Value
        
        ' Get the last row in column A of the source sheet (TotalList)
        lastRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).Row
        
        ' Initialize match found flag and lowest match
        matchFound = False
        lowestMatch = ""
        lowestDigits = 9999 ' Set to a high value
        
        ' First, look for preflight matches in TotalList
        For i = 1 To lastRow
            ' Check if the key matches column A in TotalList
            If wsSource.Cells(i, 1).Value = key Then
                ' Check if column B matches the preflight format and column L is empty
                If wsSource.Cells(i, 2).Value Like "842-PREFLIGHT-STC-*" And _
                   wsSource.Cells(i, 2).Value <> "842-PREFLIGHT-STC-FERRY" And _
                   wsSource.Cells(i, 12).Value = "" Then
                   
                    ' Extract the last three characters from the matched B value
                    lastThreeChars = Right(wsSource.Cells(i, 2).Value, 3)
                    
                    ' Check if the last three characters are numeric
                    If IsNumeric(lastThreeChars) Then
                        currentDigits = CLng(lastThreeChars)
                        
                        ' Check if this is the lowest match
                        If currentDigits < lowestDigits Then
                            lowestDigits = currentDigits
                            lowestMatch = wsSource.Cells(i, 2).Value
                            matchFound = True
                        End If
                    End If
                End If
            End If
        Next i
        
        ' If a preflight match was found, display the key and the lowest match in DataClean
        If matchFound Then
            wsOutput.Cells(outputRow, 31).Value = key ' Column AE
            wsOutput.Cells(outputRow, 32).Value = lowestMatch ' Column AF
            outputRow = outputRow + 1
        End If
        
        ' Reset for postflight search
        matchFound = False
        lowestMatch = ""
        lowestDigits = 9999 ' Reset to a high value
        
        ' Now, look for postflight matches in TotalList
        For i = 1 To lastRow
            ' Check if the key matches column A in TotalList
            If wsSource.Cells(i, 1).Value = key Then
                ' Check if column B matches the postflight format and column L is empty
                If wsSource.Cells(i, 2).Value Like "842-POSTFLIGHT-STC-*" And _
                   wsSource.Cells(i, 2).Value <> "842-PREFLIGHT-STC-FERRY" And _
                   wsSource.Cells(i, 12).Value = "" Then
                   
                    ' Extract the last three characters from the matched B value
                    lastThreeChars = Right(wsSource.Cells(i, 2).Value, 3)
                    
                    ' Check if the last three characters are numeric
                    If IsNumeric(lastThreeChars) Then
                        currentDigits = CLng(lastThreeChars)
                        
                        ' Check if this is the lowest match
                        If currentDigits < lowestDigits Then
                            lowestDigits = currentDigits
                            lowestMatch = wsSource.Cells(i, 2).Value
                            matchFound = True
                        End If
                    End If
                End If
            End If
        Next i
        
        ' If a postflight match was found, display the key and the lowest match in DataClean
        If matchFound Then
            wsOutput.Cells(outputRow, 31).Value = key ' Column AE
            wsOutput.Cells(outputRow, 32).Value = lowestMatch ' Column AF
            outputRow = outputRow + 1
        End If
        
        ' Reset for prefltlmi search
        matchFound = False
        lowestMatch = ""
        lowestDigits = 9999 ' Reset to a high value
        
        ' Now, look for prefltlmi matches in TotalList
        For i = 1 To lastRow
            ' Check if the key matches column A in TotalList
            If wsSource.Cells(i, 1).Value = key Then
                ' Check if column B matches the prefltlmi format and column L is empty
                If wsSource.Cells(i, 2).Value Like "842-PREFLTLMI-STC-*" And _
                   wsSource.Cells(i, 2).Value <> "842-PREFLIGHT-STC-FERRY" And _
                   wsSource.Cells(i, 12).Value = "" Then
                   
                    ' Extract the last three characters from the matched B value
                    lastThreeChars = Right(wsSource.Cells(i, 2).Value, 3)
                    
                    ' Check if the last three characters are numeric
                    If IsNumeric(lastThreeChars) Then
                        currentDigits = CLng(lastThreeChars)
                        
                        ' Check if this is the lowest match
                        If currentDigits < lowestDigits Then
                            lowestDigits = currentDigits
                            lowestMatch = wsSource.Cells(i, 2).Value
                            matchFound = True
                        End If
                    End If
                End If
            End If
        Next i
        
        ' If a prefltlmi match was found, display the key and the lowest match in DataClean
        If matchFound Then
            wsOutput.Cells(outputRow, 31).Value = key ' Column AE
            wsOutput.Cells(outputRow, 32).Value = lowestMatch ' Column AF
            outputRow = outputRow + 1
        End If
        
        ' Reset for quickturn search
        matchFound = False
        lowestMatch = ""
        lowestDigits = 9999 ' Reset to a high value
        
        ' Now, look for quickturn matches in TotalList
        For i = 1 To lastRow
            ' Check if the key matches column A in TotalList
            If wsSource.Cells(i, 1).Value = key Then
                ' Check if column B matches the quickturn format and column L is empty
                If wsSource.Cells(i, 2).Value Like "842-QUICKTURN-STC-*" And _
                   wsSource.Cells(i, 2).Value <> "842-PREFLIGHT-STC-FERRY" And _
                   wsSource.Cells(i, 12).Value = "" Then
                   
                    ' Extract the last three characters from the matched B value
                    lastThreeChars = Right(wsSource.Cells(i, 2).Value, 3)
                    
                    ' Check if the last three characters are numeric
                    If IsNumeric(lastThreeChars) Then
                        currentDigits = CLng(lastThreeChars)
                        
                        ' Check if this is the lowest match
                        If currentDigits < lowestDigits Then
                            lowestDigits = currentDigits
                            lowestMatch = wsSource.Cells(i, 2).Value
                            matchFound = True
                        End If
                    End If
                End If
            End If
        Next i
        
        ' If a quickturn match was found, display the key and the lowest match in DataClean
        If matchFound Then
            wsOutput.Cells(outputRow, 31).Value = key ' Column AE
            wsOutput.Cells(outputRow, 32).Value = lowestMatch ' Column AF
            outputRow = outputRow + 1
        End If
        
        ' Reset for thruflight search
        matchFound = False
        lowestMatch = ""
        lowestDigits = 9999 ' Reset to a high value
        
        ' Now, look for thruflight matches in TotalList
        For i = 1 To lastRow
            ' Check if the key matches column A in TotalList
            If wsSource.Cells(i, 1).Value = key Then
                ' Check if column B matches the thruflight format and column L is empty
                If wsSource.Cells(i, 2).Value Like "842-THRUFLIGHT-EDC-*" And _
                   wsSource.Cells(i, 12).Value = "" Then
                   
                    ' Extract the last three characters from the matched B value
                    lastThreeChars = Right(wsSource.Cells(i, 2).Value, 3)
                    
                    ' Check if the last three characters are numeric
                    If IsNumeric(lastThreeChars) Then
                        currentDigits = CLng(lastThreeChars)
                        
                        ' Check if this is the lowest match
                        If currentDigits < lowestDigits Then
                            lowestDigits = currentDigits
                            lowestMatch = wsSource.Cells(i, 2).Value
                            matchFound = True
                        End If
                    End If
                End If
            End If
        Next i
        
        ' If a thruflight match was found, display the key and the lowest match in DataClean
        If matchFound Then
            wsOutput.Cells(outputRow, 31).Value = key ' Column AE
            wsOutput.Cells(outputRow, 32).Value = lowestMatch ' Column AF
            outputRow = outputRow + 1
        End If
        
        ' Reset for thruflti search
        matchFound = False
        lowestMatch = ""
        lowestDigits = 9999 ' Reset to a high value
        
        ' Now, look for thruflti matches in TotalList
        For i = 1 To lastRow
            ' Check if the key matches column A in TotalList
            If wsSource.Cells(i, 1).Value = key Then
                ' Check if column B matches the thruflti format and column L is empty
                If wsSource.Cells(i, 2).Value Like "842-THRUFLTLMI-STC-*" And _
                   wsSource.Cells(i, 12).Value = "" Then
                   
                    ' Extract the last three characters from the matched B value
                    lastThreeChars = Right(wsSource.Cells(i, 2).Value, 3)
                    
                    ' Check if the last three characters are numeric
                    If IsNumeric(lastThreeChars) Then
                        currentDigits = CLng(lastThreeChars)
                        
                        ' Check if this is the lowest match
                        If currentDigits < lowestDigits Then
                            lowestDigits = currentDigits
                            lowestMatch = wsSource.Cells(i, 2).Value
                            matchFound = True
                        End If
                    End If
                End If
            End If
        Next i
        
        ' If a thruflti match was found, display the key and the lowest match in DataClean
        If matchFound Then
            wsOutput.Cells(outputRow, 31).Value = key ' Column AE
            wsOutput.Cells(outputRow, 32).Value = lowestMatch ' Column AF
            outputRow = outputRow + 1
        End If
        
        ' Move to the next column (skip one column)
        Col = Col + 3
    Loop
End Sub

Sub flightOrderCleanJobDelete()
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim key As String
    Dim exceptions As Collection
    Dim matchInDataClean As Boolean
    Dim outputRow As Long

    ' Set the source worksheet to TotalList
    Set wsSource = ThisWorkbook.Worksheets("TotalList")
    ' Set the output worksheet to DataClean
    Set wsOutput = ThisWorkbook.Worksheets("Compare")
    
    ' Initialize the last row in column B of the source sheet
    lastRow = wsSource.Cells(wsSource.Rows.count, 2).End(xlUp).Row
    
    ' Initialize the collection for exceptions
    Set exceptions = New Collection
    exceptions.Add "842-PREFLIGHT-STC-FERRY"
    exceptions.Add "842-PREFLTLMI-STC-FERRY"
    
    ' Loop through each row in column B of TotalList
    For i = lastRow To 1 Step -1 ' Loop backwards to avoid skipping rows after deletion
        key = wsSource.Cells(i, 2).Value
        
        ' Check if the value matches any of the patterns in list A
        If key Like "842-PREFLIGHT-STC-*" Or _
           key Like "842-POSTFLIGHT-STC-*" Or _
           key Like "842-PREFLTLMI-STC-*" Or _
           key Like "842-QUICKTURN-STC-*" Or _
           key Like "842-THRUFLIGHT-EDC-*" Or _
           key Like "842-THRUFLTLMI-STC-*" Then
           
            ' Check for exceptions
            If key = exceptions(1) Or key = exceptions(2) Then
                ' Exception 1: Exact match with ferry values
                GoTo DeleteRow
            End If
            
            If wsSource.Cells(i, 12).Value <> "" Then
                ' Exception 2: Column L has a value
                GoTo DeleteRow
            End If
            
            ' Exception 3: Check if column A and B match with column AE and AF on DataClean
            matchInDataClean = False
            outputRow = 1 ' Start checking from the first row in DataClean
            Do While wsOutput.Cells(outputRow, 31).Value <> "" ' Column AE
                If wsSource.Cells(i, 1).Value = wsOutput.Cells(outputRow, 31).Value And _
                   wsSource.Cells(i, 2).Value = wsOutput.Cells(outputRow, 32).Value Then
                    matchInDataClean = True
                    Exit Do
                End If
                outputRow = outputRow + 1
            Loop
            
            If matchInDataClean Then
                ' Exception 3: A match was found in DataClean
                GoTo DeleteRow
            End If
            
            ' If no exceptions apply, delete the row
            wsSource.Rows(i).Delete
        End If
DeleteRow:
    Next i
End Sub

Sub UpdateJobTitleDAR()

    Dim jobExportSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set jobExportSheet = ThisWorkbook.Sheets("FileExport")

    ' Find the last row in column C of the FileExport sheet
    lastRow = jobExportSheet.Cells(jobExportSheet.Rows.count, "C").End(xlUp).Row

    ' Loop through each row in column D
    For i = 2 To lastRow ' Start from row 2
        ' Check if the cell in column D is empty
        If jobExportSheet.Cells(i, "D").Value = "" Then
            ' If it is empty, copy the value from column C to column D
            jobExportSheet.Cells(i, "D").Value = jobExportSheet.Cells(i, "C").Value
        End If
    Next i
End Sub

Sub MoveDataToDataClean()
    ' Set the worksheets
    Dim totalListSheet As Worksheet
    Dim dataCleanSheet As Worksheet
    Dim lastRowTotalList As Long
    Dim i As Long
    Dim pasteRow As Long

    Set totalListSheet = ThisWorkbook.Sheets("TotalList")
    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Find the last row in TotalList
    lastRowTotalList = totalListSheet.Cells(totalListSheet.Rows.count, "A").End(xlUp).Row

    ' Set the starting row for pasting in DataClean
    pasteRow = 11

    ' Loop through each row in TotalList and copy data to DataClean
    For i = 2 To lastRowTotalList ' Assuming row 1 is a header
        dataCleanSheet.Cells(pasteRow, "A").Value = totalListSheet.Cells(i, "A").Value
        dataCleanSheet.Cells(pasteRow, "B").Value = totalListSheet.Cells(i, "D").Value
        dataCleanSheet.Cells(pasteRow, "D").Value = totalListSheet.Cells(i, "F").Value
        dataCleanSheet.Cells(pasteRow, "E").Value = totalListSheet.Cells(i, "G").Value
        dataCleanSheet.Cells(pasteRow, "F").Value = totalListSheet.Cells(i, "J").Value
        dataCleanSheet.Cells(pasteRow, "G").Value = totalListSheet.Cells(i, "K").Value
        dataCleanSheet.Cells(pasteRow, "H").Value = totalListSheet.Cells(i, "L").Value
        dataCleanSheet.Cells(pasteRow, "I").Value = totalListSheet.Cells(i, "M").Value
        dataCleanSheet.Cells(pasteRow, "J").Value = totalListSheet.Cells(i, "N").Value
        dataCleanSheet.Cells(pasteRow, "K").Value = totalListSheet.Cells(i, "O").Value

        pasteRow = pasteRow + 1
    Next i
End Sub

Sub AddColon()
    Dim dataCleanSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the Data Clean sheet
    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Find the last row in column B of the Data Clean sheet
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "B").End(xlUp).Row

    ' Loop through each row in column B starting from row 1
    For i = 1 To lastRow
        ' Check if the current cell in column B is not blank
        If dataCleanSheet.Cells(i, "B").Value <> "" Then
            ' Append ":" to the value in column B
            dataCleanSheet.Cells(i, "B").Value = dataCleanSheet.Cells(i, "B").Value & ":"
        End If
    Next i
End Sub

Sub UpdateSoldNotes()
    ' Set the Data clean sheet
    Dim dataCleanSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Find the last row in column H of the Data clean sheet
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "H").End(xlUp).Row

    ' Loop through each row in column H
    For i = 2 To lastRow ' Start from row 2
        ' Check if the cell in column H contains the word "Sell"
        If InStr(1, dataCleanSheet.Cells(i, "H").Value, "Sell", vbTextCompare) > 0 Then
            ' If it does, replace the corresponding cell in column E with "Sold"
            dataCleanSheet.Cells(i, "E").Value = "Sold"
        End If
    Next i
End Sub

Sub TextColorUpdate()
    ' Set the Data clean sheet
    Dim dataCleanSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Find the last row in column A of the Data clean sheet
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "A").End(xlUp).Row

    ' Loop through each row to check values in columns F, G, and H
    For i = 2 To lastRow ' Start from row 2
        ' Check if there is a value in column F
        If dataCleanSheet.Cells(i, "F").Value <> "" Then
            ' Make text in columns B and D Bold and Green
            With dataCleanSheet
                .Cells(i, "B").Font.Bold = True
                .Cells(i, "B").Font.Color = RGB(0, 128, 0) ' Green
                .Cells(i, "D").Font.Bold = True
                .Cells(i, "D").Font.Color = RGB(0, 128, 0) ' Green
            End With
        End If
        
        ' Check if there is a value in column G
        If dataCleanSheet.Cells(i, "G").Value <> "" Then
            ' Make text in column E Bold and Blue
            With dataCleanSheet
                .Cells(i, "E").Font.Bold = True
                .Cells(i, "E").Font.Color = RGB(0, 0, 255) ' Blue
            End With
        End If
        
        ' Check if there is a value in column H
        If dataCleanSheet.Cells(i, "H").Value <> "" Then
            ' Make text in columns B, D, and E Bold and Purple
            With dataCleanSheet
                .Cells(i, "B").Font.Bold = True
                .Cells(i, "B").Font.Color = RGB(128, 0, 128) ' Purple
                .Cells(i, "D").Font.Bold = True
                .Cells(i, "D").Font.Color = RGB(128, 0, 128) ' Purple
                .Cells(i, "E").Font.Bold = True
                .Cells(i, "E").Font.Color = RGB(128, 0, 128) ' Purple
            End With
        End If
    Next i
End Sub

Sub InsertLinesBetweenAirplanes()
    Dim dataCleanSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the Data clean sheet
    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Insert 3 lines between different values in column A
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "A").End(xlUp).Row

    ' Loop through the rows in reverse order starting from the last row
    For i = lastRow To 2 Step -1 ' Start from row 2 to avoid inserting above the header
        ' Check if the current cell is different from the cell above it
        If dataCleanSheet.Cells(i, "A").Value <> dataCleanSheet.Cells(i - 1, "A").Value Then
            ' Insert three rows above the current row
            dataCleanSheet.Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            dataCleanSheet.Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            dataCleanSheet.Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
    Next i

End Sub

Sub CreateAirplaneHeader()
    ' Set the sheets
    Dim dataCleanSheet As Worksheet
    Dim airplaneListSheet As Worksheet
    Dim lastRow As Long
    Dim lastRowAirplaneList As Long
    Dim i As Long, j As Long
    Dim valueA As Variant
    Dim valueC As Variant
    Dim valueD As Variant
    Dim foundMatch As Boolean

    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")
    Set airplaneListSheet = ThisWorkbook.Sheets("Airplane List")

    ' Find the last row in column A of the Data clean sheet
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "A").End(xlUp).Row

    ' Move the LN to the top of the job chunk
    For i = 2 To lastRow ' Start from row 2
        ' Check if the current cell in column A is different from the cell above it and is not blank
        If dataCleanSheet.Cells(i, "A").Value <> dataCleanSheet.Cells(i - 1, "A").Value And _
           dataCleanSheet.Cells(i, "A").Value <> "" Then
            ' Check if we can paste two rows above (i.e., i must be greater than 2)
            If i > 2 Then
                ' Copy the value from column A to column B of the row two above
                dataCleanSheet.Cells(i - 2, "B").Value = dataCleanSheet.Cells(i, "A").Value
            End If
        End If
    Next i

    ' Find the last row in column C of the Airplane List sheet
    lastRowAirplaneList = airplaneListSheet.Cells(airplaneListSheet.Rows.count, "C").End(xlUp).Row

    ' Loop through each value in column B of the Data clean sheet
    For i = 1 To lastRow ' Start from row 1
        valueA = dataCleanSheet.Cells(i, "B").Value ' Get the value from column B
        foundMatch = False

        ' Loop through each value in column C of the Airplane List sheet
        For j = 2 To lastRowAirplaneList ' Start from row 2
            valueC = airplaneListSheet.Cells(j, "C").Value ' Get the value from column C
            valueD = airplaneListSheet.Cells(j, "D").Value ' Get the corresponding value from column D

            ' Check if the values match
            If valueA = valueC Then
                ' If they match, replace the value in column B with the value from column D
                dataCleanSheet.Cells(i, "B").Value = valueD
                
                ' Change the font color to black, set font size to 18, make it bold, underlined, and highlight
                With dataCleanSheet.Cells(i, "B")
                    .Font.Color = RGB(0, 0, 0) ' Black color
                    .Font.Size = 18
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                    .Interior.Color = RGB(255, 255, 0) ' Highlight with yellow color
                End With
                
                foundMatch = True
                Exit For ' Exit the inner loop once a match is found
            End If
        Next j
        
        ' Merge and center cells from column B to D if a match was found
        If foundMatch Then
            With dataCleanSheet.Range(dataCleanSheet.Cells(i, "B"), dataCleanSheet.Cells(i, "D"))
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If
    Next i
End Sub

Sub InsertLinesBetweenJobTypes()
    Dim dataCleanSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the Data clean sheet
    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Find the last row in column K of the Data clean sheet
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "K").End(xlUp).Row

    ' Loop through the rows in reverse order starting from the last row
    For i = lastRow To 2 Step -1 ' Start from row 2 to avoid inserting above the header
        ' Check if the current cell in column K is not blank and different from the cell above it
        If dataCleanSheet.Cells(i, "K").Value <> "" And _
           dataCleanSheet.Cells(i, "K").Value <> dataCleanSheet.Cells(i - 1, "K").Value Then
            ' Insert one row above the current row
            dataCleanSheet.Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
    Next i
End Sub

Sub CreateJobHeader()
    Dim dataCleanSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the Data Clean sheet
    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Find the last row in column K of the Data Clean sheet
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "K").End(xlUp).Row

    ' Loop through each row in column K starting from row 2
    For i = 2 To lastRow ' Start from row 2 to avoid inserting above the header
        ' Check if the current cell in column K is not blank and different from the cell above it
        If dataCleanSheet.Cells(i, "K").Value <> "" And _
           dataCleanSheet.Cells(i, "K").Value <> dataCleanSheet.Cells(i - 1, "K").Value Then
            
            ' Copy the value from column K to column B of the row above
            dataCleanSheet.Cells(i - 1, "B").Value = dataCleanSheet.Cells(i, "K").Value
            
            ' Merge and center columns B through D for the row above
            With dataCleanSheet.Range(dataCleanSheet.Cells(i - 1, "B"), dataCleanSheet.Cells(i - 1, "D"))
                .Merge
                .HorizontalAlignment = xlCenter ' Center the text
                .VerticalAlignment = xlCenter ' Center the text vertically
            End With
            
            ' Format the text in column B
            With dataCleanSheet.Cells(i - 1, "B").Font
                .Color = RGB(255, 165, 0) ' Orange color
                .Bold = True
                .Size = 14 ' Set font size to 14
                .Underline = xlUnderlineStyleSingle ' Underline the text
            End With
            
            ' Align the text to the left
            dataCleanSheet.Cells(i - 1, "B").HorizontalAlignment = xlHAlignLeft
        End If
    Next i
End Sub

Sub MoveComments()
    Dim dataCleanSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the Data Clean sheet
    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Find the last row in column E of the Data Clean sheet
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "E").End(xlUp).Row

    ' Loop through each row in column E starting from the last row to avoid shifting issues
    For i = lastRow To 1 Step -1 ' Start from the last row
        ' Check if the current cell in column E is not blank
        If dataCleanSheet.Cells(i, "E").Value <> "" Then
            ' Insert a new row below the current row
            dataCleanSheet.Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            ' Copy the data from column E to column D of the newly inserted row
            dataCleanSheet.Cells(i + 1, "D").Value = dataCleanSheet.Cells(i, "E").Value
            
            ' Copy the formatting from the original row to the new row
            dataCleanSheet.Rows(i).Copy
            dataCleanSheet.Rows(i + 1).PasteSpecial Paste:=xlPasteFormats
            
            ' Insert the character • in column B of the newly inserted row
            dataCleanSheet.Cells(i + 1, "B").Value = "•"
            
            ' Set the text alignment to right for column B
            dataCleanSheet.Cells(i + 1, "B").HorizontalAlignment = xlHAlignRight
            
            ' Set the vertical alignment to middle for column B
            dataCleanSheet.Cells(i + 1, "B").VerticalAlignment = xlVAlignCenter
            
            ' Clear the clipboard to avoid the marching ants around the copied range
            Application.CutCopyMode = False
        End If
    Next i
End Sub

Sub CreateHeaderSection()
    Dim dataCleanSheet As Worksheet
    Dim fileConfigSheet As Worksheet
    Dim countJobSells As Long
    Dim countJobAdds As Long
    Dim pomData As String

    ' Set the Data Clean sheet and FileConfig sheet
    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")
    Set fileConfigSheet = ThisWorkbook.Sheets("FileConfig")

    ' Merge and center each row from 4 to 11 for columns B through D
    Dim i As Long
    For i = 4 To 11
        With dataCleanSheet.Range(dataCleanSheet.Cells(i, "B"), dataCleanSheet.Cells(i, "D"))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i

    ' Set text and format for cell B4
    With dataCleanSheet.Cells(4, "B")
        .Value = "Comment updates in Blue"
        .Font.Color = RGB(0, 0, 255) ' Blue color
        .Font.Bold = True
        .Font.Size = 18
    End With

    ' Set text and format for cell B5
    With dataCleanSheet.Cells(5, "B")
        .Value = "Completed Jobs in Purple"
        .Font.Color = RGB(128, 0, 128) ' Purple color
        .Font.Bold = True
        .Font.Size = 18
    End With

    ' Set text and format for cell B6
    With dataCleanSheet.Cells(6, "B")
        .Value = "New Job Adds in Green"
        .Font.Color = RGB(0, 128, 0) ' Green color
        .Font.Bold = True
        .Font.Size = 18
    End With

    ' Count the number of non-blank values in column H
    countJobSells = Application.WorksheetFunction.CountA(dataCleanSheet.Range("H:H"))
    ' Set text and format for cell B9
    With dataCleanSheet.Cells(9, "B")
        .Value = "Job Sells: " & countJobSells
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 14
    End With

    ' Count the number of non-blank values in column F
    countJobAdds = Application.WorksheetFunction.CountA(dataCleanSheet.Range("F:F"))
    ' Set text and format for cell B10
    With dataCleanSheet.Cells(10, "B")
        .Value = "Job Adds: " & countJobAdds
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 14
    End With

    ' Get the time from cell B3 on the FileConfig sheet
    pomData = fileConfigSheet.Cells(3, "B").Value
    ' Set text and format for cell B11
    With dataCleanSheet.Cells(11, "B")
        .Value = "POM Data as of " & pomData
        .HorizontalAlignment = xlLeft ' Align text to the left
        .VerticalAlignment = xlCenter
        .Font.Size = 10
    End With

    ' Auto adjust the height of each row from 4 to 11
    For i = 4 To 11
        dataCleanSheet.Rows(i).AutoFit
    Next i

    ' Logic for deleting rows based on counts
    If countJobSells = 0 Then
        dataCleanSheet.Rows(9).Delete
    End If

    If countJobSells > 0 And countJobAdds = 0 Then
        dataCleanSheet.Rows(10).Delete
    End If

    If countJobSells = 0 And countJobAdds = 0 Then
        dataCleanSheet.Rows(9).Delete
    End If
    
End Sub

Sub TextFormat()
    Dim dataCleanSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the Data Clean sheet
    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Find the last row in column D of the Data Clean sheet
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "D").End(xlUp).Row

    ' Set the text in column D to wrap
    dataCleanSheet.Range("D1:D" & lastRow).WrapText = True

    ' Turn off wrap text in column E
    dataCleanSheet.Range("E1:E" & lastRow).WrapText = False

    ' Align all non-merged text in column B to the right
    For i = 1 To lastRow
        ' Check if the cell in column B is not merged
        If Not dataCleanSheet.Cells(i, "B").MergeCells Then
            dataCleanSheet.Cells(i, "B").HorizontalAlignment = xlHAlignRight
        End If
    Next i

    ' Loop through all cells in column A to find the first non-blank value
    For i = 1 To lastRow
        If dataCleanSheet.Cells(i, "A").Value <> "" Then
            ' Place "." in column C for that row and set text color to white
            dataCleanSheet.Cells(i, "C").Value = "."
            dataCleanSheet.Cells(i, "C").Font.Color = RGB(255, 255, 255) ' Set text color to white
            Exit For ' Exit the loop after the first match
        End If
    Next i

    ' Check each row for the condition in columns F and G
    For i = 1 To lastRow
        If dataCleanSheet.Cells(i, "F").Value = "" And dataCleanSheet.Cells(i, "G").Value <> "" Then
            ' Make the text in column D bold and blue in the row below the match
            If i + 1 <= lastRow Then ' Ensure we don't go out of bounds
                With dataCleanSheet.Cells(i + 1, "D")
                    .Font.Bold = True
                    .Font.Color = RGB(0, 0, 255) ' Set text color to blue
                End With
            End If
        End If
    Next i
End Sub

Sub addFlightOrderLeftCommentPreflight()
    Dim wsDataClean As Worksheet
    Dim wsJobSort As Worksheet
    Dim lastRowDataClean As Long
    Dim lastRowJobSort As Long
    Dim i As Long
    Dim j As Long
    Dim matchFound As Boolean
    Dim matchInKeys As Boolean
    Dim key As String
    Dim xValue As Variant
    Dim outputRow As Long

    ' Set the worksheets
    Set wsDataClean = ThisWorkbook.Worksheets("DataClean")
    Set wsJobSort = ThisWorkbook.Worksheets("JobSort")
    
    ' Get the last row in column J of DataClean
    lastRowDataClean = wsDataClean.Cells(wsDataClean.Rows.count, 10).End(xlUp).Row
    ' Get the last row in column H of JobSort
    lastRowJobSort = wsJobSort.Cells(wsJobSort.Rows.count, 8).End(xlUp).Row
    
    ' Loop through each cell in column J of DataClean
    For i = 1 To lastRowDataClean
        ' Check if the cell in column J is not blank
        If wsDataClean.Cells(i, 10).Value <> "" Then
            matchFound = False
            ' Loop through each cell in column H of JobSort
            For j = 1 To lastRowJobSort
                ' Check if the cell in column H is not blank
                If wsJobSort.Cells(j, 8).Value <> "" Then
                    ' Check for a match
                    If wsDataClean.Cells(i, 10).Value = wsJobSort.Cells(j, 8).Value Then
                        matchFound = True
                        Exit For
                    End If
                End If
            Next j
            
            ' If a match is found and column H of DataClean is blank, insert a new line below
            If matchFound And wsDataClean.Cells(i, 8).Value = "" Then ' Check if column H of DataClean is blank
                wsDataClean.Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                ' Add bullet character to column B of the new row
                wsDataClean.Cells(i + 1, 2).Value = "•" ' Column B is 2
                
                ' Get the value from column A of the matched row
                key = wsDataClean.Cells(i, 1).Value
                
                ' Check for a match in columns L, O, R, etc.
                matchInKeys = False
                Dim colKey As Integer
                colKey = 12 ' Start from column L (12)
                
                Do While wsDataClean.Cells(1, colKey).Value <> ""
                    If wsDataClean.Cells(1, colKey).Value = key Then
                        matchInKeys = True
                        ' Get the value from two columns over (Z)
                        xValue = wsDataClean.Cells(1, colKey + 2).Value ' Column Z is two columns over from L
                        Exit Do
                    End If
                    colKey = colKey + 3 ' Move to the next key column (skip one column)
                Loop
                
                ' If a match was found, add the text to cell D of the new row
                If matchInKeys Then
                    If xValue = 1 Then
                        wsDataClean.Cells(i + 1, 4).Value = "This is the last preflight order" ' Column D is 4
                    Else
                        wsDataClean.Cells(i + 1, 4).Value = "There are " & xValue & " Preflight orders left" ' Column D is 4
                    End If
                End If
            End If
        End If
    Next i
End Sub

Sub addFlightOrderLeftCommentPostflight()
    Dim wsDataClean As Worksheet
    Dim wsJobSort As Worksheet
    Dim lastRowDataClean As Long
    Dim lastRowJobSort As Long
    Dim i As Long
    Dim j As Long
    Dim matchFound As Boolean
    Dim matchInKeys As Boolean
    Dim key As String
    Dim xValue As Variant
    Dim outputRow As Long
    Dim colKey As Integer

    ' Set the worksheets
    Set wsDataClean = ThisWorkbook.Worksheets("DataClean")
    Set wsJobSort = ThisWorkbook.Worksheets("JobSort")
    
    ' Get the last row in column J of DataClean
    lastRowDataClean = wsDataClean.Cells(wsDataClean.Rows.count, 10).End(xlUp).Row
    ' Get the last row in column J of JobSort
    lastRowJobSort = wsJobSort.Cells(wsJobSort.Rows.count, 10).End(xlUp).Row
    
    ' Loop through each cell in column J of DataClean
    For i = 1 To lastRowDataClean
        ' Check if the cell in column J is not blank
        If wsDataClean.Cells(i, 10).Value <> "" Then
            matchFound = False
            ' Loop through each cell in column J of JobSort
            For j = 1 To lastRowJobSort
                ' Check if the cell in column J is not blank
                If wsJobSort.Cells(j, 10).Value <> "" Then
                    ' Check for a match
                    If wsDataClean.Cells(i, 10).Value = wsJobSort.Cells(j, 10).Value Then
                        matchFound = True
                        Exit For
                    End If
                End If
            Next j
            
            ' If a match is found and column H of DataClean is blank, insert a new line below
            If matchFound And wsDataClean.Cells(i, 8).Value = "" Then ' Check if column H of DataClean is blank
                wsDataClean.Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                ' Add bullet character to column B of the new row
                wsDataClean.Cells(i + 1, 2).Value = "•" ' Column B is 2
                
                ' Get the value from column A of the matched row
                key = wsDataClean.Cells(i, 1).Value ' Key from column A
                
                ' Check for a match in columns L, O, R, etc.
                matchInKeys = False
                colKey = 12 ' Start from column L (12)
                
                Do While wsDataClean.Cells(1, colKey).Value <> ""
                    If wsDataClean.Cells(1, colKey).Value = key Then
                        matchInKeys = True
                        ' Get the value from two columns over and one row down
                        xValue = wsDataClean.Cells(2, colKey + 2).Value ' Column N for L1, P for O1, etc.
                        Exit Do
                    End If
                    colKey = colKey + 3 ' Move to the next key column (skip one column)
                Loop
                
                ' If a match was found, add the text to cell D of the new row
                If matchInKeys Then
                    If xValue = 1 Then
                        wsDataClean.Cells(i + 1, 4).Value = "This is the last postflight order" ' Column D is 4
                    Else
                        wsDataClean.Cells(i + 1, 4).Value = "There are " & xValue & " Postflight orders left" ' Column D is 4
                    End If
                End If
            End If
        End If
    Next i
End Sub

Sub addFlightOrderLeftCommentLMI()
    Dim wsDataClean As Worksheet
    Dim wsJobSort As Worksheet
    Dim lastRowDataClean As Long
    Dim lastRowJobSort As Long
    Dim i As Long
    Dim j As Long
    Dim matchFound As Boolean
    Dim matchInKeys As Boolean
    Dim key As String
    Dim xValue As Variant
    Dim outputRow As Long
    Dim colKey As Integer

    ' Set the worksheets
    Set wsDataClean = ThisWorkbook.Worksheets("DataClean")
    Set wsJobSort = ThisWorkbook.Worksheets("JobSort")
    
    ' Get the last row in column J of DataClean
    lastRowDataClean = wsDataClean.Cells(wsDataClean.Rows.count, 10).End(xlUp).Row
    ' Get the last row in column I of JobSort
    lastRowJobSort = wsJobSort.Cells(wsJobSort.Rows.count, 9).End(xlUp).Row
    
    ' Loop through each cell in column J of DataClean
    For i = 1 To lastRowDataClean
        ' Check if the cell in column J is not blank
        If wsDataClean.Cells(i, 10).Value <> "" Then
            matchFound = False
            ' Loop through each cell in column I of JobSort
            For j = 1 To lastRowJobSort
                ' Check if the cell in column I is not blank
                If wsJobSort.Cells(j, 9).Value <> "" Then
                    ' Check for a match
                    If wsDataClean.Cells(i, 10).Value = wsJobSort.Cells(j, 9).Value Then
                        matchFound = True
                        Exit For
                    End If
                End If
            Next j
            
            ' If a match is found and column H of DataClean is blank, insert a new line below
            If matchFound And wsDataClean.Cells(i, 8).Value = "" Then ' Check if column H of DataClean is blank
                wsDataClean.Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                ' Add bullet character to column B of the new row
                wsDataClean.Cells(i + 1, 2).Value = "•" ' Column B is 2
                
                ' Get the value from column A of the matched row
                key = wsDataClean.Cells(i, 1).Value ' Key from column A
                
                ' Check for a match in columns L, O, R, etc.
                matchInKeys = False
                colKey = 12 ' Start from column L (12)
                
                Do While wsDataClean.Cells(1, colKey).Value <> ""
                    If wsDataClean.Cells(1, colKey).Value = key Then
                        matchInKeys = True
                        ' Get the value from three columns over (N for L1, T for R1, etc.)
                        xValue = wsDataClean.Cells(3, colKey + 2).Value ' Column N for L1, T for R1, etc.
                        Exit Do
                    End If
                    colKey = colKey + 3 ' Move to the next key column (skip one column)
                Loop
                
                ' If a match was found, add the text to cell D of the new row
                If matchInKeys Then
                    If xValue = 1 Then
                        wsDataClean.Cells(i + 1, 4).Value = "This is the last LMI order" ' Column D is 4
                    Else
                        wsDataClean.Cells(i + 1, 4).Value = "There are " & xValue & " LMI orders left" ' Column D is 4
                    End If
                End If
            End If
        End If
    Next i
End Sub

Sub addFlightOrderLeftCommentQuickturn()
    Dim wsDataClean As Worksheet
    Dim wsJobSort As Worksheet
    Dim lastRowDataClean As Long
    Dim lastRowJobSort As Long
    Dim i As Long
    Dim j As Long
    Dim matchFound As Boolean
    Dim matchInKeys As Boolean
    Dim key As String
    Dim xValue As Variant
    Dim outputRow As Long
    Dim colKey As Integer

    ' Set the worksheets
    Set wsDataClean = ThisWorkbook.Worksheets("DataClean")
    Set wsJobSort = ThisWorkbook.Worksheets("JobSort")
    
    ' Get the last row in column J of DataClean
    lastRowDataClean = wsDataClean.Cells(wsDataClean.Rows.count, 10).End(xlUp).Row
    ' Get the last row in column K of JobSort
    lastRowJobSort = wsJobSort.Cells(wsJobSort.Rows.count, 11).End(xlUp).Row
    
    ' Loop through each cell in column J of DataClean
    For i = 1 To lastRowDataClean
        ' Check if the cell in column J is not blank
        If wsDataClean.Cells(i, 10).Value <> "" Then
            matchFound = False
            ' Loop through each cell in column K of JobSort
            For j = 1 To lastRowJobSort
                ' Check if the cell in column K is not blank
                If wsJobSort.Cells(j, 11).Value <> "" Then
                    ' Check for a match
                    If wsDataClean.Cells(i, 10).Value = wsJobSort.Cells(j, 11).Value Then
                        matchFound = True
                        Exit For
                    End If
                End If
            Next j
            
            ' If a match is found and column H of DataClean is blank, insert a new line below
            If matchFound And wsDataClean.Cells(i, 8).Value = "" Then ' Check if column H of DataClean is blank
                wsDataClean.Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                ' Add bullet character to column B of the new row
                wsDataClean.Cells(i + 1, 2).Value = "•" ' Column B is 2
                
                ' Get the value from column A of the matched row
                key = wsDataClean.Cells(i, 1).Value ' Key from column A
                
                ' Check for a match in columns L, O, R, etc.
                matchInKeys = False
                colKey = 12 ' Start from column L (12)
                
                Do While wsDataClean.Cells(1, colKey).Value <> ""
                    If wsDataClean.Cells(1, colKey).Value = key Then
                        matchInKeys = True
                        ' Get the value from three columns over (N for L1, T for R1, etc.)
                        xValue = wsDataClean.Cells(4, colKey + 2).Value ' Column N for L1, T for R1, etc.
                        Exit Do
                    End If
                    colKey = colKey + 3 ' Move to the next key column (skip one column)
                Loop
                
                ' If a match was found, add the text to cell D of the new row
                If matchInKeys Then
                    If xValue = 1 Then
                        wsDataClean.Cells(i + 1, 4).Value = "This is the last quickturn order" ' Column D is 4
                    Else
                        wsDataClean.Cells(i + 1, 4).Value = "There are " & xValue & " Quickturn orders left" ' Column D is 4
                    End If
                End If
            End If
        End If
    Next i
End Sub

' ============================================================

' EMAIL ENGINE v2.0

' Purpose:

'   Builds the Tie-In email as clean editable Outlook HTML.

'   No Excel copy/paste.

'   No Windows clipboard dependency.

'   No visible table borders.

'   Forces all report text to uppercase.

' ============================================================

Sub OpenEmailTieIn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim emailHtml As String

    Set ws = ThisWorkbook.Sheets("DataClean")

    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    If ws.Cells(ws.Rows.count, "C").End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row
    If ws.Cells(ws.Rows.count, "D").End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).Row

    If lastRow < 4 Then
        MsgBox "No Tie-In data found to email.", vbExclamation
        Exit Sub
    End If

    emailHtml = TIR_BuildEmailHtml(ws, 4, lastRow)

    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then Set outlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If outlookApp Is Nothing Then
        MsgBox "Unable to start Outlook.", vbCritical
        Exit Sub
    End If

    Set outlookMail = outlookApp.CreateItem(0)

    With outlookMail
        .To = "DL-BDSKC46FieldTiein@exchange.boeing.com"
        .Subject = "BDS Field Tie in"
        .BodyFormat = 2
        .Display
        .HTMLBody = emailHtml & .HTMLBody
    End With

    Set outlookMail = Nothing
    Set outlookApp = Nothing
End Sub


Private Function TIR_BuildEmailHtml(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long) As String
    Dim html As String
    Dim r As Long
    Dim c As Long
    Dim cell As Range
    Dim rowHtml As String
    Dim bText As String
    Dim cText As String
    Dim dText As String

    html = ""
    html = html & "<html><body style='margin:0;padding:0;font-family:Calibri,Arial,sans-serif;font-size:11pt;'>"
    html = html & "<table cellpadding='0' cellspacing='0' style='border-collapse:collapse;border:none;width:980px;font-family:Calibri,Arial,sans-serif;font-size:11pt;'>"
    html = html & "<colgroup>"
    html = html & "<col style='width:120px;'>"
    html = html & "<col style='width:170px;'>"
    html = html & "<col style='width:690px;'>"
    html = html & "</colgroup>"

    For r = firstRow To lastRow
        bText = TIR_CleanCellText(ws.Cells(r, "B").Text)
        cText = TIR_CleanCellText(ws.Cells(r, "C").Text)
        dText = TIR_CleanCellText(ws.Cells(r, "D").Text)

        If TIR_RowIsBlank(bText, cText, dText) Then
            html = html & "<tr><td colspan='3' style='border:none;height:8px;font-size:4pt;line-height:4pt;'>&nbsp;</td></tr>"
        ElseIf TIR_IsAirplaneHeaderRow(ws, r) Then
            html = html & TIR_RenderHeaderRow(ws, r, bText, cText, dText)
        ElseIf TIR_IsBulletRow(bText, cText, dText) Then
            html = html & TIR_RenderBulletRow(ws, r, bText, cText, dText)
        Else
            rowHtml = "<tr>"
            For c = 2 To 4
                Set cell = ws.Cells(r, c)
                rowHtml = rowHtml & TIR_RenderCell(cell, 1, False)
            Next c
            rowHtml = rowHtml & "</tr>"
            html = html & rowHtml
        End If
    Next r

    html = html & "</table></body></html>"
    TIR_BuildEmailHtml = html
End Function


Private Function TIR_RenderHeaderRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal bText As String, ByVal cText As String, ByVal dText As String) As String
    Dim headerText As String
    Dim bgColor As String
    Dim fontColor As String

    headerText = TIR_FirstNonBlankText(bText, cText, dText)
    bgColor = TIR_GetFillColor(ws.Cells(rowNumber, "B"))
    If bgColor = "transparent" Then bgColor = TIR_GetFillColor(ws.Cells(rowNumber, "C"))
    If bgColor = "transparent" Then bgColor = TIR_GetFillColor(ws.Cells(rowNumber, "D"))
    fontColor = TIR_GetCellColor(ws.Cells(rowNumber, "B"))

    TIR_RenderHeaderRow = "<tr><td colspan='3' style='border:none;padding:3px 8px;vertical-align:middle;text-align:center;font-weight:bold;color:" & fontColor & ";background-color:" & bgColor & ";'>" & TIR_HtmlEncode(UCase$(headerText)) & "</td></tr>"
End Function


Private Function TIR_RenderBulletRow(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal bText As String, ByVal cText As String, ByVal dText As String) As String
    Dim bulletText As String
    Dim fontColor As String
    Dim fontWeight As String

    bulletText = TIR_FirstNonBlankText(cText, dText, "")
    If Len(bulletText) = 0 Then bulletText = bText

    fontColor = TIR_GetCellColor(ws.Cells(rowNumber, "C"))
    If fontColor = "#000000" Then fontColor = TIR_GetCellColor(ws.Cells(rowNumber, "D"))

    fontWeight = TIR_GetFontWeight(ws.Cells(rowNumber, "C"))
    If fontWeight = "normal" Then fontWeight = TIR_GetFontWeight(ws.Cells(rowNumber, "D"))

    TIR_RenderBulletRow = "<tr>"
    TIR_RenderBulletRow = TIR_RenderBulletRow & "<td style='border:none;padding:1px 8px;vertical-align:top;'>&nbsp;</td>"
    TIR_RenderBulletRow = TIR_RenderBulletRow & "<td colspan='2' style='border:none;padding:1px 8px 3px 22px;vertical-align:top;color:" & fontColor & ";font-weight:" & fontWeight & ";'>"
    TIR_RenderBulletRow = TIR_RenderBulletRow & "&bull;&nbsp;" & TIR_HtmlEncode(UCase$(bulletText)) & "</td></tr>"
End Function


Private Function TIR_RenderCell(ByVal cell As Range, ByVal colspan As Long, ByVal forceHeader As Boolean) As String
    Dim cellText As String
    Dim styleText As String

    cellText = TIR_CleanCellText(cell.Text)
    If Len(cellText) = 0 Then cellText = "&nbsp;" Else cellText = TIR_HtmlEncode(UCase$(cellText))
    cellText = Replace(cellText, vbCrLf, "<br>")
    cellText = Replace(cellText, vbLf, "<br>")
    cellText = Replace(cellText, vbCr, "<br>")

    styleText = ""
    styleText = styleText & "border:none;"
    styleText = styleText & "padding:3px 8px;"
    styleText = styleText & "vertical-align:top;"
    styleText = styleText & "white-space:normal;"
    styleText = styleText & "font-family:Calibri,Arial,sans-serif;"
    styleText = styleText & "font-size:11pt;"
    styleText = styleText & "color:" & TIR_GetCellColor(cell) & ";"
    styleText = styleText & "font-weight:" & TIR_GetFontWeight(cell) & ";"
    styleText = styleText & "background-color:" & TIR_GetFillColor(cell) & ";"

    Select Case cell.HorizontalAlignment
        Case xlCenter
            styleText = styleText & "text-align:center;"
        Case xlRight
            styleText = styleText & "text-align:right;"
        Case Else
            styleText = styleText & "text-align:left;"
    End Select

    TIR_RenderCell = "<td style='" & styleText & "'>" & cellText & "</td>"
End Function


Private Function TIR_IsAirplaneHeaderRow(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim bFill As String
    Dim cFill As String
    Dim dFill As String
    Dim combinedText As String

    bFill = TIR_GetFillColor(ws.Cells(rowNumber, "B"))
    cFill = TIR_GetFillColor(ws.Cells(rowNumber, "C"))
    dFill = TIR_GetFillColor(ws.Cells(rowNumber, "D"))

    combinedText = UCase$(Trim$(CStr(ws.Cells(rowNumber, "B").Text) & " " & CStr(ws.Cells(rowNumber, "C").Text) & " " & CStr(ws.Cells(rowNumber, "D").Text)))

    If InStr(1, combinedText, "LN", vbTextCompare) > 0 And InStr(1, combinedText, "STALL", vbTextCompare) > 0 Then
        TIR_IsAirplaneHeaderRow = True
        Exit Function
    End If

    If bFill <> "transparent" And bFill = cFill And cFill = dFill Then
        If bFill <> "#FFFFFF" Then
            TIR_IsAirplaneHeaderRow = True
            Exit Function
        End If
    End If

    TIR_IsAirplaneHeaderRow = False
End Function


Private Function TIR_IsBulletRow(ByVal bText As String, ByVal cText As String, ByVal dText As String) As Boolean
    Dim leftText As String

    leftText = Trim$(bText)

    If leftText = "?" Then
        TIR_IsBulletRow = True
        Exit Function
    End If

    If leftText = ChrW(8226) Then
        TIR_IsBulletRow = True
        Exit Function
    End If

    If leftText = "ï¿½" Or leftText = "Ï¿½" Or leftText = "?" Then
        If Len(Trim$(cText & dText)) > 0 Then
            TIR_IsBulletRow = True
            Exit Function
        End If
    End If

    TIR_IsBulletRow = False
End Function


Private Function TIR_CleanCellText(ByVal rawText As String) As String
    rawText = Replace(rawText, ChrW(160), " ")
    rawText = Replace(rawText, "ï¿½", "?")
    rawText = Replace(rawText, "Ï¿½", "?")
    TIR_CleanCellText = Trim$(rawText)
End Function


Private Function TIR_RowIsBlank(ByVal bText As String, ByVal cText As String, ByVal dText As String) As Boolean
    If Len(Trim$(bText)) = 0 And Len(Trim$(cText)) = 0 And Len(Trim$(dText)) = 0 Then
        TIR_RowIsBlank = True
    Else
        TIR_RowIsBlank = False
    End If
End Function


Private Function TIR_FirstNonBlankText(ByVal firstText As String, ByVal secondText As String, ByVal thirdText As String) As String
    If Len(Trim$(firstText)) > 0 Then
        TIR_FirstNonBlankText = Trim$(firstText)
        Exit Function
    End If

    If Len(Trim$(secondText)) > 0 Then
        TIR_FirstNonBlankText = Trim$(secondText)
        Exit Function
    End If

    TIR_FirstNonBlankText = Trim$(thirdText)
End Function


Private Function TIR_GetFontWeight(ByVal cell As Range) As String
    If cell.Font.Bold Then
        TIR_GetFontWeight = "bold"
    Else
        TIR_GetFontWeight = "normal"
    End If
End Function


Private Function TIR_GetCellColor(ByVal cell As Range) As String
    If cell.Font.ColorIndex = xlColorIndexAutomatic Then
        TIR_GetCellColor = "#000000"
    Else
        TIR_GetCellColor = TIR_VbaColorToHtml(cell.Font.Color)
    End If
End Function


Private Function TIR_GetFillColor(ByVal cell As Range) As String
    If cell.Interior.ColorIndex = xlColorIndexNone Then
        TIR_GetFillColor = "transparent"
    Else
        TIR_GetFillColor = TIR_VbaColorToHtml(cell.Interior.Color)
    End If
End Function


Private Function TIR_VbaColorToHtml(ByVal colorValue As Long) As String
    Dim r As Long
    Dim g As Long
    Dim b As Long
    Dim rHex As String
    Dim gHex As String
    Dim bHex As String

    r = colorValue Mod 256
    g = (colorValue \ 256) Mod 256
    b = (colorValue \ 65536) Mod 256

    rHex = Right$("0" & Hex$(r), 2)
    gHex = Right$("0" & Hex$(g), 2)
    bHex = Right$("0" & Hex$(b), 2)

    TIR_VbaColorToHtml = "#" & rHex & gHex & bHex
End Function


Private Function TIR_HtmlEncode(ByVal textValue As String) As String
    textValue = Replace(textValue, "&", "&amp;")
    textValue = Replace(textValue, "<", "&lt;")
    textValue = Replace(textValue, ">", "&gt;")
    textValue = Replace(textValue, """", "&quot;")
    TIR_HtmlEncode = textValue
End Function




 Sub openWordTieIn()
    Dim dataCleanSheet As Worksheet
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim lastRow As Long

    ' Set reference to the DataClean sheet
    Set dataCleanSheet = ThisWorkbook.Sheets("DataClean")

    ' Find the last row in column B of DataClean
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "B").End(xlUp).Row

    ' Create a new instance of Word
    On Error Resume Next
    Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0

    ' Make Word visible
    wordApp.Visible = True

    ' Create a new document
    Set wordDoc = wordApp.Documents.Add

    ' Copy data from columns B to D starting from row 9
    dataCleanSheet.Range("B9:D" & lastRow).Copy

    ' Paste data into the Word document
    wordDoc.Content.Paste

    ' Clear the clipboard to free up memory
    Application.CutCopyMode = False

End Sub

Sub CombineWOAndJobDescription()
    Dim dataCleanSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheet
    Set dataCleanSheet = ThisWorkbook.Sheets("FileExport")
    
    ' Find the last row in column B
    lastRow = dataCleanSheet.Cells(dataCleanSheet.Rows.count, "B").End(xlUp).Row
    
    ' Loop through each row starting from row 2
    For i = 2 To lastRow
        ' Concatenate values from column B and D and place in column I
        dataCleanSheet.Cells(i, "I").Value = dataCleanSheet.Cells(i, "B").Value & ":" & " " & dataCleanSheet.Cells(i, "D").Value
    Next i
End Sub

Sub SeparateNames()
    Dim dataProcessSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim dataPoints As Variant

    ' Set the worksheet
    Set dataProcessSheet = ThisWorkbook.Sheets("FileExport")
    
    ' Find the last row in column H
    lastRow = dataProcessSheet.Cells(dataProcessSheet.Rows.count, "H").End(xlUp).Row
    
    ' Loop through each row in column H
    For i = 1 To lastRow
        ' Split the data points by comma
        dataPoints = Split(dataProcessSheet.Cells(i, "H").Value, ",")
        
        ' Loop through each data point and place it in the corresponding cell in column K and beyond
        For j = LBound(dataPoints) To UBound(dataPoints)
            dataProcessSheet.Cells(i, "K").Offset(0, j).Value = Trim(dataPoints(j)) ' Trim to remove any leading/trailing spaces
        Next j
    Next i
End Sub

Sub FillFirstDar()
    Dim wsFileExport As Worksheet
    Dim wsFirstDar As Worksheet
    Dim lastRowFileExport As Long
    Dim lastRowFirstDar As Long
    Dim i As Long, j As Long, Col As Long
    Dim keyValueA As Variant

    ' Set the worksheets
    Set wsFileExport = ThisWorkbook.Sheets("FileExport") ' List 1
    Set wsFirstDar = ThisWorkbook.Sheets("FirstDar") ' List 2

    ' Find the last rows in both sheets
    lastRowFileExport = wsFileExport.Cells(wsFileExport.Rows.count, "K").End(xlUp).Row
    lastRowFirstDar = wsFirstDar.Cells(wsFirstDar.Rows.count, "C").End(xlUp).Row

    ' Loop through each column from K (11) to AD (30) in List 1
    For Col = 11 To 30 ' Columns K to AD
        ' Loop through each key in List 1 (FileExport)
        For i = 2 To lastRowFileExport ' Assuming headers in row 1
            keyValueA = wsFileExport.Cells(i, Col).Value ' Value A from List 1 (Current Column)

            ' Loop through each key in List 2 (FirstDar)
            For j = 2 To lastRowFirstDar ' Assuming headers in row 1
                ' Check if the keys match
                If keyValueA = wsFirstDar.Cells(j, "C").Value Then ' Value A from List 2 (Column C)
                    ' Check if columns F and G are empty
                    If IsEmpty(wsFirstDar.Cells(j, "F").Value) And IsEmpty(wsFirstDar.Cells(j, "G").Value) Then
                        ' Copy Value B and C from List 1 to F and G in List 2
                        wsFirstDar.Cells(j, "F").Value = wsFileExport.Cells(i, "A").Value ' Value B to Column F
                        wsFirstDar.Cells(j, "G").Value = wsFileExport.Cells(i, "I").Value ' Value C to Column G
                    ' Check if columns M and N are empty
                    ElseIf IsEmpty(wsFirstDar.Cells(j, "M").Value) And IsEmpty(wsFirstDar.Cells(j, "N").Value) Then
                        ' Copy Value B and C from List 1 to M and N in List 2
                        wsFirstDar.Cells(j, "M").Value = wsFileExport.Cells(i, "A").Value ' Value B to Column M
                        wsFirstDar.Cells(j, "N").Value = wsFileExport.Cells(i, "I").Value ' Value C to Column N
                    ' Check if columns T and V are empty
                    ElseIf IsEmpty(wsFirstDar.Cells(j, "T").Value) And IsEmpty(wsFirstDar.Cells(j, "V").Value) Then
                        ' Copy Value B and C from List 1 to T and V in List 2
                        wsFirstDar.Cells(j, "T").Value = wsFileExport.Cells(i, "A").Value ' Value B to Column T
                        wsFirstDar.Cells(j, "V").Value = wsFileExport.Cells(i, "I").Value ' Value C to Column V
                    End If
                    Exit For ' Exit the inner loop once a match is found
                End If
            Next j
        Next i
    Next Col

End Sub

Sub FillSecondDar()
    Dim wsFileExport As Worksheet
    Dim wsSecondDar As Worksheet
    Dim lastRowFileExport As Long
    Dim lastRowSecondDar As Long
    Dim i As Long, j As Long, Col As Long
    Dim keyValueA As Variant

    ' Set the worksheets
    Set wsFileExport = ThisWorkbook.Sheets("FileExport") ' List 1
    Set wsSecondDar = ThisWorkbook.Sheets("SecondDar") ' List 2

    ' Find the last rows in both sheets
    lastRowFileExport = wsFileExport.Cells(wsFileExport.Rows.count, "K").End(xlUp).Row
    lastRowSecondDar = wsSecondDar.Cells(wsSecondDar.Rows.count, "C").End(xlUp).Row

    ' Loop through each column from K (11) to AA (27) in List 1
    For Col = 11 To 27 ' Columns K to AA
        ' Loop through each key in List 1 (FileExport)
        For i = 2 To lastRowFileExport ' Assuming headers in row 1
            keyValueA = wsFileExport.Cells(i, Col).Value ' Value A from List 1 (Current Column)

            ' Only proceed if keyValueA is not blank
            If keyValueA <> "" Then
                ' Loop through each key in List 2 (SecondDar)
                For j = 2 To lastRowSecondDar ' Assuming headers in row 1
                    ' Check if the keys match
                    If keyValueA = wsSecondDar.Cells(j, "C").Value Then ' Value A from List 2 (Column C)
                        ' Check if columns F and G are empty
                        If IsEmpty(wsSecondDar.Cells(j, "F").Value) And IsEmpty(wsSecondDar.Cells(j, "G").Value) Then
                            ' Copy Value B and C from List 1 to F and G in List 2
                            wsSecondDar.Cells(j, "F").Value = wsFileExport.Cells(i, "A").Value ' Value B to Column F
                            wsSecondDar.Cells(j, "G").Value = wsFileExport.Cells(i, "I").Value ' Value C to Column G
                        ' Check if columns M and N are empty
                        ElseIf IsEmpty(wsSecondDar.Cells(j, "M").Value) And IsEmpty(wsSecondDar.Cells(j, "N").Value) Then
                            ' Copy Value B and C from List 1 to M and N in List 2
                            wsSecondDar.Cells(j, "M").Value = wsFileExport.Cells(i, "A").Value ' Value B to Column M
                            wsSecondDar.Cells(j, "N").Value = wsFileExport.Cells(i, "I").Value ' Value C to Column N
                        End If
                        Exit For ' Exit the inner loop once a match is found
                    End If
                Next j
            End If
        Next i
    Next Col
    
End Sub

Sub FillThirdDar()
    Dim wsFileExport As Worksheet
    Dim wsThirdDar As Worksheet
    Dim lastRowFileExport As Long
    Dim lastRowThirdDar As Long
    Dim i As Long, j As Long, Col As Long
    Dim keyValueA As Variant

    ' Set the worksheets
    Set wsFileExport = ThisWorkbook.Sheets("FileExport") ' List 1
    Set wsThirdDar = ThisWorkbook.Sheets("ThirdDar") ' List 2

    ' Find the last rows in both sheets
    lastRowFileExport = wsFileExport.Cells(wsFileExport.Rows.count, "K").End(xlUp).Row
    lastRowThirdDar = wsThirdDar.Cells(wsThirdDar.Rows.count, "C").End(xlUp).Row

    ' Loop through each column from K (11) to AA (27) in List 1
    For Col = 11 To 27 ' Columns K to AA
        ' Loop through each key in List 1 (FileExport)
        For i = 2 To lastRowFileExport ' Assuming headers in row 1
            keyValueA = wsFileExport.Cells(i, Col).Value ' Value A from List 1 (Current Column)

            ' Only proceed if keyValueA is not blank
            If keyValueA <> "" Then
                ' Loop through each key in List 2 (ThirdDar)
                For j = 2 To lastRowThirdDar ' Assuming headers in row 1
                    ' Check if the keys match
                    If keyValueA = wsThirdDar.Cells(j, "C").Value Then ' Value A from List 2 (Column C)
                        ' Check if columns F and G are empty
                        If IsEmpty(wsThirdDar.Cells(j, "F").Value) And IsEmpty(wsThirdDar.Cells(j, "G").Value) Then
                            ' Copy Value B and C from List 1 to F and G in List 2
                            wsThirdDar.Cells(j, "F").Value = wsFileExport.Cells(i, "A").Value ' Value B to Column F
                            wsThirdDar.Cells(j, "G").Value = wsFileExport.Cells(i, "I").Value ' Value C to Column G
                        ' Check if columns M and N are empty
                        ElseIf IsEmpty(wsThirdDar.Cells(j, "M").Value) And IsEmpty(wsThirdDar.Cells(j, "N").Value) Then
                            ' Copy Value B and C from List 1 to M and N in List 2
                            wsThirdDar.Cells(j, "M").Value = wsFileExport.Cells(i, "A").Value ' Value B to Column M
                            wsThirdDar.Cells(j, "N").Value = wsFileExport.Cells(i, "I").Value ' Value C to Column N
                        End If
                        Exit For ' Exit the inner loop once a match is found
                    End If
                Next j
            End If
        Next i
    Next Col
    
End Sub

Sub FillFirstDarPriorityBoard()
    Dim wsFileExport As Worksheet
    Dim wsFirstDar As Worksheet
    Dim lastRowFileExport As Long
    Dim lastRowFirstDar As Long
    Dim i As Long, j As Long
    Dim pasteRow As Long
    Dim keyValueK As Variant

    ' Set the worksheets
    Set wsFileExport = ThisWorkbook.Sheets("FileExport")
    Set wsFirstDar = ThisWorkbook.Sheets("FirstDar")

    ' Find the last row in FileExport based on column K
    lastRowFileExport = wsFileExport.Cells(wsFileExport.Rows.count, "K").End(xlUp).Row
    ' Find the last row in FirstDar based on column C
    lastRowFirstDar = wsFirstDar.Cells(wsFirstDar.Rows.count, "C").End(xlUp).Row

    ' Start pasting in FirstDar from row 24
    pasteRow = 24

    ' Loop through each row in column K starting from row 2
    For i = 2 To lastRowFileExport
        ' Get the value from column K
        keyValueK = wsFileExport.Cells(i, "K").Value
        
        ' Check if the cell in column K is not blank
        If keyValueK <> "" Then
            ' Loop through each row in column C of FirstDar to find a match
            For j = 2 To lastRowFirstDar
                ' Check if the value in column C of FirstDar matches the value in column K
                If wsFirstDar.Cells(j, "C").Value = keyValueK Then
                    ' Copy values from FileExport and paste them into FirstDar
                    wsFirstDar.Cells(pasteRow, "C").Value = wsFileExport.Cells(i, "A").Value ' Column A to Column C
                    wsFirstDar.Cells(pasteRow, "D").Value = wsFileExport.Cells(i, "D").Value ' Column D to Column D
                    wsFirstDar.Cells(pasteRow, "J").Value = wsFileExport.Cells(i, "G").Value ' Column G to Column J
                    wsFirstDar.Cells(pasteRow, "R").Value = wsFileExport.Cells(i, "F").Value ' Column F to Column R
                    wsFirstDar.Cells(pasteRow, "S").Value = wsFileExport.Cells(i, "E").Value ' Column E to Column S
                    
                    ' Move to the next row in FirstDar for pasting
                    pasteRow = pasteRow + 1
                    Exit For ' Exit the inner loop once a match is found
                End If
            Next j
        End If
    Next i

    ' Copy values from column J to column AG for rows 24 to 38
    Dim k As Long
    For k = 24 To 38
        wsFirstDar.Cells(k, "AG").Value = wsFirstDar.Cells(k, "J").Value
    Next k
    
    ' Set column AG to wrap text
    wsFirstDar.Columns("AG").WrapText = True

End Sub

Sub FillSecondDarPriorityBoard()
    Dim wsFileExport As Worksheet
    Dim wsSecondDar As Worksheet
    Dim lastRowFileExport As Long
    Dim lastRowSecondDar As Long
    Dim i As Long, j As Long
    Dim pasteRow As Long
    Dim keyValueK As Variant

    ' Set the worksheets
    Set wsFileExport = ThisWorkbook.Sheets("FileExport")
    Set wsSecondDar = ThisWorkbook.Sheets("SecondDar")

    ' Find the last row in FileExport based on column K
    lastRowFileExport = wsFileExport.Cells(wsFileExport.Rows.count, "K").End(xlUp).Row
    ' Find the last row in SecondDar based on column C
    lastRowSecondDar = wsSecondDar.Cells(wsSecondDar.Rows.count, "C").End(xlUp).Row

    ' Start pasting in SecondDar from row 26
    pasteRow = 26

    ' Loop through each row in column K starting from row 2
    For i = 2 To lastRowFileExport
        ' Get the value from column K
        keyValueK = wsFileExport.Cells(i, "K").Value
        
        ' Check if the cell in column K is not blank
        If keyValueK <> "" Then
            ' Loop through each row in column C of SecondDar to find a match
            For j = 2 To lastRowSecondDar
                ' Check if the value in column C of SecondDar matches the value in column K
                If wsSecondDar.Cells(j, "C").Value = keyValueK Then
                    ' Copy values from FileExport and paste them into SecondDar
                    wsSecondDar.Cells(pasteRow, "C").Value = wsFileExport.Cells(i, "A").Value ' Column A to Column C
                    wsSecondDar.Cells(pasteRow, "D").Value = wsFileExport.Cells(i, "D").Value ' Column D to Column D
                    wsSecondDar.Cells(pasteRow, "J").Value = wsFileExport.Cells(i, "G").Value ' Column G to Column J
                    wsSecondDar.Cells(pasteRow, "R").Value = wsFileExport.Cells(i, "F").Value ' Column F to Column R
                    wsSecondDar.Cells(pasteRow, "S").Value = wsFileExport.Cells(i, "E").Value ' Column E to Column S
                    
                    ' Move to the next row in SecondDar for pasting
                    pasteRow = pasteRow + 1
                    Exit For ' Exit the inner loop once a match is found
                End If
            Next j
        End If
    Next i

    ' Copy values from column J to column AG for rows 26 to 38
    Dim k As Long
    For k = 26 To 40
        wsSecondDar.Cells(k, "AG").Value = wsSecondDar.Cells(k, "J").Value
    Next k
    
    ' Set column AG to wrap text
    wsSecondDar.Columns("AG").WrapText = True

End Sub

Sub FillThirdDarPriorityBoard()
    Dim wsFileExport As Worksheet
    Dim wsThirdDar As Worksheet
    Dim lastRowFileExport As Long
    Dim lastRowThirdDar As Long
    Dim i As Long, j As Long
    Dim pasteRow As Long
    Dim keyValueK As Variant

    ' Set the worksheets
    Set wsFileExport = ThisWorkbook.Sheets("FileExport")
    Set wsThirdDar = ThisWorkbook.Sheets("ThirdDar")

    ' Find the last row in FileExport based on column K
    lastRowFileExport = wsFileExport.Cells(wsFileExport.Rows.count, "K").End(xlUp).Row
    ' Find the last row in ThirdDar based on column C
    lastRowThirdDar = wsThirdDar.Cells(wsThirdDar.Rows.count, "C").End(xlUp).Row

    ' Start pasting in ThirdDar from row 17
    pasteRow = 17

    ' Loop through each row in column K starting from row 2
    For i = 2 To lastRowFileExport
        ' Get the value from column K
        keyValueK = wsFileExport.Cells(i, "K").Value
        
        ' Check if the cell in column K is not blank
        If keyValueK <> "" Then
            ' Loop through each row in column C of ThirdDar to find a match
            For j = 2 To lastRowThirdDar
                ' Check if the value in column C of ThirdDar matches the value in column K
                If wsThirdDar.Cells(j, "C").Value = keyValueK Then
                    ' Copy values from FileExport and paste them into ThirdDar
                    wsThirdDar.Cells(pasteRow, "C").Value = wsFileExport.Cells(i, "A").Value ' Column A to Column C
                    wsThirdDar.Cells(pasteRow, "D").Value = wsFileExport.Cells(i, "D").Value ' Column D to Column D
                    wsThirdDar.Cells(pasteRow, "J").Value = wsFileExport.Cells(i, "G").Value ' Column G to Column J
                    wsThirdDar.Cells(pasteRow, "R").Value = wsFileExport.Cells(i, "F").Value ' Column F to Column R
                    wsThirdDar.Cells(pasteRow, "S").Value = wsFileExport.Cells(i, "E").Value ' Column E to Column S
                    
                    ' Move to the next row in ThirdDar for pasting
                    pasteRow = pasteRow + 1
                    Exit For ' Exit the inner loop once a match is found
                End If
            Next j
        End If
    Next i

    ' Copy values from column J to column AG for rows 17 to 31
    Dim k As Long
    For k = 17 To 31
        wsThirdDar.Cells(k, "AG").Value = wsThirdDar.Cells(k, "J").Value
    Next k
    
    ' Set column AG to wrap text
    wsThirdDar.Columns("AG").WrapText = True

End Sub

Sub NicknameRemove()
    Dim wsSecondDar As Worksheet
    Dim lastRowSecondDar As Long
    Dim i As Long

    ' Set the worksheet
    Set wsSecondDar = ThisWorkbook.Sheets("SecondDar")

    ' Find the last row in column C
    lastRowSecondDar = wsSecondDar.Cells(wsSecondDar.Rows.count, "C").End(xlUp).Row

    ' Loop through each cell in column C
    For i = 2 To lastRowSecondDar ' Assuming headers in row 1
        ' Check if the cell value is "DC Forbes"
        If wsSecondDar.Cells(i, "C").Value = "DC Forbes" Then
            ' Replace with "Dean C Forbes"
            wsSecondDar.Cells(i, "C").Value = "Dean C Forbes"
        End If
    Next i

End Sub

Sub NicknameAdd()
    Dim wsSecondDar As Worksheet
    Dim lastRowSecondDar As Long
    Dim i As Long

    ' Set the worksheet
    Set wsSecondDar = ThisWorkbook.Sheets("SecondDar")

    ' Find the last row in column C
    lastRowSecondDar = wsSecondDar.Cells(wsSecondDar.Rows.count, "C").End(xlUp).Row

    ' Loop through each cell in column C
    For i = 2 To lastRowSecondDar ' Assuming headers in row 1
        ' Check if the cell value is "Dean C Forbes"
        If wsSecondDar.Cells(i, "C").Value = "Dean C Forbes" Then
            ' Replace with "DC Forbes"
            wsSecondDar.Cells(i, "C").Value = "DC Forbes"
        End If
    Next i

End Sub

Sub zLineFiveButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F and G in temporary variables
    tempF = ActiveSheet.Range("F5").Value
    tempG = ActiveSheet.Range("G5").Value
    
    ' Swap the values: Move M and N to F and G
    ActiveSheet.Range("F5").Value = ActiveSheet.Range("M5").Value
    ActiveSheet.Range("G5").Value = ActiveSheet.Range("N5").Value
    
    ' Move the temporary values to M and N
    ActiveSheet.Range("M5").Value = tempF
    ActiveSheet.Range("N5").Value = tempG
End Sub

Sub zLineFiveButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M5 and N5 in temporary variables
    tempM = ActiveSheet.Range("M5").Value
    tempN = ActiveSheet.Range("N5").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M5").Value = ActiveSheet.Range("T5").Value
    ActiveSheet.Range("N5").Value = ActiveSheet.Range("V5").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T5").Value = tempM
    ActiveSheet.Range("V5").Value = tempN
End Sub

Sub zLineSixButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F6 and G6 in temporary variables
    tempF = ActiveSheet.Range("F6").Value
    tempG = ActiveSheet.Range("G6").Value
    
    ' Swap the values: Move M6 and N6 to F6 and G6
    ActiveSheet.Range("F6").Value = ActiveSheet.Range("M6").Value
    ActiveSheet.Range("G6").Value = ActiveSheet.Range("N6").Value
    
    ' Move the temporary values to M6 and N6
    ActiveSheet.Range("M6").Value = tempF
    ActiveSheet.Range("N6").Value = tempG
End Sub

Sub zLineSixButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M6 and N6 in temporary variables
    tempM = ActiveSheet.Range("M6").Value
    tempN = ActiveSheet.Range("N6").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M6").Value = ActiveSheet.Range("T6").Value
    ActiveSheet.Range("N6").Value = ActiveSheet.Range("V6").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T6").Value = tempM
    ActiveSheet.Range("V6").Value = tempN
End Sub

Sub zLineSevenButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F7 and G7 in temporary variables
    tempF = ActiveSheet.Range("F7").Value
    tempG = ActiveSheet.Range("G7").Value
    
    ' Swap the values: Move M7 and N7 to F7 and G7
    ActiveSheet.Range("F7").Value = ActiveSheet.Range("M7").Value
    ActiveSheet.Range("G7").Value = ActiveSheet.Range("N7").Value
    
    ' Move the temporary values to M7 and N7
    ActiveSheet.Range("M7").Value = tempF
    ActiveSheet.Range("N7").Value = tempG
End Sub

Sub zLineEightButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F8 and G8 in temporary variables
    tempF = ActiveSheet.Range("F8").Value
    tempG = ActiveSheet.Range("G8").Value
    
    ' Swap the values: Move M8 and N8 to F8 and G8
    ActiveSheet.Range("F8").Value = ActiveSheet.Range("M8").Value
    ActiveSheet.Range("G8").Value = ActiveSheet.Range("N8").Value
    
    ' Move the temporary values to M8 and N8
    ActiveSheet.Range("M8").Value = tempF
    ActiveSheet.Range("N8").Value = tempG
End Sub

Sub zLineEightButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M8 and N8 in temporary variables
    tempM = ActiveSheet.Range("M8").Value
    tempN = ActiveSheet.Range("N8").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M8").Value = ActiveSheet.Range("T8").Value
    ActiveSheet.Range("N8").Value = ActiveSheet.Range("V8").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T8").Value = tempM
    ActiveSheet.Range("V8").Value = tempN
End Sub

Sub zLineNineButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F9 and G9 in temporary variables
    tempF = ActiveSheet.Range("F9").Value
    tempG = ActiveSheet.Range("G9").Value
    
    ' Swap the values: Move M9 and N9 to F9 and G9
    ActiveSheet.Range("F9").Value = ActiveSheet.Range("M9").Value
    ActiveSheet.Range("G9").Value = ActiveSheet.Range("N9").Value
    
    ' Move the temporary values to M9 and N9
    ActiveSheet.Range("M9").Value = tempF
    ActiveSheet.Range("N9").Value = tempG
End Sub

Sub zLineNineButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M9 and N9 in temporary variables
    tempM = ActiveSheet.Range("M9").Value
    tempN = ActiveSheet.Range("N9").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M9").Value = ActiveSheet.Range("T9").Value
    ActiveSheet.Range("N9").Value = ActiveSheet.Range("V9").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T9").Value = tempM
    ActiveSheet.Range("V9").Value = tempN
End Sub

Sub zLineTenButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F10 and G10 in temporary variables
    tempF = ActiveSheet.Range("F10").Value
    tempG = ActiveSheet.Range("G10").Value
    
    ' Swap the values: Move M10 and N10 to F10 and G10
    ActiveSheet.Range("F10").Value = ActiveSheet.Range("M10").Value
    ActiveSheet.Range("G10").Value = ActiveSheet.Range("N10").Value
    
    ' Move the temporary values to M10 and N10
    ActiveSheet.Range("M10").Value = tempF
    ActiveSheet.Range("N10").Value = tempG
End Sub

Sub zLineTenButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M10 and N10 in temporary variables
    tempM = ActiveSheet.Range("M10").Value
    tempN = ActiveSheet.Range("N10").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M10").Value = ActiveSheet.Range("T10").Value
    ActiveSheet.Range("N10").Value = ActiveSheet.Range("V10").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T10").Value = tempM
    ActiveSheet.Range("V10").Value = tempN
End Sub

Sub zLineElevenButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F11 and G11 in temporary variables
    tempF = ActiveSheet.Range("F11").Value
    tempG = ActiveSheet.Range("G11").Value
    
    ' Swap the values: Move M11 and N11 to F11 and G11
    ActiveSheet.Range("F11").Value = ActiveSheet.Range("M11").Value
    ActiveSheet.Range("G11").Value = ActiveSheet.Range("N11").Value
    
    ' Move the temporary values to M11 and N11
    ActiveSheet.Range("M11").Value = tempF
    ActiveSheet.Range("N11").Value = tempG
End Sub

Sub zLineElevenButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M11 and N11 in temporary variables
    tempM = ActiveSheet.Range("M11").Value
    tempN = ActiveSheet.Range("N11").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M11").Value = ActiveSheet.Range("T11").Value
    ActiveSheet.Range("N11").Value = ActiveSheet.Range("V11").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T11").Value = tempM
    ActiveSheet.Range("V11").Value = tempN
End Sub

Sub zLineTwelveButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F12 and G12 in temporary variables
    tempF = ActiveSheet.Range("F12").Value
    tempG = ActiveSheet.Range("G12").Value
    
    ' Swap the values: Move M12 and N12 to F12 and G12
    ActiveSheet.Range("F12").Value = ActiveSheet.Range("M12").Value
    ActiveSheet.Range("G12").Value = ActiveSheet.Range("N12").Value
    
    ' Move the temporary values to M12 and N12
    ActiveSheet.Range("M12").Value = tempF
    ActiveSheet.Range("N12").Value = tempG
End Sub

Sub zLineTwelveButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M12 and N12 in temporary variables
    tempM = ActiveSheet.Range("M12").Value
    tempN = ActiveSheet.Range("N12").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M12").Value = ActiveSheet.Range("T12").Value
    ActiveSheet.Range("N12").Value = ActiveSheet.Range("V12").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T12").Value = tempM
    ActiveSheet.Range("V12").Value = tempN
End Sub

Sub zLineThirteenButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F13 and G13 in temporary variables
    tempF = ActiveSheet.Range("F13").Value
    tempG = ActiveSheet.Range("G13").Value
    
    ' Swap the values: Move M13 and N13 to F13 and G13
    ActiveSheet.Range("F13").Value = ActiveSheet.Range("M13").Value
    ActiveSheet.Range("G13").Value = ActiveSheet.Range("N13").Value
    
    ' Move the temporary values to M13 and N13
    ActiveSheet.Range("M13").Value = tempF
    ActiveSheet.Range("N13").Value = tempG
End Sub

Sub zLineThirteenButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M13 and N13 in temporary variables
    tempM = ActiveSheet.Range("M13").Value
    tempN = ActiveSheet.Range("N13").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M13").Value = ActiveSheet.Range("T13").Value
    ActiveSheet.Range("N13").Value = ActiveSheet.Range("V13").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T13").Value = tempM
    ActiveSheet.Range("V13").Value = tempN
End Sub

Sub zLineFourteenButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F14 and G14 in temporary variables
    tempF = ActiveSheet.Range("F14").Value
    tempG = ActiveSheet.Range("G14").Value
    
    ' Swap the values: Move M14 and N14 to F14 and G14
    ActiveSheet.Range("F14").Value = ActiveSheet.Range("M14").Value
    ActiveSheet.Range("G14").Value = ActiveSheet.Range("N14").Value
    
    ' Move the temporary values to M14 and N14
    ActiveSheet.Range("M14").Value = tempF
    ActiveSheet.Range("N14").Value = tempG
End Sub

Sub zLineFourteenButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M14 and N14 in temporary variables
    tempM = ActiveSheet.Range("M14").Value
    tempN = ActiveSheet.Range("N14").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M14").Value = ActiveSheet.Range("T14").Value
    ActiveSheet.Range("N14").Value = ActiveSheet.Range("V14").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T14").Value = tempM
    ActiveSheet.Range("V14").Value = tempN
End Sub

Sub zLineFifteenButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F15 and G15 in temporary variables
    tempF = ActiveSheet.Range("F15").Value
    tempG = ActiveSheet.Range("G15").Value
    
    ' Swap the values: Move M15 and N15 to F15 and G15
    ActiveSheet.Range("F15").Value = ActiveSheet.Range("M15").Value
    ActiveSheet.Range("G15").Value = ActiveSheet.Range("N15").Value
    
    ' Move the temporary values to M15 and N15
    ActiveSheet.Range("M15").Value = tempF
    ActiveSheet.Range("N15").Value = tempG
End Sub

Sub zLineFifteenButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M15 and N15 in temporary variables
    tempM = ActiveSheet.Range("M15").Value
    tempN = ActiveSheet.Range("N15").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M15").Value = ActiveSheet.Range("T15").Value
    ActiveSheet.Range("N15").Value = ActiveSheet.Range("V15").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T15").Value = tempM
    ActiveSheet.Range("V15").Value = tempN
End Sub

Sub zLineSixteenButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F16 and G16 in temporary variables
    tempF = ActiveSheet.Range("F16").Value
    tempG = ActiveSheet.Range("G16").Value
    
    ' Swap the values: Move M16 and N16 to F16 and G16
    ActiveSheet.Range("F16").Value = ActiveSheet.Range("M16").Value
    ActiveSheet.Range("G16").Value = ActiveSheet.Range("N16").Value
    
    ' Move the temporary values to M16 and N16
    ActiveSheet.Range("M16").Value = tempF
    ActiveSheet.Range("N16").Value = tempG
End Sub

Sub zLineSixteenButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M16 and N16 in temporary variables
    tempM = ActiveSheet.Range("M16").Value
    tempN = ActiveSheet.Range("N16").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M16").Value = ActiveSheet.Range("T16").Value
    ActiveSheet.Range("N16").Value = ActiveSheet.Range("V16").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T16").Value = tempM
    ActiveSheet.Range("V16").Value = tempN
End Sub

Sub zLineSeventeenButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F17 and G17 in temporary variables
    tempF = ActiveSheet.Range("F17").Value
    tempG = ActiveSheet.Range("G17").Value
    
    ' Swap the values: Move M17 and N17 to F17 and G17
    ActiveSheet.Range("F17").Value = ActiveSheet.Range("M17").Value
    ActiveSheet.Range("G17").Value = ActiveSheet.Range("N17").Value
    
    ' Move the temporary values to M17 and N17
    ActiveSheet.Range("M17").Value = tempF
    ActiveSheet.Range("N17").Value = tempG
End Sub

Sub zLineSeventeenButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M17 and N17 in temporary variables
    tempM = ActiveSheet.Range("M17").Value
    tempN = ActiveSheet.Range("N17").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M17").Value = ActiveSheet.Range("T17").Value
    ActiveSheet.Range("N17").Value = ActiveSheet.Range("V17").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T17").Value = tempM
    ActiveSheet.Range("V17").Value = tempN
End Sub

Sub zLineEighteenButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F18 and G18 in temporary variables
    tempF = ActiveSheet.Range("F18").Value
    tempG = ActiveSheet.Range("G18").Value
    
    ' Swap the values: Move M18 and N18 to F18 and G18
    ActiveSheet.Range("F18").Value = ActiveSheet.Range("M18").Value
    ActiveSheet.Range("G18").Value = ActiveSheet.Range("N18").Value
    
    ' Move the temporary values to M18 and N18
    ActiveSheet.Range("M18").Value = tempF
    ActiveSheet.Range("N18").Value = tempG
End Sub

Sub zLineEighteenButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M18 and N18 in temporary variables
    tempM = ActiveSheet.Range("M18").Value
    tempN = ActiveSheet.Range("N18").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M18").Value = ActiveSheet.Range("T18").Value
    ActiveSheet.Range("N18").Value = ActiveSheet.Range("V18").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T18").Value = tempM
    ActiveSheet.Range("V18").Value = tempN
End Sub

Sub zLineNineteenButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F19 and G19 in temporary variables
    tempF = ActiveSheet.Range("F19").Value
    tempG = ActiveSheet.Range("G19").Value
    
    ' Swap the values: Move M19 and N19 to F19 and G19
    ActiveSheet.Range("F19").Value = ActiveSheet.Range("M19").Value
    ActiveSheet.Range("G19").Value = ActiveSheet.Range("N19").Value
    
    ' Move the temporary values to M19 and N19
    ActiveSheet.Range("M19").Value = tempF
    ActiveSheet.Range("N19").Value = tempG
End Sub

Sub zLineNineteenButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M19 and N19 in temporary variables
    tempM = ActiveSheet.Range("M19").Value
    tempN = ActiveSheet.Range("N19").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M19").Value = ActiveSheet.Range("T19").Value
    ActiveSheet.Range("N19").Value = ActiveSheet.Range("V19").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T19").Value = tempM
    ActiveSheet.Range("V19").Value = tempN
End Sub

Sub zLineTwentyButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F20 and G20 in temporary variables
    tempF = ActiveSheet.Range("F20").Value
    tempG = ActiveSheet.Range("G20").Value
    
    ' Swap the values: Move M20 and N20 to F20 and G20
    ActiveSheet.Range("F20").Value = ActiveSheet.Range("M20").Value
    ActiveSheet.Range("G20").Value = ActiveSheet.Range("N20").Value
    
    ' Move the temporary values to M20 and N20
    ActiveSheet.Range("M20").Value = tempF
    ActiveSheet.Range("N20").Value = tempG
End Sub

Sub zLineTwentyButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M20 and N20 in temporary variables
    tempM = ActiveSheet.Range("M20").Value
    tempN = ActiveSheet.Range("N20").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M20").Value = ActiveSheet.Range("T20").Value
    ActiveSheet.Range("N20").Value = ActiveSheet.Range("V20").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T20").Value = tempM
    ActiveSheet.Range("V20").Value = tempN
End Sub

Sub zLineTwentyOneButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F21 and G21 in temporary variables
    tempF = ActiveSheet.Range("F21").Value
    tempG = ActiveSheet.Range("G21").Value
    
    ' Swap the values: Move M21 and N21 to F21 and G21
    ActiveSheet.Range("F21").Value = ActiveSheet.Range("M21").Value
    ActiveSheet.Range("G21").Value = ActiveSheet.Range("N21").Value
    
    ' Move the temporary values to M21 and N21
    ActiveSheet.Range("M21").Value = tempF
    ActiveSheet.Range("N21").Value = tempG
End Sub

Sub zLineTwentyOneButtonLRTwo()
    Dim tempM As Variant
    Dim tempN As Variant
    
    ' Store the values from cells M21 and N21 in temporary variables
    tempM = ActiveSheet.Range("M21").Value
    tempN = ActiveSheet.Range("N21").Value
    
    ' Swap the values: Move T and V to M and N
    ActiveSheet.Range("M21").Value = ActiveSheet.Range("T21").Value
    ActiveSheet.Range("N21").Value = ActiveSheet.Range("V21").Value
    
    ' Move the temporary values to T and V
    ActiveSheet.Range("T21").Value = tempM
    ActiveSheet.Range("V21").Value = tempN
End Sub

Sub zLineTwentyTwoButtonLROne()
    Dim tempF As Variant
    Dim tempG As Variant
    
    ' Store the values from cells F22 and G22 in temporary variables
    tempF = ActiveSheet.Range("F22").Value
    tempG = ActiveSheet.Range("G22").Value
    
    ' Swap the values: Move M22 and N22 to F22 and G22
    ActiveSheet.Range("F22").Value = ActiveSheet.Range("M22").Value
    ActiveSheet.Range("G22").Value = ActiveSheet.Range("N22").Value
    
    ' Move the temporary values to M22 and N22
    ActiveSheet.Range("M22").Value = tempF
    ActiveSheet.Range("N22").Value = tempG
End Sub

Public Sub ImportData()
' ***INITIALIZE***
    ' Database Objects
    Dim SqlDb As New ADODB.Connection
    Dim SqlData As New ADODB.Recordset
    Dim SqlConn As String
    SqlConn = "Provider=SQLOLEDB;Data Source=pbebdsmessier.nos.boeing.com;User ID=SvcIedwhExcel;Password=0x9F4DE51F3477553AE46A752DD6B66FCE;"
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
   ' msgAlert = MsgBox("Data Updated with Run Time of " & Round(Timer - StartTime, 1) & " second", vbOKOnly, "Data Pull Complete")
End Sub







































