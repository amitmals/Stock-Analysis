'HW-02 Submission
'Created by Amit Malik
'Step1: The script is run over all the sheets starting with the 1st
'Step2: The data in each sheet is 1st sorted on Date then Ticker name
'       All the data now is ready in a serial manner
'Step3: Initialize the open $ for the 1st ticker from C2 as "yearStart" for the ticker
'       Intialize the var j to which will be used to write the output in colI to colL as 2
'Step4: Start looping thru colA from 2 till end of file
'Step5: Compare the values in colA between 2 consecutive entries if the data is valid(Some data has 0 Start/End/Vol)
'       IF the tickers are the same then just update the volumne in colL
'       IF the tickers are not the same
'           Now we know we are in the last entry for the ticker
'           Update the volumne in colL for the ticker the last time
'           We can now get the close $ for year end for the ticker and save that in "yearEnd"
'           Calc the Yearly change and Percentage and save it in colJ & colK
'           added the formatting as needed
'           We can now get the open $ for year start for the next ticker and save that in "yearStart"
'           update the counter j so we write the next ticker info in the correct colI to colL
'Step6: Now that the yearStart is populated repeat the step5 for the next entry in colA for the loop started in step4
'Step7: Now loop through colI to colL and find the max/min yearly changes and the max total vol
'       Populate the data in colP-ColQ
'Step8: Autofit and adjust the formatting as requested in the HW                 

Sub akm()

    'Lets define our variables
    Dim i as Long
    Dim j As Long
    'Used to store the lastRow in excel col under consideration
    Dim lastRow As Long
    'yearStart and yearEnd store the Opening and closing prices for any ticeker
    Dim yearStart as Double
    Dim yearEnd As Double

    'Step1: Loop thru all the sheets in the doc
    For Each currentWS In Worksheets
        MsgBox "Lets work on: " + currentWS.Name    
        'Step2: Lets sort the columns out. Data is sorted by ticker and then date
        'Sorting makes sure the data is alligned so all the entries with 1 ticker are together 
        currentWS.Range("A1:G1", currentWS.Range("A1:G1").End(xlDown)).Sort Key1:=currentWS.Range("B1"), Order1:=xlAscending, Header:=xlNo
        currentWS.Range("A1:G1", currentWS.Range("A1:G1").End(xlDown)).Sort Key1:=currentWS.Range("A1"), Order1:=xlAscending, Header:=xlNo

        'Lets get the last entry in the worksheet
        lastRow = currentWS.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Step3: j->counter for unique tickers column I. This is used to place the date in ColI
        j = 2
        yearStart = currentWS.Cells(2, 3).Value
        
        'Step4: Lets cycle through all the entries
        For i = 2 To lastRow
            'Step5: Data Check:Only consider the row as valid if there is data in it
            If (currentWS.Cells(i,3).Value + currentWS.Cells(i,4).Value + currentWS.Cells(i,7).Value) > 0 Then
                'If the next entry in the WS ColA is diff then we have a new ticker
                If currentWS.Cells(i, 1).Value <> currentWS.Cells(i + 1, 1).Value Then
                    'Add the new Ticker in the Ticker Column
                    currentWS.Cells(j, 9).Value = currentWS.Cells(i, 1).Value
                    'Sum the Volumne
                    currentWS.Cells(j, 12).Value = currentWS.Cells(j, 12).Value + currentWS.Cells(i, 7)
                    'For the last Entry for the Ticker, since data is sorted based on date
                    '-> we get yearEnd value
                    yearEnd = currentWS.Cells(i, 6).Value
                    'Lets calc the Yearly change and percentage
                    currentWS.Cells(j, 10).Value = yearEnd - yearStart
                    currentWS.Cells(j, 11).Value = (yearEnd - yearStart) / yearStart
                    'Cond format the Yearly change and format the Percentage
                    If currentWS.Cells(j, 10) >= 0 Then
                        currentWS.Cells(j, 10).Interior.ColorIndex = 4
                    Else: currentWS.Cells(j, 10).Interior.ColorIndex = 3
                    End If
                    'We now can get the information for the yearStart for the next ticker
                    yearStart = currentWS.Cells(i + 1, 3).Value
                    'increase j->counter for unique tickers
                    j = j + 1
                Else
                    'When the ticker is same between the 2 rows. Update the volumne only
                    currentWS.Cells(j, 12).Value = currentWS.Cells(j, 12).Value + currentWS.Cells(i, 7)
                End If
            Else
            'If the row is invalid, the yearStart for the next row is moved
            yearStart = currentWS.Cells(i + 1, 3).Value    
            End If
            'Step6: Keep looping
        Next i

        currentWS.Cells(1, 9).Value = "Ticker"
        currentWS.Cells(1, 10).Value = "Yearly Change"
        currentWS.Cells(1, 11).Value = "Percentage"
        currentWS.Cells(1, 12).Value = "Total Stock Volumne"

        'Hard Level-Add additional information in ColO-ColQ
        'Lets get the last entry in the worksheet
        lastRow = currentWS.Cells(Rows.Count, 9).End(xlUp).Row
        'Step7: Let run thru the data we created earlier and find the max/min as needed
        For i = 2 To lastRow
            if currentWS.Cells(i,11).Value > currentWS.Cells(2,17).Value then
                currentWS.Cells(2,17).Value = currentWS.Cells(i,11).Value
                currentWS.Cells(2,16).Value = currentWS.Cells(i,9).Value
            end if
            if currentWS.Cells(i,11).Value < currentWS.Cells(3,17).Value then
                currentWS.Cells(3,17).Value = currentWS.Cells(i,11).Value
                currentWS.Cells(3,16).Value = currentWS.Cells(i,9).Value
            end if
            if currentWS.Cells(i,12).Value > currentWS.Cells(4,17).Value then
                currentWS.Cells(4,17).Value = currentWS.Cells(i,12).Value
                currentWS.Cells(4,16).Value = currentWS.Cells(i,9).Value
            end if           
        Next i

        'Step8: Let print output headers and autofit all the columns
        currentWS.Cells(2, 15).Value = "Greatest % Increase"
        currentWS.Cells(3, 15).Value = "Greatest % Decrease"
        currentWS.Cells(4, 15).Value = "Greatest Total Volume"
        currentWS.Cells(1, 16).Value = "Ticker"
        currentWS.Cells(1, 17).Value = "Value"
        currentWS.Range("Q2:Q3").NumberFormat = "0.00%"
        currentWS.Range("K2:K" &lastRow).NumberFormat = "0.00%"
'        currentWS.Range("J2:J" &lastRow).NumberFormat = "0.00000000"
        currentWS.Range("A:Q").Columns.Autofit
        
        MsgBox "Output is created for: " + currentWS.Name 
    
    Next currentWS
    MsgBox "Macro has finished all your work!!!"
End Sub
