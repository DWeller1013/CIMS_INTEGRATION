'------------------------------------------------------------------------------------------------------------------------
'
' Program description: Macro to pull all Detail #'s, COST, CC, RATE and DESCRIPTION from
'                      DIE/Part sheets and paste into 'CIMS' sheet.
'                      *To easily copy paste data into CIMS->Quotes.
'
'------------------------------------------------------------------------------------------------------------------------
' 230307(dmw)-Removing Labor(Hourly rate) from INTRO sheet and from filling out each DIE and Part sheet. All code that changed those values commented out.
' 220609(dmw)-Formula updated on Sheet '600 Laser-Quality'  (F18). Fixed margin not calculating correctly. Ex: Multiplied 10% not 110%.
' 220601(dmw)-Updating spreadsheet, no code changed per John Deman's request.
'               ~Changed sheets 700/800 to add more columns with more information for Material costs.
'                   ~Added columns: $ per LB, Tryout Stock PCS and Tryout Stock Cost
'               ~Updated formulas in those columns to adjust for additional information and change of costs.
' 220527(dmw)-Improve performance of macro by disabling screen updating, automatic calculations and events at the
'             beginning of macro. Re-enabled at the end of the macro.
' 220526(dmw)-Per Ryan Walsh's request, fixed issue where grand total differed between sheets
'               ~Created a new property for detail class (marginTotal) and added to CIMS sheet
'               ~Values are calculated based off margin costs and not total costs
'               ~Fixed circular reference error. (200 DIE H34 calculated itself in sum).
'               ~Also, fixed issue where empty rows would not delete on 'CIMS' sheet.
'               ~Commented out 'deleteRows' subprocedure since 'deleteEmptyData' procedure clears all necessary rows in CIMS sheet.
'               ~Added in CIMS sheet a grand total, grand total (margin) and margin difference (margin - total)
' 220505(dmw)-More changes to workbook per John Demans request.
'               ~'INTRO' sheet to add more boxes that are applied by 'Universal Rate Selector Button'.
'                   ~Adding 3 more boxes to separate the 600/700 sheets as parts. (Tooling and parts have separate rates / margins
'                       ~1 extra input box for 600/700 rate (Parts).
'                       ~2 extra boxes to add margin for both tooling and part sheets.
'               ~Intro sheet 'Total' was miscalculated. Was grabbing total without margin. Adjusted where cells pulled from
'                   for each cell 100-600.
' 220503(dmw)-No changes to code. John Deman showed me bunch of screenshots and changes he would like to formulas
'             and the formatting / style of the tables. Just wants everything cleaned up overall and more compact.
' 220428(dmw)-Removing the add-in idea as it would remove itself every time excel was closed.
'             All functionality done through macro-enabled template sheet.
'               ~Adding a button ('CIMS CONVERSION') that calls to Subprocedure 'CIMS'
' 220427(dmw)-Requested by John Deman / Alan McMullen to add additional features.
'               ~ Added Universal Rate selector (Button that takes user input value and sets rate across workbook)
'                   ~ Added Reset button to go along with this.
'               ~ Added Material Cost Aluminum on sheets 700/800.
'               ~ Changed machining to vary with rate not just static cost.
'               ~ Added cost of ALL dies to INTRO page
'               ~ Added Cast Iron / Kirksite switch feature (2 Radio buttons to toggle between the two)
' 211213(dmw)-Adding error handling
' 211202(dmw)-Set columns to autowidth and CIMS sheet position after Grand Total
' 211130(dmw)-Deleted unused / commented out codes. Set columns formatting (decimal places, currency, etc.).
' 211129(dmw)-setTotals function added
' 211124(dmw)-Deleted many subs/functions after meeting with Ryan and finding more requirements and
'             deleting what is not necessary and not needed in both the code / sheets.
' 211119(dmw)-Functions working as intended.
' 211117(dmw)-Testing adding functions to retrieve and copy paste range rather than subroutines.
' 211116(dmw)-sub addSetC working correctly
' 211111(dmw)-Was able to get subs addSetA / B working, C still WIP
' 211109(dmw)-More functionality
' 211108(dmw)-Testing, attempting to make code more "generic".
'            -Also adding more comments where necessary.
'
'------------------------------------------------------------------------------------------------------------------------

' Main function for macro. Takes data from all sheets and copies it into 'CIMS' sheet to copy paste onto CIMS site.
Sub CIMS()
    
    '220527(dmw)-Improve performance by turning off screen updating, automatic calculations and events
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Initialize variables so sheet starts at 2 and cell value starts at 3
    ' due to Detail # starting at A3 for each sheet. And skipping the "INTRO" sheet.
    ' Because of the specific integers used and Sheet name, it must be consistent across all excel files used
    Dim sht As Integer
    Dim maxSht As Integer
    
    ' Create and set collection to store all objects into
    Dim col As Collection
    Set col = New Collection
    
    ' Set max sheet equal to index of grand total
    ' start sht at 100 DIE. Skipping any previous sheets and looping from 100 -> Grand Total
    sht = Sheets("100 DIE").Index
    maxSht = Sheets("Grand Total").Index
    
    ' Check if 'CIMS' sheet exists, if it does, clear and paste new data
    ' if it does not exist, then create blank 'CIMS' sheet.
    If Not (WorksheetExists("CIMS")) Then
        Worksheets.Add.Name = "CIMS"
    Else
        ' Clear CIMS sheet at start of program and make it Active.
        Sheets("CIMS").Cells.Clear
    End If
  
    ' Move CIMS sheet after Grand Total sheet
    If (WorksheetExists("Grand Total")) Then
        Worksheets("CIMS").Move after:=Worksheets("Grand Total")
    End If
  
    ' While loop to go through the all DIE / Part sheets. Stopping at "Grand Total" Sheet.
    While (sht < maxSht)
    
        ' Activate sheet index is currently on and copy selected set to "CIMS" sheet.
        Sheets(sht).Activate
        'Call addSet
        'Call addObjects
        'Call printColl(col)
        
        '-------------------------------------
        ' was not passing values correctly so moved 'Call addSet' here
        
        ' Initialize collection to store all details
        Dim details As detail
    
        ' Initialize start/last row variables and set to specific text string found in sheets
        ' Last row will always be one row below "# of Parts/Tools"
        'Dim lastRow As Integer
        Dim numOfTools As Integer
        'lastRow = Range("A:F").Find(What:="Total", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
        numOfTools = Range("A:F").Find(What:="# of", SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
        lastRow = numOfTools + 1
        
        ' For loop to go through rows and add each detail as object/class and add to collection
        Dim i As Integer
        For i = 3 To lastRow - 3
            
            ' 211213(dmw)-Probably delete this if statement and replace with error handling
            ' it only grabs integers, incase of typo it should be properly handled
            If Application.WorksheetFunction.IsNumber(Range("A" & i).value) = True _
                And Application.WorksheetFunction.IsNumber(Range("B" & i).value) = True Then
                
                ' Create new detail and set values equal to values found in cell
                Set details = New detail
                'details.Index = i - 2
                ' Set detail properties to values in each row across all tables
                ' on each sheet
                details.Set_Detail = Range("A" & i).value
                details.Set_Cost = Range("B" & i).value
                details.Set_CC = Range("C" & i).value
                details.Set_Rate = Range("D" & i).value
                details.Set_Hours = Range("E" & i).value
                details.Set_Description = Range("F" & i).value
                details.Set_Total = Range("G" & i).value
                details.Set_marginTotal = Range("G" & i).value * (Range("H2").value + 1)
                
                ' If # of tools present is > 1 AND Cost Code is NOT 301 - then multiply hours by # Of Tools.
                If (Range("G" & numOfTools).value > 1 And details.Cost <> "301") Then
                    details.Set_Hours = details.Hours * Range("G" & numOfTools).value
                End If
    
                ' Add detail to collection
                col.Add Item:=details
                
            ' If column 'BCDE' are all empty, skip to next line
            ' 'TOTAL' is placed in the A or F row, so check all other columns to be empty
            ElseIf IsEmpty(Range("B" & i)) And _
                IsEmpty(Range("C" & i)) Then
                ' Do nothing and go to next line
            
            ' Add error handling in case If statement is not true
            Else
                ' Grab sheet name for where error occured
                Dim shtName As String
                shtName = ActiveSheet.Name
                MsgBox "Incorrect value on sheet: " & shtName & "." & vbCrLf & "Row " & i & ".", vbOKOnly, "Input error"
                End
            End If
    
        Next i
    
        Set details = Nothing ' release variable
    
        '-------------------------------------
        
        ' Increment and grab next sheet
        sht = sht + 1
        
    Wend
    
    ' Print collection to CIMS
    'Call printColl(col)
    
    '-------------------------------
    ' was not passing values correctly so moved 'Call printColl' here
    
    ' Collection should be populated by all objects at this point
    ' Activate CIMS sheet and populate rows with each object
    
    Dim det As detail
    Dim detNum As Integer
    detNum = 1

    With ThisWorkbook.Worksheets("CIMS")

        For Each detail In col

            ThisWorkbook.Worksheets("CIMS").Range("A" & detNum).value = col.Item(detNum).detail
            ThisWorkbook.Worksheets("CIMS").Range("B" & detNum).value = col.Item(detNum).Cost
            ThisWorkbook.Worksheets("CIMS").Range("C" & detNum).value = col.Item(detNum).CC
            ThisWorkbook.Worksheets("CIMS").Range("D" & detNum).value = col.Item(detNum).Rate
            ThisWorkbook.Worksheets("CIMS").Range("E" & detNum).value = col.Item(detNum).Hours
            ThisWorkbook.Worksheets("CIMS").Range("F" & detNum).value = col.Item(detNum).Description
            ThisWorkbook.Worksheets("CIMS").Range("G" & detNum).value = col.Item(detNum).Total
            ThisWorkbook.Worksheets("CIMS").Range("H" & detNum).value = col.Item(detNum).marginTotal
            
            detNum = detNum + 1

        Next

    End With
    
    '---------------------------------
    
    ' Activate CIMS sheet
    Sheets("CIMS").Activate
    
    ' 220526(dmw)-Commented out as subprocedure below (deleteEmptyData) clears necessary rows
    ' Delete irrelevant data
    'Call deleteRows
    
    ' Based on 'CC' set 'RATE' & 'HOURS' equal to 'TOTAL'
    Call setTotals
    
    ' Insert header row and align all data to the left
    Call addHeader
    
    ' Set currency format for 'TOTAL', 'MARGIN TOTAL' and 'GRAND TOTAL' columns
    Columns(7).NumberFormat = "$#,##0.00"
    Columns(8).NumberFormat = "$#,##0.00"
    Columns(10).NumberFormat = "$#,##0.00"
    Columns(11).NumberFormat = "$#,##0.00"
    Columns(13).NumberFormat = "$#,##0.00"
    
    ' Set 2 decimal places for 'RATE' & 'HOURS' columns
    Columns(4).NumberFormat = "0.00"
    Columns(5).NumberFormat = "0.00"
    
    ' 220526(dmw) - Adding Grand total (margin) and difference between margin and total
    ' Initialize lr as last row variable
    ' Grab sum of all detail totals and place into 'Grand Total'
    Dim LR As Long
    LR = Cells(rows.Count, "A").End(xlUp).Row
    
    ' Grand total
    Sheets("CIMS").Range("J2") = "=SUM(G2:G" & LR & ")"
    
    ' Grand total (margin)
    Sheets("CIMS").Range("K2") = "=SUM(H2:H" & LR & ")"
    
    ' Margin difference
    Sheets("CIMS").Range("M2") = "=K2-J2"
    
    
    'Center text in first row (header row)
    Worksheets("CIMS").rows(1).HorizontalAlignment = xlCenter
    
    ' Set auto width on all columns
    Worksheets("CIMS").Columns("A:M").AutoFit
    
    ' 220526(dmw)-Sub procedure working correctly
    ' Delete empty data from CIMS sheet
    Call deleteEmptyData
    
    ' Copy from 2nd row to last row (LR)
    'Sheets("CIMS").Range("A2:F" & LR).Copy
    'Range("$A$1").Select
    
    ' 220527(dmw)-Improve performance by re-enabling screen updating, automatic calculations and events at end of macro
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' 220505(dmw) - Added more input boxes. Making code more generic by removing statically assigned cells.
'                   ~Adding variables to search for labor and margin rates and assign those values throughout sheet.
' 220425(dmw) - Adding button to Intro form.
' Button used to grab value from 'Universal Rate' cell and apply it to all 'Rate' values throughout workbook.
Sub universalRateBtn_Click()
    
    With Sheets("INTRO")
    
        ' 230307(dmw) - Commenting all laborValue relevant code out.
        ' Initialize / find labor value
        'Dim laborValueRN As Integer
        'Dim laborValue As Integer ' For rate on sheets 100-600
        'Dim laborValueParts As Integer ' For rate on sheets 700 & 800
        
        ' Find rates value and assign to variables
        'laborValueRN = .Range("A:F").Find(What:="Labor", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        'laborValue = .Range("B" & laborValueRN + 1).value
        'laborValueParts = .Range("B" & laborValueRN + 2).value
        
        ' Initialize / find margin value
        Dim marginValueRN As Integer
        Dim marginValue As Double ' For magin on sheets 100-600
        Dim marginValueParts As Double ' For margin on sheets 700 & 800
        
        ' Find margin value and assign to variables
        marginValueRN = .Range("A:F").Find(What:="Margin (%)", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        marginValue = .Range("B" & marginValueRN + 1).value
        marginValueParts = .Range("B" & marginValueRN + 2).value
        
        ' Confirmation box to appear for user
        confirmationBox = MsgBox("Are you sure you wish to set all Rate values across entire workbook to: " & vbCr & vbCr & _
        vbCr & vbCr & "Margin value (%): " & vbCr & vbTab & "Tooling: " & CStr(marginValue * 100) & "%" & vbCr & vbTab & "Parts: " & _
        CStr(marginValueParts * 100) & "%", vbYesNoCancel, "Apply universal rate")
        
        ' If user hits yes
        If confirmationBox = vbYes Then
    
            ' Catch errors if non-integer is input
            On Error GoTo inputError
            
            ' 220505(dmw) - Commenting below 4 lines of code out, changing variables and no need to declare a second time.
            ' -------------------------------------
            ' Initialize universalRate variable
            'Dim universalRate As Integer
            ' Set universal rate equal to user input value
            'universalRate = Sheets("INTRO").Range("B17")
               
            ' Go through each sheet in workbook and set rate value
            'Sheets("100 DIE").Range("D3:D14").value = laborValue
            'Sheets("200 DIE").Range("D3:D13").value = laborValue
            'Sheets("300 DIE").Range("D3:D13").value = laborValue
            'Sheets("400 DIE").Range("D3:D13").value = laborValue
            'Sheets("500 DIE").Range("D3:D13").value = laborValue
            'Sheets("500 Hammer Form").Range("D3:D7").value = laborValue
            'Sheets("600 Laser-Quality").Range("D3:D12").value = laborValue
            'Sheets("700 Parts").Range("D3:D15").value = laborValueParts
            'Sheets("800 Parts").Range("D3:D15").value = laborValueParts
            
            ' Go through each sheet in workbook and set margin value
            Sheets("100 DIE").Range("H2").value = marginValue
            Sheets("200 DIE").Range("H2").value = marginValue
            Sheets("300 DIE").Range("H2").value = marginValue
            Sheets("400 DIE").Range("H2").value = marginValue
            Sheets("500 DIE").Range("H2").value = marginValue
            Sheets("500 Hammer Form").Range("H2").value = marginValue
            Sheets("600 Laser-Quality").Range("H2").value = marginValue
            Sheets("700 Parts").Range("H2").value = marginValueParts
            Sheets("800 Parts").Range("H2").value = marginValueParts
            
            Exit Sub
            
        End If
    
    End With
    
' Catch error if user inputs non-integers and display a message box
inputError:
        inputErrorBox = MsgBox("Error: You must enter number values only into universal rate.", vbOKOnly, "Input Error")
              
End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' 220505(dmw) - Changed range on 'Intro' sheet from B17 to B18:C19.
' 220425(dmw) - Adding button to Intro form.
' Button used to reset all 'Rate' values in workbook.
Sub universalRateResetBtn_Click()
    
    ' Confirmation box to appear for user
    confirmationBox = MsgBox("Are you sure you wish to reset all Rate values across entire workbook to empty/null?", vbYesNoCancel, "Reset All")
    
    If confirmationBox = vbYes Then
    
        ' Go through each sheet in workbook and set value to blank
        Sheets("100 DIE").Range("D3:D14").value = ""
        Sheets("200 DIE").Range("D3:D13").value = ""
        Sheets("300 DIE").Range("D3:D13").value = ""
        Sheets("400 DIE").Range("D3:D13").value = ""
        Sheets("500 DIE").Range("D3:D13").value = ""
        Sheets("500 Hammer Form").Range("D3:D7").value = ""
        Sheets("600 Laser-Quality").Range("D3:D12").value = ""
        Sheets("700 Parts").Range("D3:D15").value = ""
        Sheets("800 Parts").Range("D3:D15").value = ""
        
        Sheets("INTRO").Range("B18:C19").value = ""
        
        ' Go through each sheet in workbook and set margin value to blank
        Sheets("100 DIE").Range("H2").value = ""
        Sheets("200 DIE").Range("H2").value = ""
        Sheets("300 DIE").Range("H2").value = ""
        Sheets("400 DIE").Range("H2").value = ""
        Sheets("500 DIE").Range("H2").value = ""
        Sheets("500 Hammer Form").Range("H2").value = ""
        Sheets("600 Laser-Quality").Range("H2").value = ""
        Sheets("700 Parts").Range("H2").value = ""
        Sheets("800 Parts").Range("H2").value = ""
        
    End If
    
End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' ------------------------------------------------------------------------
' 220425(dmw) - | Radio buttons to toggle between Kirksite and Cast Iron |
' ------------------------------------------------------------------------
' If Kirksite Radio button is selected change values
Sub kirksiteRadioBtn_Click()
    
    ' Initialize ctr (counter) variable
    Dim ctr As Integer
    ctr = 1
    
    ' Initialize pourChargeRN variable (Pour Charge Row Number)
    Dim pourChargeRN As Integer
    
    ' While loop to go through 100-500 DIE sheets and change section
    While ctr <= 5
        
        With Sheets(CStr(ctr) & "00 DIE")
        
            ' Declare variable for pour charge row number
            'Dim pourChargeRN As Integer
            pourChargeRN = .Range("A:F").Find(What:="Pour Charge -", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            
            ' Set Header for Pour Chage section and formula
            .Range("A" & pourChargeRN).value = "Pour Charge - Kirksite"
            .Range("E" & pourChargeRN + 2).Formula = "=D" & pourChargeRN + 2 & "*0.31*1.1"
            
            ctr = ctr + 1
            
        End With
        
    Wend
    
    ' Change formula and pour charge header to kirksite on 500 hammer form sheet
    With Sheets("500 Hammer Form")
    
        ' Declare variable for pour charge row number
        'Dim pourChargeRN As Integer
        pourChargeRN = .Range("A:F").Find(What:="Pour Charge -", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
        ' Set Header for Pour Chage section and formula
        .Range("A" & pourChargeRN).value = "Pour Charge - Kirksite"
        .Range("E" & pourChargeRN + 2).Formula = "=D" & pourChargeRN + 2 & "*0.31*1.1"
        
    End With

End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' ------------------------------------------------------------------------
' 220425(dmw) - | Radio buttons to toggle between Kirksite and Cast Iron |
' ------------------------------------------------------------------------
' If Cast Iron Radio button is selected change values
Sub castIronRadioBtn_Click()
    
    ' Initialize ctr (counter) variable
    Dim ctr As Integer
    ctr = 1
    
    ' Initialize pourChargeRN variable (Pour Charge Row Number)
    Dim pourChargeRN As Integer
    
    ' While loop to go through 100-500 DIE sheets and change section
    While ctr <= 5
        
        With Sheets(CStr(ctr) & "00 DIE")
        
            ' Declare variable for pour charge row number
            'Dim pourChargeRN As Integer
            pourChargeRN = .Range("A:F").Find(What:="Pour Charge -", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            
            ' Set Header for Pour Chage section and formula
            .Range("A" & pourChargeRN).value = "Pour Charge - Cast Iron"
            .Range("E" & pourChargeRN + 2).Formula = "=D" & pourChargeRN + 2 & "*1.19*1.1"
            
            ctr = ctr + 1
            
        End With
        
    Wend
    
    ' Change formula and pour charge header to kirksite on 500 hammer form sheet
    With Sheets("500 Hammer Form")
    
        ' Declare variable for pour charge row number
        'Dim pourChargeRN As Integer
        pourChargeRN = .Range("A:F").Find(What:="Pour Charge -", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
        ' Set Header for Pour Chage section and formula
        .Range("A" & pourChargeRN).value = "Pour Charge - Cast Iron"
        .Range("E" & pourChargeRN + 2).Formula = "=D" & pourChargeRN + 2 & "*1.19*1.1"
        
    End With

End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' Function used to check if Worksheet exists or not and return True or False
' after sorting through all sheets in file.
' (used to check if 'CIMS' and 'Grand Total' sheet exist)
Private Function WorksheetExists(ByVal WorksheetName As String) As Boolean

    Dim sht As Worksheet
    
    ' Loop through each sheet in current workbook and through every worksheet
    For Each sht In ActiveWorkbook.Worksheets
    
        If sht.Name = WorksheetName Then
            WorksheetExists = True
            Exit Function
        End If
        
    Next sht
    
    WorksheetExists = False

End Function

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' MOVED - No longer a sub procedure or called and now code copied to main function 'CIMS'.
' 211206(dmw)-Another test subprocedure.
' Rather than copying range of values, those rows
' will be copied into an object so the values can
' be manipulated and then pasted onto 'CIMS' sheet.
Sub addObjects()

    ' Initialize collection to store all details
    Dim details As detail
    Dim col As Collection
    Set col = New Collection
    
    ' Initialize start/last row variables and set to specific text string found in sheets
    Dim lastRow As Integer
    Dim numOfTools As Integer
    'lastRow = Range("A:F").Find(What:="Total", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    numOfTools = Range("A:F").Find(What:="# of", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastRow = numOfTools + 1
        
    ' For loop to go through rows and add each detail as object/class and add to collection
    Dim i As Integer
    For i = 3 To lastRow - 3
    
        If Application.WorksheetFunction.IsText(Range("A" & i).value) = False Then
        
            ' Create new detail and set values equal to values found in cell
            Set details = New detail
            'details.Index = i - 2
            ' Set detail properties to values in each row across all tables
            ' on each sheet
            details.Set_Detail = Range("A" & i).value
            details.Set_Cost = Range("B" & i).value
            details.Set_CC = Range("C" & i).value
            details.Set_Rate = Range("D" & i).value
            details.Set_Hours = Range("E" & i).value
            details.Set_Description = Range("F" & i).value
            details.Set_Total = Range("G" & i).value
            details.Set_marginTotal = Range("G" & i).value * (Range("H2").value + 1)
            
            ' If # of tools present is > 1  AND Cost Code is not 301 - then multiply hours by # Of Tools.
            If (Range("G" & numOfTools).value > 1 & details.Cost <> 301) Then
                details.Set_Hours = details.Hours * Range("G" & numOfTools).value
            End If
            
            ' Add detail to collection
            'col.Add Name, "detail" & i
            col.Add Item:=details
            
        End If
               
    Next i

    Set details = Nothing ' release variable

End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' MOVED - No longer a sub procedure or called and now code copied to main function 'CIMS'.
' Sub procedure to take collection as parameter and paste values into CIMS worksheet
Sub printColl(ByVal myCol As Collection)

    ' Collection should be populated by all objects at this point
    ' Activate CIMS sheet and populate rows with each object
    'Sheets("CIMS").Activate
    
    Dim det As detail
    Dim detNum As Integer
    detNum = 1
    
    ' Grab 'CIMS' worksheet within current workbook.
    With ThisWorkbook.Worksheets("CIMS")
    
        For Each detail In myCol
        
            Range("A" & detNum).value = col.Item(detNum).detail
            Range("B" & detNum).value = col.Item(detNum).Cost
            Range("C" & detNum).value = col.Item(detNum).CC
            Range("D" & detNum).value = col.Item(detNum).Rate
            Range("E" & detNum).value = col.Item(detNum).Hours
            Range("F" & detNum).value = col.Item(detNum).Description
            
            detNum = detNum + 1
            
        Next
    
    End With
End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' MOVED - No longer a sub procedure or called and now code copied to main function 'CIMS'.
' 211129(dmw)-Will function correctly as log as header rows take up rows 1 & 2
'             and the details start on row 3 (A3:F).
' 211109(dmw)-Add first set of details (YELLOW BACKGROUND)
Sub addSet()
    
    ' Find "Total" string and set lastRow equal to that cell and copy the range
    Dim lastRow As Integer
    lastRow = Range("A:F").Find(What:="Total", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Range("A3:G" & lastRow).Copy
    
    ' Activate CIMS sheet and paste in last available blank cell
    Sheets("CIMS").Activate
    Cells(Range("A" & rows.Count).End(xlUp).Row + 1, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
          
End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' 211109(dmw)-Turning header section from Sub Main() to it's own sub procedure
Sub addHeader()

    ' Insert new row for A1
    rows(1).EntireRow.Insert
    
    ' Create headers in first row, created after "Copy" so it is not included when pasted to CIMS.
    Range("$A$1").value = "DETAIL #"
    Range("$B$1").value = "COST"
    Range("$C$1").value = "CC"
    Range("$D$1").value = "RATE"
    Range("$E$1").value = "HOURS"
    Range("$F$1").value = "DESCRIPTION"
    Range("$G$1").value = "TOTAL"
    Range("$H$1").value = "MARGIN TOTAL"
    Range("$J$1").value = "GRAND TOTAL"
    Range("$K$1").value = "GRAND TOTAL (MARGIN)"
    Range("$M$1").value = "MARGIN DIFFERENCE (MARGIN - TOTAL)"
    
End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' 211129(dmw)-Function to set total for details containing 301 cost code
'        ~Rows containing 301 cost code do not have hours, per Ryan and total is manually entered.
'        ~To combat the CIMS issue of not being able to manually type total.
'        ~If CC is 'Subcontract Work' Then set rate = total on sheet and hours = 1.
Function setTotals()

    Dim currentSheet As Worksheet
    Set currentSheet = ThisWorkbook.ActiveSheet
    
    Dim i As Long
    For i = 1 To 5000
        If (Cells(i, 3).value = "Subcontract Work" Or Cells(i, 3).value = "MISC" _
            Or Cells(i, 3).value = "Material Purchase") Then
            Cells(i, 4).value = Cells(i, 7).value
            Cells(i, 5).value = 1
        End If
    Next i
    
End Function

' 211108(dmw)-Sub to delete all empty rows for both 'A' and 'B' within 'CIMS' worksheet.
Sub deleteEmptyRow()

    With Sheets("CIMS")
    
        On Error Resume Next
        Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        Columns("B").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        
    End With
    
End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' 220526(dmw)-Commented out code that calls this subprocedure. Sub 'deleteEmptyData' covers what this sub should do.
' 211108(dmw)-Delete all rows where A is either a non integer or blank within 'CIMS' worksheet.
Sub deleteRows()

    ' Delete all rows that do not include an integer or value at all
    ' in columns A & B
    Dim LR3 As Long, i3 As Long
    With Sheets("CIMS")
        ' Set LR3 equal to last row containing data
        LR3 = .Range("A" & .rows.Count).End(xlUp).Row
        
        ' Check each row for blank A/B column or non-numeric and delete row
        For i3 = LR3 To 2 Step -1
            If Not IsNumeric(.Range("A" & i3).value) Or _
            .Range("A" & i3).value = "" Or .Range("B" & i3).value = "" _
            Then .rows(i3).Delete
        Next i3
        
        ' Delete all rows where "A" column value contains a value other than a die number (<100 >800)
        Dim rows As Range, Cell As Range, value As Long
        Set Cell = Range("A2")
        Do Until Cell.value = ""
            value = Val(Cell.value)
            If (value < 100 Or value > 800) Then
                If rows Is Nothing Then
                    Set rows = Cell.EntireRow
                Else
                    Set rows = Union(Cell.EntireRow, rows)
                End If
            End If
            Set Cell = Cell.Offset(1)
        Loop
        If Not rows Is Nothing Then rows.Delete
    End With
    
    ' Don't throw error and continue on with code to delete rows
    On Error Resume Next
    Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Columns("B").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
End Sub

' -------------------------------------------------------------------------------
' -------------------------------------------------------------------------------
' 220526(dmw)-Subprocedure not working properly. Reworking code to provide proper functionality.
' 220214(dmw)-Delete all rows containing 0 under Rate, Cost & Total
' on CIMS sheet. Sub created to remove unnecessary data that should not
' be copied over.
Sub deleteEmptyData()

    ' Declare variables (i as counter for For Loop, LR as last row)
    Dim i As Integer
    Dim LR As Long
    LR = Range("G" & rows.Count).End(xlUp).Row
    
    With Sheets("CIMS")
    
        For i = LR To 2 Step -1
            If Range("G" & i).value = 0 Then rows(i).Delete
        Next i
        
    End With

End Sub
