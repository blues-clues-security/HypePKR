UNCLASSIFIED

'''Sheet 1'''
Private Sub button1()

Dim btn As Button
'Application.ScreenUpdating = False
ActiveSheet.Buttons.Delete
Dim t As Range

Set t = ActiveSheet.Range("M1", "O1")
Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
        .OnAction = "parse_this_shit"
        .Caption = "Start"
        .Name = "Start"
    End With
    
Set t = ActiveSheet.Range("Q1", "S1")
Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
        .OnAction = "delete_all"
        .Caption = "Clear"
        .Name = "Clear"
    End With

'Application.ScreenUpdating = True



End Sub


Private Sub ip_list_generate()

'''Create IP List sheet to enter device information
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "IP List"
    Range("A1").Value = "IP"
    Range("B1").Value = "Device Name"
    
'''Info Message
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "Create IP List Here Before Running Nessus Parsing Script"
    With Selection.Font
        .Name = "Calibri"
        .Size = 28
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "Don't forget to switch back to main data page!"
    With Selection.Font
        .Name = "Calibri"
        .Size = 24
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With


End Sub
'''End Sheet'''
'''Start Workbook'''

Sub delete_all()

Application.DisplayAlerts = False

ActiveWorkbook.Worksheets("Working").Delete
ActiveSheet.Range("A:A", "M:M").Value = ""

Application.DisplayAlerts = True

End Sub

Sub parse_this_shit()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Declaring Variables''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim selectBox As String
Dim i As Long

'''Creates a dialog box for user to enter which type of data they are parsing
selectBox = InputBox("Select the following:" & Chr(13) & " 1 - Nessus" & Chr(13) & " 2 - SCC" & Chr(13) & " 0 - Exit Program")

'''This is some error validation to check if any text was entered instead of numbers, this will loop until a number is entered
Do Until IsNumeric(selectBox) = True
If IsNumeric(selectBox) = True Then
    selectBox = InputBox("You need to use numbers, silly" & Chr(13) & "Select the following:" & Chr(13) & " 1 - Nessus" & Chr(13) & " 2 - SCC" & Chr(13) & " 0 - Exit Program")
Else: If selectBox = "" Then Exit Sub
End If
Loop

'''This is a conditional loop that will run one of the macros below based on the selection made in the dialog box
'' 1 is Nessus, 2 is SCC
For i = 1 To selectBox
Select Case True
    Case selectBox = 1
'''Adding in portion to check if an IP List was loaded
        Select Case MsgBox("Did you forget to update your IP List?" & Chr(13) & "The 'IP List' sheet is used to correlate Hostname to IP." & Chr(13) & "If you have another Excel sheet with this info, copy and paste it here", vbYesNo, "IP List Check")
            Case vbYes
                Worksheets("IP List").Activate
                Exit Sub
            Case vbNo
                Call nessus_parse2
        End Select
    Case selectBox = 2
        MsgBox ("Currently SCC does not support IP List capability." & Chr(13) & "Please ensure all your XCCDF CSV's are named according to the hostname of the device it scanned")
        Call scc_parse
    Case selectBox = 0
        Exit Sub
    Case selectBox = ""
        Exit Sub
    Case selectBox <> 1, 2, 0
''I,Robot reference
        MsgBox "My responses are limited. You must ask the right questions.", , "Dr. Lanning says..."
End Select
selectBox = InputBox("Select the following:" & Chr(13) & " 1 - Nessus" & Chr(13) & " 2 - SCC" & Chr(13) & " 0 - Exit Program")
Next i

End Sub




Sub nessus_parse2()

Dim path, ThisWb, filename, output_sheet, primary_sheet, newsht_name() As String
Dim counter As Long
Dim wbDest, Wkb As Workbook
Dim shtDest, ws As Worksheet
Dim CopyRng, Dest As Range

'''Turns off showing data flying all over the screen
Application.ScreenUpdating = False

'''Used for clarity
ThisWb = ActiveWorkbook.Name

'''Where user inputs files for parsing
path = InputBox("Please enter the full file path to the Nessus CSV's")

'''To explicitly define path
'path = "C:\Users\thomas.blauvelt\Desktop\Blauvelt\OF_working\Scans\for_rmp\SCC"

'''Counter used to input device name for each row of data
counter = Range("A1", Range("A64000").End(xlUp)).Count

'''Sets the variable shtDest which looks at the workbook as an array, using the 2nd sheet (0 is the first)
Set shtDest = ActiveWorkbook.Sheets(1)
filename = Dir(path & "\*.csv", vbNormal)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''File Loops'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Grabs data from each file within the directory and inputs the device name into each based off the filename
''*****NOTE*******
''If you use "." in the filename this will error
If Len(filename) = 0 Then Exit Sub
Do Until filename = vbNullString
    If Not filename = ThisWb Then
        Set Wkb = Workbooks.Open(filename:=path & "\" & filename)
        counter = Range("A1", Range("A64000").End(xlUp)).Count
        newsht_name = Split(filename, ".")
        Range("D1").Value = "Device Name"
        Set CopyRng = Wkb.Sheets(1).Range(Cells(2, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column))
        Set Dest = shtDest.Range("A" & shtDest.Cells(Rows.Count, 1).End(xlUp).Row + 1)
        CopyRng.Copy
        Dest.PasteSpecial xlPasteFormats
        Dest.PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        Wkb.Close False
    End If
    
    filename = Dir()
Loop

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Declaring Variables''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Variables are in order in which they appear in code

Dim ouput_sheet As String

Dim sheet_range As String
Dim column_h As String
Dim columns_all As String
Dim column_e As String

Dim incrementer As Integer
Dim pluginid As Long
Dim IP As String
Dim ip_check As String
Dim pluginid_combined As String

Dim cve_combined As String
Dim ip_combined As String
Dim ip_array() As String
Dim hostname_array() As String
Dim ip_address As String
Dim hostname_combined As String
Dim i As Long

Dim plugin_high As Long

Dim ip_count As Integer
Dim index As Integer
Dim ip_holder As Integer

primary_sheet = ActiveSheet.Name
output_sheet = "Working"

Application.ScreenUpdating = False

'''Creates a new sheet, sets the name, and width of the rows
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = output_sheet

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Moving data to correct headings''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Note - need to be on the nessus data output sheet in order for this to work properly
    
'''Column A Move
    Worksheets(primary_sheet).Activate
    Columns("I:I").Copy
    Worksheets(output_sheet).Activate
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Value = "Finding Title"
    

'''Column B Move
    Worksheets(primary_sheet).Activate
    Columns("J:J").Copy
    Worksheets(output_sheet).Activate
    Range("B1").Select
    ActiveSheet.Paste
    Range("B1").Value = "Finding Description"
    
'''Column C Move
    Worksheets(primary_sheet).Activate
    Columns("D:D").Copy
    Worksheets(output_sheet).Activate
    Range("C1").Select
    ActiveSheet.Paste
    Range("C1").Value = "Vulnerability"
        
'''Column D Move
    Worksheets(primary_sheet).Activate
    'Columns("D:D").Copy
    Worksheets(output_sheet).Activate
    'Range("B1").Select
    'ActiveSheet.Paste
    Range("D1").Value = "Device Name"
        
'''Column E Move
    Worksheets(primary_sheet).Activate
    Columns("E:E").Copy
    Worksheets(output_sheet).Activate
    Range("E1").Select
    ActiveSheet.Paste
    Range("E1").Value = "IP Address"
        
'''Column F Move
    Worksheets(primary_sheet).Activate
    Columns("B:B").Copy
    Worksheets(output_sheet).Activate
    Range("F1").Select
    ActiveSheet.Paste
    Range("F1").Value = "Reference"
        
'''Column G Move
    Worksheets(primary_sheet).Activate
    Columns("K:K").Copy
    Worksheets(output_sheet).Activate
    Range("G1").Select
    ActiveSheet.Paste
    Range("G1").Value = "Fix Recommendation"
    
'''Column H Move
    Worksheets(primary_sheet).Activate
    Columns("A:A").Copy
    Worksheets(output_sheet).Activate
    Range("H1").Select
    ActiveSheet.Paste
    Range("H1").Value = "Comment"
    
'''Selects the new output sheet to view and sets the column width for readability
    ActiveSheet.Name = output_sheet
    Selection.RowHeight = 15
    Columns("A:H").ColumnWidth = 20

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Sorting on PluginID (Comment)''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''Creates a counter value for every row in data
counter = Range("A1", Range("A64000").End(xlUp)).Count - 1
'''Creates a string value for every row in data
sheet_range = Trim(Str(counter))
'''Creates total range for each column for the loop to use
column_h = "H2:H" + sheet_range
column_all = "A1:H" + sheet_range
column_e = "E2:E" + sheet_range

''Sort data by PluginID (ascending) then by IP Address (ascending)
ActiveWorkbook.Worksheets(output_sheet).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(output_sheet).Sort.SortFields.Add Key:=Range( _
        column_h), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(output_sheet).Sort.SortFields.Add Key:=Range( _
        column_e), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets(output_sheet).Sort
        .SetRange Range(column_all)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Group CVEs (Reference) by IP Address''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Place 1 in first count column because it will be the first unique value
''This is used to create a column for unique Plugin IDs and IPs
incrementer = 1
Range("I2").Value = incrementer

''Sets pluginID to equal the first value in H for comparison
pluginid = Trim(Range("H2").Value)

''Sets IP to equal the first value in E for comparison
IP = Trim(Range("E2").Value)

''Start loop to create groups
For i = 1 To counter

    ''Creates lookup variable for each value in pluginID
    pluginid_combined = Trim(Range("H1").Offset(i, 0).Value)
    
    ''Creates lookup variable for each value in IP
    ip_check = Trim(Range("E1").Offset(i, 0).Value)
    
    '' Place a 1 in Column I for unique Plugin IDs and IPs
    If pluginid_combined = pluginid Then
    
        If ip_check = IP Then
            
        Else
            IP = Trim(Range("E1").Offset(i, 0).Value)
            incrementer = 1 + incrementer
            Range("I1").Offset(i, 0).Value = incrementer
        End If
    

    Else
        pluginid = Trim(Range("H1").Offset(i, 0).Value)
        incrementer = 1 + incrementer
        Range("I1").Offset(i, 0).Value = incrementer
        
        If ip_check = IP Then
            
        Else
            IP = Trim(Range("E1").Offset(i, 0).Value)
            
        End If

    End If
    
Next


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Creates CVE and IP Address pull out''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Starts from bottom to top comparing the number values to roll up CVE
i_low = Trim(Range("I2").Value)
i = 0

Do While counter <> i

    If Range("I1").Offset(counter, 0).Value <> "" Then
        
        Range("I1").Offset(counter, 1).Value = Range("E1").Offset(counter, 0).Value

        If cve_combined = "" Then
            cve_combined = Range("F1").Offset(counter, 0).Value
        Else
            cve_combined = Range("F1").Offset(counter, 0).Value & ", " & cve_combined
        End If
        
        Range("I1").Offset(counter, 2).Value = cve_combined
        cve_combined = ""
        
    Else
        
        If cve_combined = "" Then
            cve_combined = Range("F1").Offset(counter, 0).Value
        Else
            cve_combined = Range("F1").Offset(counter, 0).Value & ", " & cve_combined
        End If
        
    End If

counter = counter - 1

Loop

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Clear out CVE entries that are no longer necessary ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Deletes all entries where I has no number
Worksheets(output_sheet).Activate
Range("I1").Value = "Delete Column"
'On Error Resume Next
Columns("I").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Input Device Names for each IP Address'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Worksheets("IP List").Activate
ip_count = Range("A1", Range("A64000").End(xlUp)).Count - 1

ReDim Preserve ip_array(ip_count)
ReDim Preserve hostname_array(ip_count)

For i = 1 To ip_count

ip_array(i) = Range("A1").Offset(i, 0).Value
hostname_array(i) = Range("B1").Offset(i, 0).Value

Next

Worksheets(output_sheet).Activate
counter = Range("A1", Range("A64000").End(xlUp)).Count
For i = 1 To counter

ip_address = Range("e1").Offset(i, 0).Value
        ip_holder = 0
        For index = 1 To ip_count
            If ip_array(index) = ip_address Then
                ip_holder = index
            End If
        Next
            If ip_holder <> 0 Then
                Range("d1").Offset(i, 0).Value = hostname_array(ip_holder)
                
            Else
                Range("d1").Offset(i, 0).Value = ip_address & "?"
            End If
        
Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Delete Column I''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Columns("I:I").Delete

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Combines IP Addresses by Plugin ID''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
counter = Range("A1", Range("A64000").End(xlUp)).Count - 1
plugin_high = Trim(Range("H1").Offset(counter, 0).Value)


i = -1
Do While counter <> i

    If Range("H1").Offset(counter, 0).Value <> plugin_high Then

        If ip_combined = "" Then
            Range("K1").Offset(counter + 1, 0).Value = Range("E1").Offset(counter + 1, 0).Value
        Else
            Range("K1").Offset(counter + 1, 0).Value = ip_combined
            Range("L1").Offset(counter + 1, 0).Value = hostname_combined
            ip_combined = ""
            hostname_combined = ""
        End If
        ip_combined = Range("E1").Offset(counter, 0).Value
        hostname_combined = Range("D1").Offset(counter, 0).Value
        
        If counter > 0 Then
            plugin_high = Range("H1").Offset(counter, 0).Value
        End If

    Else

        If ip_combined = "" Then
            ip_combined = Range("E1").Offset(counter, 0).Value
            hostname_combined = Range("D1").Offset(counter, 0).Value
            
        Else
            ip_combined = Range("E1").Offset(counter, 0).Value & ", " & ip_combined
            hostname_combined = Range("D1").Offset(counter, 0).Value & ", " & hostname_combined
        End If
    End If


counter = counter - 1
Loop


'''Clear out CVE entries that are no longer necessary
''Deletes all entries where I has no number
Worksheets(output_sheet).Activate
Range("K1").Value = "Delete Column"
Columns("K").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

'''Copy over Consolidated IPs and CVEs
Columns("K:K").Copy
Range("E1").Select
ActiveSheet.Paste
Range("E1").Value = "IP Address"

Columns("J:J").Copy
Range("F1").Select
ActiveSheet.Paste
Range("F1").Value = "Reference"

Columns("L:L").Copy
Range("D1").Select
ActiveSheet.Paste
Range("D1").Value = "Device Name"


Columns("I:L").Delete


Application.ScreenUpdating = True



End Sub

Sub scc_parse()

'''Variables in entire document
Dim path, ThisWb, filename, output_sheet, primary_sheet, newsht_name() As String
Dim counter As Long
Dim wbDest, Wkb As Workbook
Dim shtDest, ws As Worksheet
Dim CopyRng, Dest As Range


'''Turns off showing data flying all over the screen
Application.ScreenUpdating = False

'''Used for clarity
ThisWb = ActiveWorkbook.Name

'''Where user inputs files for parsing
path = InputBox("Please enter the full file path to the SCC CSV's")

'''To explicitly define path
'path = "C:\Users\thomas.blauvelt\Desktop\Blauvelt\OF_working\Scans\for_rmp\SCC"

'''Sets the variable shtDest which looks at the workbook as an array, using the 2nd sheet (0 is the first)
Set shtDest = ActiveWorkbook.Sheets(1)
filename = Dir(path & "\*.csv", vbNormal)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''File Loops'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Grabs data from each file within the directory and inputs the device name into each based off the filename
''*****NOTE*******
''If you use "." in the filename this will error
If Len(filename) = 0 Then Exit Sub
Do Until filename = vbNullString
    If Not filename = ThisWb Then
        Set Wkb = Workbooks.Open(filename:=path & "\" & filename)
        counter = Range("A1", Range("A64000").End(xlUp)).Count
        newsht_name = Split(filename, ".")
        Range("G1").Value = "Hostname"
        Range("G2", "G" & counter).Value = newsht_name(0)
        Set CopyRng = Wkb.Sheets(1).Range(Cells(2, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column))
        Set Dest = shtDest.Range("A" & shtDest.Cells(Rows.Count, 1).End(xlUp).Row + 1)
        CopyRng.Copy
        Dest.PasteSpecial xlPasteFormats
        Dest.PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        Wkb.Close False
    End If
    
    filename = Dir()
Loop

'''Sets all the headers on first sheet before copy
Range("A1").Value = "Vulnerability"
Range("B1").Value = "Finding Title"
Range("C1").Value = "Finding Description"
Range("D1").Value = "Fix Recommendation"
Range("E1").Value = "Comments"
Range("F1").Value = "Reference"

'''Counter used to input device name for each row of data
counter = Range("A1", Range("A64000").End(xlUp)).Count
counter_trim = Trim(Str(counter))

'''Sort by Vulnerability(A) then by CCI(F) then Finding Description(C) in order to pull out information
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("A2:A" + counter_trim _
    ), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
    "high,medium,low", DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("E2:E" + counter_trim _
    ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("F2:F" + counter_trim _
    ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("C2:C" + counter_trim _
    ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Sheet1").Sort
    .SetRange Range("A2:G" + counter_trim)
    .Header = xlGuess
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Data Manipulation''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''If Reference and Finding Description is the same, combine rows
'''Place 1 in first count column because it will be the first unique value
''This is used to create a column for unique Plugin IDs and IPs
incrementer = 1
Range("I2").Value = incrementer

'''Counter used to input device name for each row of data
counter = Range("A1", Range("A64000").End(xlUp)).Count
counter_trim = Trim(Str(counter))

'''Sets pluginID to equal the first value in H for comparison
'pluginid = Trim(Range("H2").Value)
'
'''Sets IP to equal the first value in E for comparison
'IP = Trim(Range("E2").Value)


''Sets CCI to equal the first value in F for comparison as well as sets first Description in C
'cci_counter = Trim(Range("F2").Value)
description_counter = Trim(Range("B2").Value) + Trim(Range("C2").Value)

''Sets IP to equal the first value in E for comparison
hostname = Trim(Range("G2").Value)

''Start loop to create groups
For i = 1 To counter

    ''Creates lookup variable for each value in pluginID
     description_combined = Trim(Range("B2").Offset(i, 0).Value) + Trim(Range("C2").Offset(i, 0).Value) 'pluginid_combined
    
    ''Creates lookup variable for each value in IP
     hostname_check = Trim(Range("G2").Offset(i, 0).Value) 'ip_check
    
    '' Place a 1 in Column I for unique Plugin IDs and IPs
    If description_counter = description_combined Then 'cci_combined = cci_counter And
    
'        If hostname_check = hostname Then
'            hostname = Trim(Range("G2").Offset(i, 0).Value)
'            incrementer = 1 + incrementer
'            Range("I1").Offset(i, 0).Value = incrementer
'        Else
'
'        End If
'

    Else
        description_counter = Trim(Range("B2").Offset(i, 0).Value) + Trim(Range("C2").Offset(i, 0).Value) 'pluginid
        incrementer = 1 + incrementer
        Range("I2").Offset(i, 0).Value = incrementer
        
       ''Used for troubleshooting
        'Range("J2").Offset(i, 0).Value = description_counter
        
        If hostname_check = hostname Then
            
        Else
            hostname = Trim(Range("G1").Offset(i, 0).Value)
            
        End If
    End If

Next


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Creates Device Name pull out''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Starts from bottom to top comparing the number values to roll up CVE
i_low = Trim(Range("I2").Value)
i = 0

Do While counter <> i

    If Range("I1").Offset(counter, 0).Value <> "" Then
        
        'Range("I1").Offset(counter, 1).Value = Range("G1").Offset(counter, 0).Value

        If cve_combined = "" Then
            cve_combined = Range("G1").Offset(counter, 0).Value
        Else
            cve_combined = Range("G1").Offset(counter, 0).Value & ", " & cve_combined
        End If
        
        Range("I1").Offset(counter, 2).Value = cve_combined
        cve_combined = ""
        
    Else
        
        If cve_combined = "" Then
            cve_combined = Range("G1").Offset(counter, 0).Value
        Else
            cve_combined = Range("G1").Offset(counter, 0).Value & ", " & cve_combined
        End If
        
    End If

counter = counter - 1

Loop

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Clear out CVE entries that are no longer necessary ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Counter used to input device name for each row of data
counter = Range("A1", Range("A64000").End(xlUp)).Count
counter_trim = Trim(Str(counter))
stat_counter = Range("A1", Range("A64000").End(xlUp)).Count

''Deletes all entries where I has no number
Range("I1").Value = "Delete Column"
Columns("I").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Columns("I").Delete
Range("J2:J" + counter_trim).Copy
Range("G2").Select
ActiveSheet.Paste
Columns("J").Delete

'''Creating holder for statistics
counter = Range("A1", Range("A64000").End(xlUp)).Count
counter_trim = Trim(Str(counter))
unique_vulnerabilities = Range("A1", Range("A64000").End(xlUp)).Count
notafinding = 0
notapplicable = 0

i = 1
Do While i < counter

    If Range("E1").Offset(counter, 0) = "NotAFinding" Then
        notafinding = notafinding + 1
    End If
    
    If Range("E1").Offset(counter, 0) = "Not_Applicable" Then
        notapplicable = notapplicable + 1
    End If
    

counter = counter - 1

Loop


'''Counter
counter = Range("A1", Range("A64000").End(xlUp)).Count
counter_trim = Trim(Str(counter))

''Deletes if NotAFinding in comments
i = 1
Do While i < counter

    If Range("E1").Offset(counter, 0) <> "Open" Then
        Range("E1").Offset(counter, 0).EntireRow.Delete
    End If

counter = counter - 1

Loop

total_open = Range("A1", Range("A64000").End(xlUp)).Count


'''Variables used below for column moves added for clarity
output_sheet = "Working"
primary_sheet = ActiveSheet.Name

'''Creates a new worksheet to move all the data to
''*****NOTE*******
''The data is moved in order to change column headings and have data in an appropriate order making it easier to copy and paste
Sheets.Add After:=Sheets(Sheets.Count)

'''Sets the name of the new sheet
ActiveSheet.Name = output_sheet

'''Column A Move
Worksheets(primary_sheet).Activate
Columns("B:B").Copy
Worksheets(output_sheet).Activate
Range("A1").Select
ActiveSheet.Paste

'''Column B Move
Worksheets(primary_sheet).Activate
Columns("C:C").Copy
Worksheets(output_sheet).Activate
Range("B1").Select
ActiveSheet.Paste

'''Column C Move
Worksheets(primary_sheet).Activate
Columns("A:A").Copy
Worksheets(output_sheet).Activate
Range("C1").Select
ActiveSheet.Paste

'''Column D Move
Worksheets(primary_sheet).Activate
Columns("G:G").Copy
Worksheets(output_sheet).Activate
Range("D1").Select
ActiveSheet.Paste
Range("D1").Value = "Device Name"

'''Column E Move
Worksheets(output_sheet).Activate
Range("E1").Value = "IP Address"

'''Column F Move
Worksheets(primary_sheet).Activate
Columns("F:F").Copy
Worksheets(output_sheet).Activate
Range("F1").Select
ActiveSheet.Paste

'''Column G Move
Worksheets(primary_sheet).Activate
Columns("D:D").Copy
Worksheets(output_sheet).Activate
Range("G1").Select
ActiveSheet.Paste

'''Column H Move
Worksheets(primary_sheet).Activate
Columns("E:E").Copy
Worksheets(output_sheet).Activate
Range("H1").Select
ActiveSheet.Paste

'''Column I (Informational/Statistics)
Worksheets(output_sheet).Activate
Range("I1").Value = "Total Vulnerabilities Evaluated"
Range("I2").Value = Str(stat_counter)
Range("I4").Value = "Total Unique Vulnerabilities"
Range("I5").Value = Str(unique_vulnerabilities)
Range("I7").Value = "Total Not a Finding"
Range("I8").Value = Str(notafinding)
Range("I10").Value = "Total Not Applicable"
Range("I11").Value = Str(notapplicable)
Range("I13").Value = "Total Open"
Range("I14").Value = Str(total_open)


'''Sets the widths of all columns into an easier to read format
Columns("A:A").ColumnWidth = 15.86
Columns("B:B").ColumnWidth = 21
Columns("C:C").ColumnWidth = 14.71
Columns("D:D").ColumnWidth = 33.57
Columns("E:E").ColumnWidth = 10.71
Columns("F:F").ColumnWidth = 12.14
Columns("G:G").ColumnWidth = 20.43
Columns("H:H").ColumnWidth = 13.86


'''Counter used to input device name for each row of data
counter = Range("A1", Range("A64000").End(xlUp)).Count
counter_trim = Trim(Str(counter))

'''Creates total range for each column for the loop to use
column_all = "A1:H" + counter_trim

'''Move all the data down and add a title
Worksheets(output_sheet).Activate
Worksheets(output_sheet).Range(column_all).Copy
Range("A2").Select
ActiveSheet.Paste
Range("A1").Value = "SCC Findings"
Range("B1:H1").Value = ""

'''Format Title
Range("A1:H1").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
Selection.Merge
Selection.Font.Bold = True
Selection.Font.Size = 12
Selection.Font.Size = 14


'''Turns the screen update back on in order to see all the magic that happened
Application.ScreenUpdating = True

End Sub







